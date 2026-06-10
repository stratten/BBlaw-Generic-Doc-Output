import { Router } from 'express';
import { z } from 'zod';
import { DOC_ARTIFACTS_CAPABILITY_ID } from '../config/capability';
import { createStoredArtifact, getStoredArtifact, getStoredArtifactFile, publicArtifactMetadata } from '../services/artifactStore';
import { logger } from '../services/logger';
import { renderDocument } from '../services/renderDocument';
import type { PortalIdentity } from '../types/artifacts';

export const artifactsRouter = Router();

const contentBlockSchema = z.object({
  type: z.enum(['heading', 'paragraph', 'list', 'table', 'pageBreak']),
  text: z.string().optional(),
  level: z.union([z.literal(1), z.literal(2), z.literal(3)]).optional(),
  ordered: z.boolean().optional(),
  items: z.array(z.string()).optional(),
  rows: z.array(z.array(z.string())).optional(),
});

const createArtifactSchema = z.object({
  title: z.string().trim().min(1).max(180),
  markdown: z.string().optional(),
  blocks: z.array(contentBlockSchema).optional(),
  formats: z.array(z.enum(['docx', 'pdf'])).min(1).max(2),
  style: z.enum(['plain', 'memo', 'letter', 'report']).optional(),
  sourceNotes: z.string().max(2000).optional(),
  conversationId: z.string().max(120).optional(),
  messageId: z.string().max(120).optional(),
}).refine((value) => Boolean(value.markdown?.trim() || value.blocks?.length), {
  message: 'Provide markdown or blocks.',
  path: ['markdown'],
});

function header(req: { headers: Record<string, string | string[] | undefined> }, name: string): string | undefined {
  const value = req.headers[name.toLowerCase()];
  if (Array.isArray(value)) return value[0];
  return value;
}

function portalIdentity(req: Parameters<typeof header>[0]): PortalIdentity {
  return {
    id: header(req, 'x-portal-user-id'),
    email: header(req, 'x-portal-user-email'),
    name: header(req, 'x-portal-user-name'),
  };
}

function hasCapability(req: Parameters<typeof header>[0]): boolean {
  if (header(req, 'x-gateway-broker') === '1') return true;
  const capabilities = (header(req, 'x-portal-capabilities') || '')
    .split(',')
    .map((capability) => capability.trim())
    .filter(Boolean);
  return capabilities.includes('*') || capabilities.includes(DOC_ARTIFACTS_CAPABILITY_ID);
}

function requireCapability(req: Parameters<typeof header>[0], res: { status: (code: number) => { json: (body: unknown) => void } }): boolean {
  if (hasCapability(req)) return true;
  res.status(403).json({ error: `Missing required capability: ${DOC_ARTIFACTS_CAPABILITY_ID}` });
  return false;
}

artifactsRouter.post('/artifacts', async (req, res, next) => {
  if (!requireCapability(req, res)) return;
  const parsed = createArtifactSchema.safeParse(req.body);
  if (!parsed.success) {
    res.status(400).json({ error: 'Invalid artifact request', issues: parsed.error.issues });
    return;
  }

  try {
    const { document, files } = await renderDocument(parsed.data);
    const metadata = await createStoredArtifact({
      title: document.title,
      style: document.style,
      files,
      creator: portalIdentity(req),
      sourceNotes: parsed.data.sourceNotes,
      conversationId: parsed.data.conversationId,
      messageId: parsed.data.messageId,
      baseDownloadPath: process.env.PUBLIC_BASE_PATH || '',
    });

    logger.info('generated document artifact', {
      artifactId: metadata.artifactId,
      title: metadata.title,
      formats: metadata.files.map((file) => file.format),
      fileSizes: metadata.files.map((file) => file.sizeBytes),
      userId: metadata.creator.id,
      userEmail: metadata.creator.email,
      conversationId: metadata.conversationId,
      messageId: metadata.messageId,
    });

    res.status(201).json(publicArtifactMetadata(metadata));
  } catch (error) {
    next(error);
  }
});

artifactsRouter.get('/artifacts/:artifactId', async (req, res, next) => {
  if (!requireCapability(req, res)) return;
  try {
    const metadata = await getStoredArtifact(req.params.artifactId);
    if (!metadata) {
      res.status(404).json({ error: 'Artifact not found' });
      return;
    }
    res.json(publicArtifactMetadata(metadata));
  } catch (error) {
    next(error);
  }
});

artifactsRouter.get('/artifacts/:artifactId/files/:fileId', async (req, res, next) => {
  if (!requireCapability(req, res)) return;
  try {
    const file = await getStoredArtifactFile(req.params.artifactId, req.params.fileId);
    if (!file) {
      res.status(404).json({ error: 'Artifact file not found' });
      return;
    }
    res.setHeader('content-type', file.metadata.contentType);
    res.setHeader('content-length', file.buffer.byteLength);
    res.setHeader('content-disposition', `attachment; filename="${file.metadata.filename.replace(/"/g, '')}"`);
    res.send(file.buffer);
  } catch (error) {
    next(error);
  }
});
