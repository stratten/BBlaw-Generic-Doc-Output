import { mkdir, readFile, stat, writeFile } from 'fs/promises';
import path from 'path';
import { nanoid } from 'nanoid';
import type { ArtifactFileMetadata, ArtifactMetadata, PortalIdentity, StoredArtifactFile } from '../types/artifacts';
import type { RenderedFile } from './renderDocument';

export interface CreateStoredArtifactInput {
  title: string;
  style: ArtifactMetadata['style'];
  files: RenderedFile[];
  creator: PortalIdentity;
  sourceNotes?: string;
  conversationId?: string;
  messageId?: string;
  baseDownloadPath?: string;
}

const DEFAULT_TTL_DAYS = Number(process.env.ARTIFACT_TTL_DAYS || 14);
const storageRoot = process.env.ARTIFACT_STORAGE_DIR || path.join(process.cwd(), '.artifacts');

function metadataPath(artifactId: string): string {
  return path.join(storageRoot, artifactId, 'metadata.json');
}

function artifactDir(artifactId: string): string {
  return path.join(storageRoot, artifactId);
}

function fileDownloadUrl(artifactId: string, fileId: string, baseDownloadPath = ''): string {
  return `${baseDownloadPath}/artifacts/${encodeURIComponent(artifactId)}/files/${encodeURIComponent(fileId)}`;
}

function expiresAt(createdAt: Date): string | undefined {
  if (!Number.isFinite(DEFAULT_TTL_DAYS) || DEFAULT_TTL_DAYS <= 0) return undefined;
  const expires = new Date(createdAt);
  expires.setDate(expires.getDate() + DEFAULT_TTL_DAYS);
  return expires.toISOString();
}

export async function createStoredArtifact(input: CreateStoredArtifactInput): Promise<ArtifactMetadata> {
  const artifactId = nanoid(16);
  const createdAt = new Date();
  const dir = artifactDir(artifactId);
  await mkdir(dir, { recursive: true });

  const files: ArtifactFileMetadata[] = [];
  for (const file of input.files) {
    const fileId = nanoid(12);
    const storagePath = path.join(dir, `${fileId}-${file.filename}`);
    await writeFile(storagePath, file.buffer);
    files.push({
      fileId,
      format: file.format,
      filename: file.filename,
      contentType: file.contentType,
      sizeBytes: file.buffer.byteLength,
      downloadUrl: fileDownloadUrl(artifactId, fileId, input.baseDownloadPath),
      storagePath,
    });
  }

  const metadata: ArtifactMetadata = {
    artifactId,
    title: input.title,
    style: input.style,
    status: 'ready',
    files,
    sourceNotes: input.sourceNotes,
    conversationId: input.conversationId,
    messageId: input.messageId,
    createdAt: createdAt.toISOString(),
    expiresAt: expiresAt(createdAt),
    creator: input.creator,
  };
  await writeFile(metadataPath(artifactId), JSON.stringify(metadata, null, 2));
  return metadata;
}

export async function getStoredArtifact(artifactId: string): Promise<ArtifactMetadata | undefined> {
  try {
    const body = await readFile(metadataPath(artifactId), 'utf8');
    return JSON.parse(body) as ArtifactMetadata;
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === 'ENOENT') return undefined;
    throw error;
  }
}

export async function getStoredArtifactFile(artifactId: string, fileId: string): Promise<StoredArtifactFile | undefined> {
  const metadata = await getStoredArtifact(artifactId);
  const file = metadata?.files.find((candidate) => candidate.fileId === fileId);
  if (!file) return undefined;
  await stat(file.storagePath);
  const buffer = await readFile(file.storagePath);
  return { metadata: file, buffer };
}

export function publicArtifactMetadata(metadata: ArtifactMetadata): Omit<ArtifactMetadata, 'files'> & { files: Omit<ArtifactFileMetadata, 'storagePath'>[] } {
  return {
    ...metadata,
    files: metadata.files.map(({ storagePath: _storagePath, ...file }) => file),
  };
}
