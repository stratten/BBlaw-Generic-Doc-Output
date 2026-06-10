import type { CapabilityManifest } from '../types/capability';

export const DOC_ARTIFACTS_CAPABILITY_ID = 'doc-artifacts';

export const capabilityManifest: CapabilityManifest = {
  id: DOC_ARTIFACTS_CAPABILITY_ID,
  name: 'Document Artifacts',
  description: 'Generate downloadable Word and PDF artifacts from arbitrary markdown or structured content.',
  icon: 'file-text',
  version: '1.0.0',
  inputSchema: {
    type: 'structured',
    accepts: ['application/json'],
    streaming: false,
    requiredFields: [
      {
        key: 'title',
        type: 'string',
        label: 'Title',
        description: 'Document title used for headings and generated filenames.',
      },
      {
        key: 'formats',
        type: 'string[]',
        label: 'Output formats',
        description: 'One or more output formats: docx, pdf.',
      },
    ],
    optionalFields: [
      {
        key: 'markdown',
        type: 'string',
        label: 'Markdown',
        description: 'Markdown content to render into the requested artifacts.',
      },
      {
        key: 'blocks',
        type: 'ContentBlock[]',
        label: 'Structured content blocks',
        description: 'Normalized blocks for callers that do not want markdown parsing.',
      },
      {
        key: 'style',
        type: 'string',
        label: 'Style',
        description: 'One of plain, memo, letter, or report.',
      },
      {
        key: 'sourceNotes',
        type: 'string',
        label: 'Source notes',
        description: 'Traceability notes shown in metadata but not rendered into the document body.',
      },
    ],
  },
  endpoints: {
    create: '/artifacts',
    metadata: '/artifacts/:artifactId',
    download: '/artifacts/:artifactId/files/:fileId',
  },
  surfaces: [
    {
      id: 'user',
      label: 'Document Artifacts',
      description: 'Create and download generic work-product artifacts.',
    },
    {
      id: 'system',
      label: 'Artifact API',
      description: 'Gateway and chat tool access to generated document artifacts.',
    },
  ],
  routes: [
    {
      surface: 'system',
      method: 'POST',
      path: '/artifacts',
      label: 'Create artifact',
      description: 'Generate one or more files from markdown or content blocks.',
    },
    {
      surface: 'system',
      method: 'GET',
      path: '/artifacts/:artifactId',
      label: 'Get artifact metadata',
    },
    {
      surface: 'system',
      method: 'GET',
      path: '/artifacts/:artifactId/files/:fileId',
      label: 'Download artifact file',
    },
  ],
  instructions: 'Use this capability only when the user explicitly asks to create or export a Word/PDF artifact.',
  formattingGuidelines: [
    'Treat markdown as markup for document structure, not as trusted HTML.',
    'Keep formal package generation in the separate document-generation service.',
  ],
  models: ['service-renderer'],
  defaultModel: 'service-renderer',
  estimatedDuration: '10-30 seconds',
};
