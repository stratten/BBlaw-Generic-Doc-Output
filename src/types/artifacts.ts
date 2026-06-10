export type ArtifactFormat = 'docx' | 'pdf';
export type ArtifactStyle = 'plain' | 'memo' | 'letter' | 'report';

export interface ContentBlock {
  type: 'heading' | 'paragraph' | 'list' | 'table' | 'pageBreak';
  text?: string;
  level?: 1 | 2 | 3;
  ordered?: boolean;
  items?: string[];
  rows?: string[][];
}

export interface NormalizedDocument {
  title: string;
  style: ArtifactStyle;
  blocks: ContentBlock[];
  sourceNotes?: string;
}

export interface CreateArtifactRequest {
  title: string;
  markdown?: string;
  blocks?: ContentBlock[];
  formats: ArtifactFormat[];
  style?: ArtifactStyle;
  sourceNotes?: string;
  conversationId?: string;
  messageId?: string;
}

export interface PortalIdentity {
  id?: string;
  email?: string;
  name?: string;
}

export interface ArtifactFileMetadata {
  fileId: string;
  format: ArtifactFormat;
  filename: string;
  contentType: string;
  sizeBytes: number;
  downloadUrl: string;
  storagePath: string;
}

export interface ArtifactMetadata {
  artifactId: string;
  title: string;
  style: ArtifactStyle;
  status: 'ready';
  files: ArtifactFileMetadata[];
  sourceNotes?: string;
  conversationId?: string;
  messageId?: string;
  createdAt: string;
  expiresAt?: string;
  creator: PortalIdentity;
}

export interface StoredArtifactFile {
  metadata: ArtifactFileMetadata;
  buffer: Buffer;
}
