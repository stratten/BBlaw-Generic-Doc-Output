export interface InputSchemaField {
  key: string;
  type: string;
  label: string;
  placeholder?: string;
  description?: string;
  accept?: string[];
  maxCount?: number;
  options_endpoint?: string;
  schema_endpoint?: string;
}

export interface InputSchema {
  type: 'conversation' | 'structured';
  accepts?: string[];
  maxFiles?: number;
  maxFileSizeMB?: number;
  streaming?: boolean;
  requiredFields?: InputSchemaField[];
  optionalFields?: InputSchemaField[];
}

export type CapabilitySurfaceId = 'admin' | 'user' | 'system';

export interface CapabilitySurface {
  id: CapabilitySurfaceId;
  label: string;
  description?: string;
  entryPath?: string;
}

export interface CapabilityRoute {
  surface: CapabilitySurfaceId;
  method: string;
  path: string;
  label: string;
  description?: string;
}

export interface CapabilityManifest {
  id: string;
  name: string;
  description: string;
  icon?: string;
  version: string;
  inputSchema: InputSchema;
  endpoints?: Record<string, unknown>;
  surfaces?: CapabilitySurface[];
  routes?: CapabilityRoute[];
  instructions?: string;
  formattingGuidelines?: string[];
  models: string[];
  defaultModel: string;
  estimatedDuration?: string;
}
