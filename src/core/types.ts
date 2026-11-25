export type HostKind = 'word' | 'excel' | 'powerpoint';
export type OutputKind = 'text' | 'formula' | 'outline' | 'document' | 'table';
export type WriteKind = 'replace' | 'insert' | 'formula' | 'slides' | 'document' | 'table';

export type ModelApiFormat = 'openai' | 'anthropic';

export interface ModelProfile {
  id: string;
  name: string;
  baseUrl: string;
  /** Profiles created before format selection use OpenAI Chat Completions. */
  apiFormat?: ModelApiFormat;
  model: string;
  apiKey: string;
  timeoutSeconds: number;
  rememberKey: boolean;
}

export type ProfileDetails = Pick<
  ModelProfile,
  'id' | 'name' | 'baseUrl' | 'apiFormat' | 'model' | 'timeoutSeconds'
>;

/** Saved shared credentials never enter the browser's profile state. */
export interface SharedModelProfile extends ProfileDetails {
  kind: 'shared';
  hasKey: boolean;
  revision: number;
}

export type RuntimeModelProfile = ModelProfile | SharedModelProfile;

export interface TaskSpec {
  id: string;
  label: string;
  description: string;
  instruction: string;
  needsSource: boolean;
  needsInstruction?: boolean;
  output: OutputKind;
}

export interface SelectionSnapshot {
  text: string;
  label: string;
  warning?: string;
  /** A snapshot owns its tracked Office range until it is disposed. */
  replace?: (text: string) => Promise<void>;
  dispose?: () => Promise<void>;
}

export interface WritePlan {
  label: string;
  before: string;
  warning?: string;
  commit: (text: string) => Promise<{ warning?: string } | void>;
  dispose?: () => Promise<void>;
}

export interface HostBridge {
  presentation?: import('../presentation/types').PresentationAccess;
  host: HostKind;
  mode: 'office' | 'demo';
  capabilities: {
    read: boolean;
    replace: boolean;
    insert: boolean;
    formula: boolean;
    slides: boolean;
    document?: boolean;
    table?: boolean;
  };
  readSelection: () => Promise<SelectionSnapshot>;
  prepareWrite: (
    kind: WriteKind,
    source?: SelectionSnapshot,
    output?: string,
  ) => Promise<WritePlan>;
  dispose?: () => void;
}

export interface SlideContent {
  title: string;
  bullets: string[];
}
export interface ReportSection {
  heading: string;
  paragraphs: string[];
  bullets: string[];
  table?: { headers: string[]; rows: string[][] };
}
export interface ReportContent {
  title: string;
  sections: ReportSection[];
}
export type TableCell = string | number | boolean | null | { formula: string };
export type ColumnFormat = 'general' | 'text' | 'integer' | 'number' | 'percent';
export interface TableContent {
  name: string;
  headers: string[];
  rows: TableCell[][];
  formats: ColumnFormat[];
}
export interface ChatMessage {
  role: 'system' | 'user';
  content: string | ChatContentPart[];
}
export type ChatContentPart =
  | { type: 'text'; text: string }
  | { type: 'image_url'; image_url: { url: string; detail: 'low' | 'high' | 'auto' } };

export interface TextChatMessage {
  role: 'system' | 'user';
  content: string;
}
