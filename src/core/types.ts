export type HostKind = 'word' | 'excel' | 'powerpoint';
export type OutputKind = 'text' | 'formula' | 'outline';
export type WriteKind = 'replace' | 'insert' | 'formula' | 'slides';

export interface ModelProfile {
  id: string;
  name: string;
  baseUrl: string;
  model: string;
  apiKey: string;
  timeoutSeconds: number;
  rememberKey: boolean;
}

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
  /** A snapshot owns its tracked Office range until it is disposed. */
  replace?: (text: string) => Promise<void>;
  dispose?: () => Promise<void>;
}

export interface WritePlan {
  label: string;
  before: string;
  commit: (text: string) => Promise<void>;
  dispose?: () => Promise<void>;
}

export interface HostBridge {
  host: HostKind;
  mode: 'office' | 'demo';
  capabilities: {
    read: boolean;
    replace: boolean;
    insert: boolean;
    formula: boolean;
    slides: boolean;
  };
  readSelection: () => Promise<SelectionSnapshot>;
  prepareWrite: (kind: WriteKind, source?: SelectionSnapshot) => Promise<WritePlan>;
  dispose?: () => void;
}

export interface SlideContent {
  title: string;
  bullets: string[];
}
export interface ChatMessage {
  role: 'system' | 'user';
  content: string;
}
