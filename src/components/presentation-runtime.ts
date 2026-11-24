import type { ChatMessage } from '../core/types';
import type { MaterialContext, PresentationAccess } from '../presentation/types';
export interface PresentationRuntime {
  context: MaterialContext;
  access?: PresentationAccess;
  busy: boolean;
  demo: boolean;
  modelStamp: string;
  run: (label: string, work: (signal: AbortSignal) => Promise<void>) => Promise<void>;
  request: (messages: ChatMessage[], signal: AbortSignal, demoResult: string) => Promise<string>;
  notify: (message: string) => void;
}
