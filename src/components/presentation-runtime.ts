import type { ChatMessage, CompletionPurpose } from '../core/types';
import type { MaterialContext, PresentationAccess } from '../presentation/types';
export interface PresentationRuntime {
  context: MaterialContext;
  access?: PresentationAccess;
  busy: boolean;
  demo: boolean;
  modelStamp: string;
  run: (
    label: string,
    work: (signal: AbortSignal) => Promise<void>,
    cancellable?: boolean,
  ) => Promise<void>;
  request: (
    messages: ChatMessage[],
    signal: AbortSignal,
    demoResult: string,
    purpose?: CompletionPurpose,
  ) => Promise<string>;
  notify: (message: string) => void;
}

export function englishCount(count: number, noun: string, plural = noun + 's'): string {
  return `${count} ${count === 1 ? noun : plural}`;
}
