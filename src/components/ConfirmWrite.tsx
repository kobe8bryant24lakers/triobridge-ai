import { useEffect, useRef } from 'react';
import { AlertCircle, ArrowRight, LoaderCircle } from 'lucide-react';
import type { WritePlan } from '../core/types';

interface Props {
  plan: WritePlan;
  output: string;
  busy: boolean;
  error: string;
  onCancel: () => void;
  onConfirm: () => void;
}
export function ConfirmWrite({ plan, output, busy, error, onCancel, onConfirm }: Props) {
  const dialog = useRef<HTMLDialogElement>(null);
  const cancel = useRef<HTMLButtonElement>(null);
  useEffect(() => {
    const previous = document.activeElement as HTMLElement | null;
    dialog.current?.showModal();
    cancel.current?.focus();
    return () => {
      dialog.current?.close();
      previous?.focus();
    };
  }, []);
  return (
    <dialog
      ref={dialog}
      className="confirm-dialog"
      aria-labelledby="confirm-title"
      onKeyDown={(event) => {
        if (event.key !== 'Tab') return;
        // Keep a consistent cycle in Office WebViews, including Safari's default
        // keyboard policy, which otherwise skips buttons and visits browser chrome.
        const targets = Array.from(
          event.currentTarget.querySelectorAll<HTMLElement>(
            'button:not(:disabled), [tabindex="0"]',
          ),
        );
        if (!targets.length) return;
        event.preventDefault();
        const index = targets.indexOf(document.activeElement as HTMLElement);
        const next = event.shiftKey
          ? index <= 0
            ? targets.length - 1
            : index - 1
          : (index + 1) % targets.length;
        targets[next].focus();
      }}
      onCancel={(e) => {
        e.preventDefault();
        if (!busy) onCancel();
      }}
    >
      <h2 id="confirm-title">确认写入</h2>
      <p className="target-label">{plan.label}</p>
      <div className="write-comparison">
        <div>
          <h3 id="confirm-before-heading">目标当前内容</h3>
          <pre tabIndex={0} role="region" aria-labelledby="confirm-before-heading">
            {plan.before || '（空）'}
          </pre>
        </div>
        <ArrowRight size={18} />
        <div>
          <h3 id="confirm-after-heading">即将写入的结果</h3>
          <pre tabIndex={0} role="region" aria-labelledby="confirm-after-heading">
            {output}
          </pre>
        </div>
      </div>
      <p className="field-help">
        请检查目标和内容。写入后可使用 Office 的撤销功能恢复；添加幻灯片可能需要多次撤销。
      </p>
      {error && (
        <div className="notice error" role="alert">
          <AlertCircle size={17} />
          {error}
        </div>
      )}
      <div className="button-row">
        <button ref={cancel} className="secondary" onClick={onCancel} disabled={busy}>
          返回检查
        </button>
        <button className="primary" onClick={onConfirm} disabled={busy || Boolean(error)}>
          {busy && <LoaderCircle size={16} className="spin" />}
          {busy ? '正在写入…' : '确认写入'}
        </button>
      </div>
    </dialog>
  );
}
