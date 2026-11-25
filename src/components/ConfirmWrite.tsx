import { t } from '../i18n';
import { uiMessage } from '../i18n-errors';
import { useEffect, useRef } from 'react';
import { AlertCircle, ArrowRight, LoaderCircle } from 'lucide-react';
import type { OutputKind, WritePlan } from '../core/types';
import { outlineToText } from '../core/output';
import { parseReport, parseTable, reportToText, tableToText } from '../core/structured-output';
import { WordTextPreview } from './WordTextPreview';

interface Props {
  plan: WritePlan;
  output: string;
  outputKind?: OutputKind;
  wordText?: boolean;
  returnFocus?: HTMLElement | null;
  busy: boolean;
  error: string;
  onCancel: () => void;
  onConfirm: () => void;
}
export function ConfirmWrite({
  plan,
  output,
  outputKind,
  wordText,
  returnFocus,
  busy,
  error,
  onCancel,
  onConfirm,
}: Props) {
  const dialog = useRef<HTMLDialogElement>(null);
  const cancel = useRef<HTMLButtonElement>(null);
  useEffect(() => {
    const previous = returnFocus || (document.activeElement as HTMLElement | null);
    const element = dialog.current;
    element?.showModal();
    cancel.current?.focus();
    return () => {
      element?.close();
      requestAnimationFrame(() => {
        if (!element?.open && previous?.isConnected) previous.focus();
      });
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
      <h2 id="confirm-title">{t('确认写入', 'Confirm Write')}</h2>
      <p className="target-label">{uiMessage(plan.label)}</p>
      <div className="write-comparison">
        <div>
          <h3 id="confirm-before-heading">{t('目标当前内容', 'Current content')}</h3>
          <pre tabIndex={0} role="region" aria-labelledby="confirm-before-heading">
            {plan.before || t('（空）', '(Empty)')}
          </pre>
        </div>
        <ArrowRight size={18} />
        <div>
          <h3 id="confirm-after-heading">{t('即将写入的结果', 'Content to write')}</h3>
          {wordText ? (
            <WordTextPreview
              text={output}
              tabIndex={0}
              role="region"
              aria-labelledby="confirm-after-heading"
            />
          ) : (
            <pre tabIndex={0} role="region" aria-labelledby="confirm-after-heading">
              {outputKind === 'outline'
                ? outlineToText(output)
                : outputKind === 'document'
                  ? reportToText(parseReport(output))
                  : outputKind === 'table'
                    ? tableToText(parseTable(output))
                    : output}
            </pre>
          )}
        </div>
      </div>
      {plan.warning && <p className="write-warning">{uiMessage(plan.warning)}</p>}
      <p className="field-help">
        {outputKind === 'table'
          ? t(
              '请检查新工作表名称和内容。需要恢复时，请核对表名后删除本次新建的工作表。',
              "Check the new worksheet's name and contents. To reverse this change, verify the sheet name and delete the new worksheet.",
            )
          : t(
              '请检查目标和内容。写入后可使用 Office 的撤销功能恢复；添加幻灯片可能需要多次撤销。',
              'Check the destination and content. You can undo the change in Office. Adding slides may take more than one Undo.',
            )}
      </p>
      {error && (
        <div className="notice error" role="alert">
          <AlertCircle size={17} />
          {uiMessage(error)}
        </div>
      )}
      <div className="button-row">
        <button ref={cancel} className="secondary" onClick={onCancel} disabled={busy}>
          {t('返回检查', 'Back to Review')}
        </button>
        <button className="primary" onClick={onConfirm} disabled={busy || Boolean(error)}>
          {busy && <LoaderCircle size={16} className="spin" />}
          {busy ? t('正在写入…', 'Writing…') : t('确认写入', 'Confirm Write')}
        </button>
      </div>
    </dialog>
  );
}
