import { useRef } from 'react';
import { Send, Trash2 } from 'lucide-react';
import { MAX_EDITING_FOLLOW_UP_LENGTH, type EditingTurn } from '../core/editing-conversation';
import { t, useUiLanguage } from '../i18n';

interface Props {
  id: string;
  turns: readonly EditingTurn[];
  value: string;
  onChange: (value: string) => void;
  onSubmit: () => void;
  onReset: () => void;
  disabled?: boolean;
  unavailable?: boolean;
  stale?: boolean;
}

export function EditingConversation({
  id,
  turns,
  value,
  onChange,
  onSubmit,
  onReset,
  disabled = false,
  unavailable = false,
  stale = false,
}: Props) {
  useUiLanguage();
  const composing = useRef(false);
  const input = useRef<HTMLTextAreaElement>(null);
  const tooLong = value.length > MAX_EDITING_FOLLOW_UP_LENGTH;
  const canSend = !disabled && !unavailable && !stale && Boolean(value.trim()) && !tooLong;

  return (
    <section className="editing-conversation" aria-labelledby={`${id}-heading`}>
      <div className="editing-conversation-heading">
        <h3 id={`${id}-heading`}>{t('编辑对话', 'Editing conversation')}</h3>
        <button
          type="button"
          className="text-button"
          disabled={disabled || !turns.length}
          onClick={() => {
            onReset();
            input.current?.focus();
          }}
        >
          <Trash2 size={14} aria-hidden="true" />
          {t('清空编辑对话', 'Clear editing conversation')}
        </button>
      </div>
      <p className="field-help">
        {t(
          'AI 会结合近期对话继续调整当前预览。确认写入后才修改文档，清空记录会保留当前预览。',
          'AI uses recent conversation to refine the current preview. Changes require your confirmation. Clearing the history keeps the preview.',
        )}
      </p>
      {turns.length > 0 && (
        <details className="editing-conversation-history">
          <summary>
            {t(
              `查看 ${turns.length} 轮编辑记录`,
              `View ${turns.length} editing ${turns.length === 1 ? 'round' : 'rounds'}`,
            )}
          </summary>
          <ol>
            {turns.map((turn, index) => (
              <li className="editing-conversation-turn" key={index}>
                <p>
                  <strong>{t(`第 ${index + 1} 轮`, `Round ${index + 1}`)}</strong>
                </p>
                <p>
                  {turn.instruction ||
                    t('按所选任务生成初稿', 'Generate the first draft for the selected task')}
                </p>
                <details>
                  <summary>{t('查看本轮结果', 'View this result')}</summary>
                  <pre className="editing-conversation-result" tabIndex={0}>
                    {turn.result}
                  </pre>
                </details>
              </li>
            ))}
          </ol>
        </details>
      )}
      {stale && (
        <p className="field-help" role="status">
          {t(
            '资料或任务已改变，请重新读取或生成预览后继续调整。',
            'The source or task changed. Read it again or generate a new preview before continuing.',
          )}
        </p>
      )}
      {unavailable && (
        <p className="field-help" role="status">
          {t(
            '请先连接可用模型，再继续编辑对话。',
            'Connect an available model to continue the editing conversation.',
          )}
        </p>
      )}
      <form
        onSubmit={(event) => {
          event.preventDefault();
          if (canSend) onSubmit();
        }}
      >
        <label htmlFor={`${id}-input`}>{t('继续调整', 'Refine the result')}</label>
        <textarea
          id={`${id}-input`}
          ref={input}
          rows={3}
          value={value}
          disabled={disabled}
          placeholder={t(
            '例如：保留主要结论，把第二段改得更简洁。',
            'Example: Keep the main conclusion and shorten the second paragraph.',
          )}
          aria-invalid={tooLong}
          aria-describedby={`${id}-help${tooLong ? ` ${id}-error` : ''}`}
          onChange={(event) => onChange(event.target.value)}
          onCompositionStart={() => {
            composing.current = true;
          }}
          onCompositionEnd={() => {
            composing.current = false;
          }}
          onKeyDown={(event) => {
            if (
              event.key === 'Enter' &&
              !event.shiftKey &&
              !event.nativeEvent.isComposing &&
              !composing.current &&
              event.keyCode !== 229
            ) {
              event.preventDefault();
              if (canSend) onSubmit();
            }
          }}
        />
        {tooLong && (
          <p className="field-error" role="alert" id={`${id}-error`}>
            {t(
              '修改要求超过 4,000 字符。内容已保留，请缩短后发送。',
              'Your changes exceed 4,000 characters. Your text is intact; shorten it before sending.',
            )}
          </p>
        )}
        <div className="editing-conversation-actions">
          <span className="field-help" id={`${id}-help`}>
            {t('Enter 发送 · Shift + Enter 换行', 'Enter to send · Shift + Enter for a new line')}
          </span>
          <button type="submit" className="secondary" disabled={!canSend}>
            <Send size={15} aria-hidden="true" />
            {t('发送修改要求', 'Send changes')}
          </button>
        </div>
      </form>
    </section>
  );
}
