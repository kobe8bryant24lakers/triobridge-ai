import { useEffect, useRef, useState } from 'react';
import {
  ArrowRight,
  LoaderCircle,
  Send,
  Settings2,
  ShieldCheck,
  Square,
  Trash2,
} from 'lucide-react';
import {
  MAX_CONVERSATION_INPUT_LENGTH,
  replyToConversation,
  type ConversationTurn,
} from '../core/conversation';
import { customFormatSuffix } from '../core/model-presets';
import type { ProfileState } from '../core/profile-store';
import type { HostKind } from '../core/types';
import { t, useUiLanguage } from '../i18n';
import { uiMessage } from '../i18n-errors';
import { officeError } from '../office/bridge';
import { WordTextPreview } from './WordTextPreview';

interface Props {
  host: HostKind;
  profiles: ProfileState;
  visible?: boolean;
  onChooseModel: (id: string) => void | Promise<void>;
  onSettings: () => void;
  onBusyChange: (busy: boolean) => void;
}

interface Exchange {
  id: number;
  question: string;
  answer?: string;
  modelLabel: string;
  status: 'pending' | 'complete' | 'error' | 'stopped';
  error?: string;
}

// Deliberately receives no document bridge: model replies can only enter this transcript.
export function Conversation({
  host,
  profiles,
  visible = true,
  onChooseModel,
  onSettings,
  onBusyChange,
}: Props) {
  useUiLanguage();
  const [modelId, setModelId] = useState(profiles.activeId);
  const [switchingModel, setSwitchingModel] = useState(false);
  const [modelError, setModelError] = useState('');
  const [draft, setDraft] = useState('');
  const [exchanges, setExchanges] = useState<Exchange[]>([]);
  const [generating, setGenerating] = useState(false);
  const request = useRef<AbortController | null>(null);
  const sequence = useRef(0);
  const mounted = useRef(false);
  const composing = useRef(false);
  const input = useRef<HTMLTextAreaElement>(null);
  const transcript = useRef<HTMLDivElement>(null);
  const model = profiles.profiles.find((profile) => profile.id === modelId);
  const unavailable =
    profiles.mode === 'checking' || (profiles.mode === 'shared' && profiles.connection !== 'ready');
  const active = generating || switchingModel;
  const tooLong = draft.length > MAX_CONVERSATION_INPUT_LENGTH;
  const canSend = Boolean(model && !unavailable && !active && draft.trim() && !tooLong);

  useEffect(() => setModelId(profiles.activeId), [profiles.activeId]);
  useEffect(() => {
    onBusyChange(active);
    return () => onBusyChange(false);
  }, [active, onBusyChange]);
  useEffect(() => {
    mounted.current = true;
    return () => {
      mounted.current = false;
      request.current?.abort();
      request.current = null;
    };
  }, []);
  useEffect(() => {
    if (visible && transcript.current)
      transcript.current.scrollTop = transcript.current.scrollHeight;
  }, [exchanges, visible]);

  async function chooseModel(id: string) {
    if (active) return;
    setSwitchingModel(true);
    setModelError('');
    try {
      await onChooseModel(id);
      if (mounted.current && profiles.mode !== 'shared') setModelId(id);
    } catch (error) {
      if (mounted.current) setModelError(officeError(error));
    } finally {
      if (mounted.current) setSwitchingModel(false);
    }
  }

  async function send(retry?: Exchange) {
    if (request.current || active || !model || unavailable || (!retry && !canSend)) return;
    const question = retry?.question ?? draft.trim();
    const id = retry?.id ?? ++sequence.current;
    const pending: Exchange = {
      id,
      question,
      modelLabel: `${model.name} · ${model.model}`,
      status: 'pending',
    };
    const history: ConversationTurn[] = exchanges
      .filter((item) => item.status === 'complete' && item.id < id)
      .flatMap((item) => [
        { role: 'user' as const, content: item.question },
        { role: 'assistant' as const, content: item.answer! },
      ]);
    history.push({ role: 'user', content: question });
    const controller = new AbortController();
    request.current = controller;
    setGenerating(true);
    setModelError('');
    setExchanges((items) =>
      retry ? items.map((item) => (item.id === id ? pending : item)) : [...items, pending],
    );
    if (!retry) setDraft('');
    try {
      const answer = await replyToConversation(model, history, host, controller.signal);
      if (request.current !== controller) return;
      setExchanges((items) =>
        items.map((item) => (item.id === id ? { ...item, answer, status: 'complete' } : item)),
      );
    } catch (error) {
      if (request.current !== controller) return;
      setExchanges((items) =>
        items.map((item) =>
          item.id === id ? { ...item, error: officeError(error), status: 'error' } : item,
        ),
      );
    } finally {
      if (request.current === controller) {
        request.current = null;
        setGenerating(false);
      }
    }
  }

  function stop() {
    const controller = request.current;
    request.current = null;
    controller?.abort();
    setGenerating(false);
    setExchanges((items) =>
      items.map((item) => (item.status === 'pending' ? { ...item, status: 'stopped' } : item)),
    );
    input.current?.focus();
  }

  const suggestions = [
    [
      t('插件怎么用', 'Using the add-in'),
      t(
        '这个插件怎么用？请介绍主要功能和操作步骤。',
        'How do I use this add-in? Explain its main features and steps.',
      ),
    ],
    [
      t('梳理 PPT 思路', 'Plan a presentation'),
      t(
        '帮我梳理一份 PPT 的思路，先问我主题、听众和希望达到的目标。',
        'Help me plan a presentation. Start by asking about the topic, audience, and goal.',
      ),
    ],
    [
      t('聊聊其他问题', 'Ask anything'),
      t(
        '怎样把一个复杂的问题拆解成可执行的步骤？',
        'How can I break a complex problem into actionable steps?',
      ),
    ],
  ];

  return (
    <section className="conversation" aria-label={t('对话工作区', 'Conversation workspace')}>
      <div className="page-heading">
        <h1>{t('对话模式', 'Conversation')}</h1>
        <p>
          {t(
            '了解插件用法、梳理 PPT 思路，或聊聊其他问题。',
            'Get help with the add-in, plan a presentation, or ask about anything else.',
          )}
        </p>
      </div>
      <div className="model-strip">
        <div className="model-select">
          <label htmlFor="conversation-model">{t('对话模型', 'Conversation model')}</label>
          <select
            id="conversation-model"
            value={modelId}
            disabled={active || unavailable}
            onChange={(event) => void chooseModel(event.target.value)}
          >
            <option value="" disabled>
              {t('请添加模型', 'Add a model')}
            </option>
            {profiles.profiles.map((profile) => (
              <option key={profile.id} value={profile.id}>
                {profile.name} · {profile.model}
                {customFormatSuffix(profile)}
              </option>
            ))}
          </select>
        </div>
        <button
          className="icon-button"
          onClick={onSettings}
          disabled={active}
          aria-label={t('打开对话模型设置', 'Open conversation model settings')}
          title={t('模型设置', 'Model Settings')}
        >
          <Settings2 size={19} aria-hidden="true" />
        </button>
      </div>
      <p className="conversation-boundary">
        <ShieldCheck size={16} aria-hidden="true" />
        {t(
          '只回答问题，不读取或修改当前文档。',
          'Answers only. This mode does not read or change your document.',
        )}
      </p>
      {unavailable ? (
        <p className="notice" role="status">
          {profiles.mode === 'checking'
            ? t('正在加载共享模型设置…', 'Loading shared model settings…')
            : t(
                '本机模型服务暂不可用，请在模型设置中检查连接。',
                'The local model service is unavailable. Check the connection in Model Settings.',
              )}
        </p>
      ) : (
        !model && (
          <div className="notice conversation-model-notice">
            <span>
              {t('连接自己的模型后即可开始对话。', 'Connect your model to start a conversation.')}
            </span>
            <button className="text-button" onClick={onSettings}>
              {t('配置模型', 'Set up a model')}
              <ArrowRight size={15} aria-hidden="true" />
            </button>
          </div>
        )
      )}
      {modelError && (
        <p className="notice error" role="alert">
          {uiMessage(modelError)}
        </p>
      )}
      <div className="conversation-toolbar">
        <h2>{t('本次对话', 'This conversation')}</h2>
        <button
          className="text-button"
          disabled={active || !exchanges.length}
          onClick={() => {
            setExchanges([]);
            input.current?.focus();
          }}
        >
          <Trash2 size={15} aria-hidden="true" />
          {t('清空对话', 'Clear conversation')}
        </button>
      </div>
      {exchanges.length ? (
        <div
          className="conversation-transcript"
          role="log"
          aria-label={t('对话记录', 'Conversation history')}
          aria-live="polite"
          aria-relevant="additions text"
          tabIndex={0}
          ref={transcript}
        >
          {exchanges.map((item, index) => (
            <div className="conversation-exchange" key={item.id}>
              <article className="conversation-question">
                <h3>{t('你', 'You')}</h3>
                <p>{item.question}</p>
              </article>
              <article className="conversation-answer">
                <h3>
                  {t('助手', 'Assistant')}
                  <span>{item.modelLabel}</span>
                </h3>
                {item.status === 'complete' ? (
                  <WordTextPreview text={item.answer!} />
                ) : item.status === 'pending' ? (
                  <p className="conversation-pending">
                    <LoaderCircle size={16} className="spin" aria-hidden="true" />
                    {t('正在回答…', 'Thinking…')}
                  </p>
                ) : (
                  <>
                    <p
                      className={item.status === 'error' ? 'field-error' : 'muted'}
                      role={item.status === 'error' ? 'alert' : undefined}
                    >
                      {item.status === 'error'
                        ? uiMessage(item.error!)
                        : t('已停止回答。', 'Response stopped.')}
                    </p>
                    {index === exchanges.length - 1 && (
                      <button
                        className="text-button"
                        disabled={active || !model || unavailable}
                        onClick={() => void send(item)}
                      >
                        {t('重试', 'Retry')}
                      </button>
                    )}
                  </>
                )}
              </article>
            </div>
          ))}
        </div>
      ) : (
        <div className="conversation-empty">
          <p>{t('从一个问题开始', 'Start with a question')}</p>
          <div
            className="conversation-suggestions"
            role="group"
            aria-label={t('提问示例', 'Suggested questions')}
          >
            {suggestions.map(([label, prompt]) => (
              <button
                className="text-button"
                key={label}
                onClick={() => {
                  setDraft(prompt);
                  input.current?.focus();
                }}
              >
                {label}
                <ArrowRight size={14} aria-hidden="true" />
              </button>
            ))}
          </div>
        </div>
      )}
      <form
        className="conversation-composer"
        onSubmit={(event) => {
          event.preventDefault();
          void send();
        }}
      >
        <label htmlFor="conversation-input">{t('你的问题', 'Your message')}</label>
        <textarea
          id="conversation-input"
          ref={input}
          rows={3}
          value={draft}
          placeholder={t(
            '输入问题，也可以粘贴需要讨论的内容…',
            'Ask a question or paste something to discuss…',
          )}
          aria-invalid={tooLong}
          aria-describedby={`conversation-input-help${tooLong ? ' conversation-input-error' : ''}`}
          onChange={(event) => setDraft(event.target.value)}
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
              void send();
            }
          }}
        />
        {tooLong && (
          <p id="conversation-input-error" className="field-error" role="alert">
            {t(
              '问题超过 4,000 字符。内容已保留，请缩短后发送。',
              'Your message exceeds 4,000 characters. Your text is intact; shorten it before sending.',
            )}
          </p>
        )}
        <div className="conversation-send-row">
          <span className="field-help" id="conversation-input-help">
            {t('Enter 发送 · Shift + Enter 换行', 'Enter to send · Shift + Enter for a new line')}
          </span>
          {generating ? (
            <button className="secondary" type="button" onClick={stop}>
              <Square size={15} aria-hidden="true" />
              {t('停止回答', 'Stop response')}
            </button>
          ) : (
            <button className="primary" type="submit" disabled={!canSend}>
              <Send size={16} aria-hidden="true" />
              {t('发送', 'Send')}
            </button>
          )}
        </div>
        <p className="field-help conversation-retention">
          {t(
            '发送时会将问题及近期对话发送给所选模型。记录仅保留在当前窗口，刷新后清空。',
            'Your message and recent conversation are sent to the selected model. History stays in this window and clears on reload.',
          )}
        </p>
      </form>
    </section>
  );
}
