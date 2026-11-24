import { useEffect, useRef, useState } from 'react';
import { LoaderCircle, Square, X } from 'lucide-react';
import type { ChatMessage, HostBridge, RuntimeModelProfile } from '../core/types';
import type { MaterialContext } from '../presentation/types';
import { complete } from '../core/model-client';
import { officeError } from '../office/bridge';
import { MaterialsPicker } from './MaterialsPicker';
import { DeckGenerator } from './DeckGenerator';
import { PresentationEdit } from './PresentationEdit';
import { DataChart } from './DataChart';
import type { PresentationRuntime } from './presentation-runtime';

export type PresentationMode = 'generate' | 'edit' | 'chart' | 'text';
export function PresentationWorkspace({
  bridge,
  model,
  demoModel,
  sharedUnavailable,
  mode,
  onBusyChange,
}: {
  bridge: HostBridge;
  model?: RuntimeModelProfile;
  demoModel: boolean;
  sharedUnavailable: boolean;
  mode: PresentationMode;
  onBusyChange: (busy: boolean) => void;
}) {
  const [context, setContext] = useState<MaterialContext>({ materials: [], includeImages: false }),
    [busy, setBusy] = useState(''),
    [notice, setNotice] = useState<{ text: string; error?: boolean }>();
  const operation = useRef<AbortController | null>(null),
    alive = useRef(true);
  useEffect(() => {
    alive.current = true;
    return () => {
      alive.current = false;
      operation.current?.abort();
    };
  }, []);
  useEffect(() => {
    onBusyChange(Boolean(busy));
    return () => onBusyChange(false);
  }, [busy, onBusyChange]);
  useEffect(() => setNotice(undefined), [mode]);
  async function run(label: string, work: (signal: AbortSignal) => Promise<void>) {
    if (operation.current) return;
    const controller = new AbortController();
    operation.current = controller;
    setBusy(label);
    setNotice(undefined);
    try {
      await work(controller.signal);
      if (alive.current && controller.signal.aborted)
        setNotice({ text: '已停止，之前的资料与预览仍保留。' });
    } catch (e) {
      if (alive.current)
        setNotice({
          text: controller.signal.aborted ? '已停止，之前的资料与预览仍保留。' : officeError(e),
          error: !controller.signal.aborted,
        });
    } finally {
      operation.current = null;
      if (alive.current) setBusy('');
    }
  }
  async function request(messages: ChatMessage[], signal: AbortSignal, demoResult: string) {
    if (demoModel) return demoResult;
    if (!model) throw new Error('请先在模型设置中连接模型，再生成 PPT。');
    if (sharedUnavailable) throw new Error('本机共享模型服务尚未连接，请检查模型设置。');
    return complete(model, messages, signal);
  }
  const runtime: PresentationRuntime = {
    context,
    access: bridge.presentation,
    busy: Boolean(busy),
    demo: demoModel,
    modelStamp: `${model?.id || '__demo__'}:${model?.model || ''}`,
    run,
    request,
    notify: (text) => {
      if (alive.current) setNotice({ text });
    },
  };
  const cancellable = /生成整套|生成已有|推荐图表|读取本地|读取现有/.test(busy);
  return (
    <div
      className="presentation-workspace"
      hidden={mode === 'text'}
      onChangeCapture={() => {
        if (!busy) setNotice(undefined);
      }}
    >
      <MaterialsPicker context={context} onChange={setContext} busy={Boolean(busy)} run={run} />
      {busy && (
        <div className="ppt-progress" role="status">
          <LoaderCircle size={18} className="spin" />
          <span>{busy}…</span>
          {cancellable && (
            <button className="text-button" onClick={() => operation.current?.abort()}>
              <Square size={13} />
              停止
            </button>
          )}
        </div>
      )}
      {notice && (
        <div
          className={`notice ${notice.error ? 'error' : 'success'}`}
          role={notice.error ? 'alert' : 'status'}
        >
          <span>{notice.text}</span>
          <button
            className="icon-button"
            aria-label="关闭 PPT 提示"
            onClick={() => setNotice(undefined)}
          >
            <X size={16} />
          </button>
        </div>
      )}
      <div hidden={mode !== 'generate'} aria-hidden={mode !== 'generate'}>
        <DeckGenerator {...runtime} />
      </div>
      <div hidden={mode !== 'edit'} aria-hidden={mode !== 'edit'}>
        <PresentationEdit {...runtime} />
      </div>
      <div hidden={mode !== 'chart'} aria-hidden={mode !== 'chart'}>
        <DataChart {...runtime} />
      </div>
    </div>
  );
}
