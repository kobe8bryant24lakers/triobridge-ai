import { useEffect, useRef, useState } from 'react';
import {
  ArrowDownToLine,
  ArrowRight,
  Check,
  CheckCheck,
  Copy,
  FileCheck2,
  FileInput,
  LoaderCircle,
  Pencil,
  ScanText,
  Settings2,
  ShieldCheck,
  Sparkles,
  Square,
  X,
} from 'lucide-react';
import { complete } from '../core/model-client';
import { outlineToText, parseFormula, parseOutline } from '../core/output';
import {
  HOST_NAMES,
  MAX_INSTRUCTION_LENGTH,
  MAX_SOURCE_LENGTH,
  TASKS,
  buildMessages,
} from '../core/tasks';
import type { ProfileState } from '../core/profile-store';
import type { HostBridge, SelectionSnapshot, WriteKind, WritePlan } from '../core/types';
import { officeError } from '../office/bridge';
import { demoComplete } from '../office/demo';
import { ConfirmWrite } from './ConfirmWrite';

interface Props {
  bridge: HostBridge;
  profiles: ProfileState;
  onChooseModel: (id: string) => void;
  onSettings: () => void;
  onBusyChange: (busy: boolean) => void;
}

async function copyText(text: string) {
  if (navigator.clipboard) {
    await navigator.clipboard.writeText(text);
    return;
  }
  const input = document.createElement('textarea');
  input.value = text;
  input.style.position = 'fixed';
  input.style.opacity = '0';
  document.body.append(input);
  input.select();
  try {
    if (!document.execCommand('copy')) throw new Error('复制失败，请选中结果手动复制。');
  } finally {
    input.remove();
  }
}

export function Workbench({ bridge, profiles, onChooseModel, onSettings, onBusyChange }: Props) {
  const tasks = TASKS[bridge.host];
  const [taskId, setTaskId] = useState(tasks[0].id);
  const task = tasks.find((item) => item.id === taskId)!;
  const [modelId, setModelId] = useState(
    profiles.activeId || (bridge.mode === 'demo' ? '__demo__' : ''),
  );
  const model = profiles.profiles.find((p) => p.id === modelId);
  const isDemoModel = bridge.mode === 'demo' && modelId === '__demo__';
  const [source, setSource] = useState('');
  const snapshot = useRef<SelectionSnapshot | undefined>(undefined);
  const [sourceLabel, setSourceLabel] = useState('尚未读取选区');
  const [instruction, setInstruction] = useState('');
  const [language, setLanguage] = useState('英语');
  const [result, setResult] = useState('');
  const [resultModel, setResultModel] = useState('');
  const [editing, setEditing] = useState(false);
  const [busy, setBusy] = useState<'reading' | 'generating' | 'preparing' | 'writing' | ''>('');
  const [notice, setNotice] = useState<{ text: string; error?: boolean }>();
  const [written, setWritten] = useState(false);
  const [plan, setPlan] = useState<WritePlan>();
  const [writeValue, setWriteValue] = useState('');
  const [writeError, setWriteError] = useState('');
  const request = useRef<AbortController | null>(null);
  const activePlan = useRef<WritePlan | undefined>(undefined);
  const outputRef = useRef<HTMLElement>(null);
  const disposed = useRef(false);
  const active = Boolean(busy) || Boolean(plan);

  function chooseModel(id: string) {
    try {
      if (id !== '__demo__') onChooseModel(id);
      setModelId(id);
      invalidate();
    } catch (error) {
      setNotice({ text: officeError(error), error: true });
    }
  }

  useEffect(() => {
    onBusyChange(active);
    return () => onBusyChange(false);
  }, [active, onBusyChange]);
  useEffect(
    () => () => {
      disposed.current = true;
      request.current?.abort();
      void snapshot.current?.dispose?.().catch(() => {});
      void activePlan.current?.dispose?.().catch(() => {});
    },
    [],
  );

  function invalidate() {
    setResult('');
    setWritten(false);
    setEditing(false);
    setNotice(undefined);
  }
  function releaseSource() {
    void snapshot.current?.dispose?.().catch(() => {});
    snapshot.current = undefined;
  }
  function editSource(value: string) {
    releaseSource();
    setSource(value);
    setSourceLabel(value ? '手动输入 · 可发送给模型' : '尚未读取选区');
    invalidate();
  }
  async function read() {
    setBusy('reading');
    setNotice(undefined);
    try {
      const next = await bridge.readSelection();
      if (disposed.current) {
        void next.dispose?.().catch(() => {});
        return;
      }
      releaseSource();
      snapshot.current = next;
      setSource(next.text);
      setSourceLabel(next.label);
      invalidate();
    } catch (error) {
      setNotice({ text: officeError(error), error: true });
    } finally {
      setBusy('');
    }
  }
  async function generate() {
    setNotice(undefined);
    try {
      const messages = buildMessages(task, source, instruction, language);
      if (!model && !isDemoModel) throw new Error('请先添加并选择一个模型配置。');
      const controller = new AbortController();
      request.current = controller;
      setBusy('generating');
      setResult('');
      setWritten(false);
      setEditing(false);
      const text = isDemoModel
        ? await demoComplete(task, source, instruction, controller.signal)
        : await complete(model!, messages, controller.signal);
      if (controller.signal.aborted || disposed.current) return;
      setResult(text);
      setResultModel(isDemoModel ? '固定演示样例 · 未调用 AI' : `${model!.name} / ${model!.model}`);
      setTimeout(() => outputRef.current?.focus(), 0);
    } catch (error) {
      if (!disposed.current)
        setNotice({
          text: officeError(error),
          error: !request.current?.signal.aborted,
        });
    } finally {
      if (!disposed.current) setBusy('');
      request.current = null;
    }
  }
  async function prepare(kind: WriteKind) {
    setBusy('preparing');
    setNotice(undefined);
    try {
      const text = kind === 'formula' ? parseFormula(result) : result.trim();
      if (kind === 'slides') parseOutline(text);
      if (!text) throw new Error('结果为空，请先生成内容。');
      const next = await bridge.prepareWrite(kind, snapshot.current);
      setWriteValue(text);
      setWriteError('');
      activePlan.current = next;
      setPlan(next);
    } catch (error) {
      setNotice({ text: officeError(error), error: true });
    } finally {
      setBusy('');
    }
  }
  function dismissPlan() {
    void activePlan.current?.dispose?.().catch(() => {});
    activePlan.current = undefined;
    setPlan(undefined);
  }
  async function commit() {
    if (!plan) return;
    setBusy('writing');
    try {
      await plan.commit(writeValue);
      setWritten(true);
      dismissPlan();
      releaseSource();
      setSourceLabel('写入完成 · 再次处理前请重新读取');
      setNotice({
        text:
          bridge.mode === 'demo'
            ? '已写入模拟文档，可展开下方文档查看。真实 Office 文件未被修改。'
            : '已写入 Office。需要恢复时，请使用 Office 撤销。',
      });
    } catch (error) {
      setWriteError(
        `${officeError(error)}${task.output === 'outline' ? ' 如已出现部分新页面，请先检查并撤销，再重试。' : ''}`,
      );
    } finally {
      setBusy('');
    }
  }
  async function copy() {
    try {
      await copyText(task.output === 'outline' ? outlineToText(result) : result);
      setNotice({ text: '结果已复制。' });
    } catch (error) {
      setNotice({ text: officeError(error), error: true });
    }
  }
  function download() {
    const url = URL.createObjectURL(new Blob([result], { type: 'text/plain;charset=utf-8' }));
    const a = document.createElement('a');
    a.href = url;
    a.download = `triobridge-${task.id}.${task.output === 'outline' ? 'json' : 'txt'}`;
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }
  let slides;
  let formatError = '';
  if (result && task.output !== 'text') {
    try {
      if (task.output === 'outline') slides = parseOutline(result);
      else parseFormula(result);
    } catch (error) {
      formatError = officeError(error);
    }
  }
  const canReplace =
    bridge.capabilities.replace && Boolean(snapshot.current?.replace) && task.output === 'text';
  const canInsert = bridge.capabilities.insert && task.output === 'text';
  const canFormula = bridge.capabilities.formula && task.output === 'formula';
  const canSlides = bridge.capabilities.slides && task.output === 'outline';

  return (
    <>
      <div className="page-heading work-heading">
        <div>
          <h1>{HOST_NAMES[bridge.host]} 工作台</h1>
          <p>
            {bridge.host === 'word'
              ? '从你的文字出发，让下一稿更好。'
              : bridge.host === 'excel'
                ? '理解表格，让计算更有把握。'
                : '整理想法，把重点讲清楚。'}
          </p>
        </div>
        <span className="connection-badge">
          <span />
          {bridge.mode === 'demo' ? '浏览器演示' : '已连接 Office'}
        </span>
      </div>
      <div className="model-strip">
        <div className="model-select">
          <label htmlFor="active-model">当前模型</label>
          <select
            id="active-model"
            value={modelId}
            disabled={active}
            onChange={(e) => chooseModel(e.target.value)}
          >
            <option value="" disabled>
              请添加模型
            </option>
            {bridge.mode === 'demo' && <option value="__demo__">演示模型 · 无需 API Key</option>}
            {profiles.profiles.map((p) => (
              <option key={p.id} value={p.id}>
                {p.name} · {p.model}
              </option>
            ))}
          </select>
        </div>
        <button
          className="icon-button"
          aria-label="打开模型设置"
          title="模型设置"
          onClick={onSettings}
          disabled={active}
        >
          <Settings2 size={19} />
        </button>
      </div>
      {bridge.mode === 'demo' && (
        <p className="demo-note">
          <ShieldCheck size={16} />
          {isDemoModel
            ? '当前使用示例内容与固定输出，体验流程无需连接模型。'
            : '当前文档为模拟文档；生成时会实际调用选中的模型服务。'}
        </p>
      )}
      <div className="workspace-grid">
        <section className="compose-area" aria-label="输入与任务">
          <div className="section-heading">
            <h2>选择任务</h2>
            <span className="muted">{tasks.length} 项能力</span>
          </div>
          <div className="task-options" role="group" aria-label="选择任务">
            {tasks.map((item) => (
              <button
                key={item.id}
                disabled={active}
                aria-pressed={taskId === item.id}
                className={taskId === item.id ? 'task-option selected' : 'task-option'}
                onClick={() => {
                  setTaskId(item.id);
                  invalidate();
                }}
              >
                {item.label}
                {taskId === item.id && <Check size={14} />}
              </button>
            ))}
          </div>
          <p className="task-description">{task.description}</p>
          <div className="section-heading source-heading">
            <h2>
              <label htmlFor="source-text">参考内容</label>
            </h2>
            <button
              className="text-button"
              onClick={read}
              disabled={active || !bridge.capabilities.read}
            >
              {busy === 'reading' ? (
                <LoaderCircle size={16} className="spin" />
              ) : (
                <ScanText size={16} />
              )}
              {bridge.mode === 'demo' ? '读取示例选区' : '读取选区'}
            </button>
          </div>
          <textarea
            id="source-text"
            className="source-input"
            value={source}
            maxLength={MAX_SOURCE_LENGTH}
            disabled={active}
            onChange={(e) => editSource(e.target.value)}
            placeholder={
              bridge.host === 'excel'
                ? '在 Excel 中选中数据或公式，点击「读取选区」。也可以粘贴表格内容。'
                : bridge.host === 'powerpoint'
                  ? '在幻灯片中进入文本编辑状态并选中文字，点击「读取选区」。也可以直接粘贴文案。'
                  : '选中文档中的一段文字，再点击「读取选区」。也可以直接在这里粘贴内容。'
            }
            aria-describedby="source-info"
          />
          <div className="source-meta" id="source-info">
            <span>{sourceLabel}</span>
            <span>{source.length.toLocaleString()} 字符</span>
          </div>
          {task.id === 'translate' && (
            <label className="language-field">
              目标语言
              <select
                value={language}
                disabled={active}
                onChange={(e) => {
                  setLanguage(e.target.value);
                  invalidate();
                }}
              >
                {['英语', '简体中文', '繁体中文', '日语', '韩语', '法语', '德语', '西班牙语'].map(
                  (lang) => (
                    <option key={lang}>{lang}</option>
                  ),
                )}
              </select>
            </label>
          )}
          <label className="instruction-label" htmlFor="instruction">
            {task.needsInstruction ? '具体要求' : '补充要求'}{' '}
            <span className="muted">{task.needsInstruction ? '必填' : '可选'}</span>
          </label>
          <textarea
            id="instruction"
            className="instruction-input"
            value={instruction}
            maxLength={MAX_INSTRUCTION_LENGTH}
            disabled={active}
            onChange={(e) => {
              setInstruction(e.target.value);
              invalidate();
            }}
            placeholder={
              task.id === 'formula'
                ? '例如：计算 B2:B4 的合计，结果放在 B6'
                : task.id === 'outline'
                  ? '例如：为季度工作回顾生成 5 页演示大纲，面向团队成员'
                  : task.id === 'ask'
                    ? '例如：哪个区域的季度合计最高？'
                    : task.id === 'generate'
                      ? '例如：写一封项目进度邮件，包含已完成事项和下周计划'
                      : '例如：语气更简洁，保留所有数字与专业术语'
            }
          />
          <div className="generate-row">
            <button
              className="primary generate-button"
              onClick={generate}
              disabled={
                active ||
                (!model && !isDemoModel) ||
                (task.needsSource && !source.trim()) ||
                (task.needsInstruction && !instruction.trim())
              }
            >
              {busy === 'generating' ? (
                <LoaderCircle size={18} className="spin" />
              ) : (
                <Sparkles size={18} />
              )}
              {busy === 'generating' ? '正在生成…' : isDemoModel ? '生成演示预览' : '生成预览'}
              {!busy && <ArrowRight size={17} />}
            </button>
            {busy === 'generating' && (
              <button className="secondary" onClick={() => request.current?.abort()}>
                <Square size={14} />
                停止
              </button>
            )}
          </div>
          <p className="privacy-note">
            <ShieldCheck size={14} />
            {isDemoModel
              ? '演示模式不会向模型发送内容，也不会修改 Office 文件。'
              : '生成时仅将上方参考内容和要求发送给所选模型服务。'}
          </p>
        </section>
        <section
          ref={outputRef}
          tabIndex={-1}
          className={`result-area ${result ? 'has-result' : ''}`}
          aria-labelledby="result-heading"
          aria-busy={busy === 'generating'}
        >
          <div className="section-heading">
            <h2 id="result-heading">结果预览</h2>
            {result ? (
              <span className={`result-status ${written ? 'applied' : ''}`}>
                {written ? <CheckCheck size={14} /> : <FileCheck2 size={14} />}
                {written ? '已写入' : '待确认'}
              </span>
            ) : (
              <span className="muted">生成后由你确认</span>
            )}
          </div>
          {!result ? (
            <div className="empty-result">
              {busy === 'generating' ? (
                <>
                  <LoaderCircle size={36} className="spin" />
                  <h3>正在整理你的内容</h3>
                  <p>
                    生成过程中不会修改文档。
                    <br />
                    你可以随时停止请求。
                  </p>
                </>
              ) : (
                <>
                  <div className="preview-symbol">
                    <FileCheck2 size={34} strokeWidth={1.5} />
                  </div>
                  <h3>下一稿，从这里开始</h3>
                  <p>
                    读取内容，选择任务。
                    <br />
                    AI 结果会先显示在这里，
                    <br />
                    检查后再写入文档。
                  </p>
                  <div className="flow-hint">
                    <span>读取</span>
                    <ArrowRight size={13} />
                    <span>预览</span>
                    <ArrowRight size={13} />
                    <span>确认</span>
                  </div>
                </>
              )}
            </div>
          ) : (
            <>
              <p className="result-model">{resultModel}</p>
              {editing ? (
                <label className="result-edit-label">
                  编辑结果
                  <textarea
                    className="result-editor"
                    value={result}
                    maxLength={100_000}
                    disabled={active}
                    onChange={(e) => {
                      setResult(e.target.value);
                      setWritten(false);
                    }}
                  />
                </label>
              ) : slides ? (
                <ol className="slide-outline">
                  {slides.map((slide, i) => (
                    <li key={i}>
                      <span className="slide-number">{String(i + 1).padStart(2, '0')}</span>
                      <div>
                        <h3>{slide.title}</h3>
                        <ul>
                          {slide.bullets.map((bullet, b) => (
                            <li key={b}>{bullet}</li>
                          ))}
                        </ul>
                      </div>
                    </li>
                  ))}
                </ol>
              ) : (
                <pre className={`result-text ${task.output === 'formula' ? 'formula-text' : ''}`}>
                  {result}
                </pre>
              )}
              {formatError && (
                <div role="alert" className="notice error">
                  {formatError}
                </div>
              )}
              <div className="result-tools">
                <button
                  className="text-button"
                  onClick={copy}
                  disabled={active || Boolean(formatError)}
                >
                  <Copy size={16} />
                  复制
                </button>
                <button
                  className="text-button"
                  onClick={() => setEditing(!editing)}
                  disabled={active}
                >
                  {editing ? <Check size={16} /> : <Pencil size={16} />}
                  {editing ? '完成编辑' : '编辑'}
                </button>
                <button className="text-button" onClick={download} disabled={active}>
                  <ArrowDownToLine size={16} />
                  下载
                </button>
              </div>
              <div className="write-actions">
                {canReplace && (
                  <button
                    className="primary"
                    onClick={() => prepare('replace')}
                    disabled={active || written || !result.trim()}
                  >
                    <FileInput size={17} />
                    替换已读取选区
                  </button>
                )}
                {canInsert && (
                  <button
                    className={canReplace ? 'secondary' : 'primary'}
                    onClick={() => prepare('insert')}
                    disabled={active || written || !result.trim()}
                  >
                    插入到光标处
                  </button>
                )}
                {canFormula && (
                  <button
                    className="primary"
                    onClick={() => prepare('formula')}
                    disabled={active || written || Boolean(formatError)}
                  >
                    <FileInput size={17} />
                    写入选中单元格
                  </button>
                )}
                {canSlides && (
                  <button
                    className="primary"
                    onClick={() => prepare('slides')}
                    disabled={active || written || Boolean(formatError)}
                  >
                    <FileInput size={17} />
                    创建幻灯片
                  </button>
                )}
              </div>
              <p className="field-help">
                {canFormula
                  ? '先在 Excel 中选中一个目标单元格。确认时会显示位置和原内容。'
                  : canSlides
                    ? '确认后在演示文稿末尾添加页面；生成的是基础文字页，可在 Office 中排版。'
                    : canReplace || canInsert
                      ? '点击写入按钮后，还会展示目标和内容供你确认。'
                      : bridge.host === 'excel' && task.output === 'text'
                        ? '分析和解释供参考，可复制到你的工作记录。'
                        : '当前宿主或选区不支持自动写入，可复制结果后手动粘贴。'}
              </p>
            </>
          )}
        </section>
      </div>
      {notice && (
        <div
          className={`notice global-notice ${notice.error ? 'error' : 'success'}`}
          role={notice.error ? 'alert' : 'status'}
        >
          <span>{notice.text}</span>
          <button
            className="icon-button"
            aria-label="关闭提示"
            onClick={() => setNotice(undefined)}
          >
            <X size={16} />
          </button>
        </div>
      )}
      {plan && (
        <ConfirmWrite
          plan={plan}
          output={writeValue}
          busy={busy === 'writing'}
          error={writeError}
          onCancel={dismissPlan}
          onConfirm={commit}
        />
      )}
    </>
  );
}
