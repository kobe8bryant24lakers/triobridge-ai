import { nativeWorkbench } from './offline/transport';
import { useEffect, useMemo, useState, useSyncExternalStore } from 'react';
import {
  BookOpen,
  ChevronRight,
  FileText,
  Grid2X2,
  LayoutDashboard,
  Monitor,
  Presentation,
  Settings2,
} from 'lucide-react';
import { ModelSettings } from './components/ModelSettings';
import { Workbench } from './components/Workbench';
import { HOST_NAMES } from './core/tasks';
import { createProfileController, type LocalProfileState } from './core/profile-store';
import type { HostBridge, HostKind } from './core/types';
import { createDemoBridge } from './office/demo';
import type { DemoScenario } from './office/complex-samples';
import type { PresentationMode } from './components/PresentationWorkspace';

const ICONS = { word: FileText, excel: Grid2X2, powerpoint: Presentation };
interface Props {
  officeBridge?: HostBridge;
}

export default function App({ officeBridge }: Props) {
  const [host, setHost] = useState<HostKind>(() => {
    const demoHost = new URLSearchParams(window.location.search).get('demo-host');
    return (
      officeBridge?.host || (demoHost === 'excel' || demoHost === 'powerpoint' ? demoHost : 'word')
    );
  });
  const [page, setPage] = useState<'work' | 'settings' | 'help'>('work');
  const [presentationMode, setPresentationMode] = useState<PresentationMode>('generate');
  const profileController = useMemo(() => createProfileController(), []);
  const profiles = useSyncExternalStore(profileController.subscribe, profileController.getSnapshot);
  useEffect(() => {
    profileController.start();
    return () => profileController.stop();
  }, [profileController]);
  const [busy, setBusy] = useState(false);
  const [demoDocument, setDemoDocument] = useState('');
  const [demoScenario, setDemoScenario] = useState<DemoScenario>('basic');
  const bridge = useMemo(
    () => officeBridge || createDemoBridge(host, setDemoDocument, demoScenario),
    [host, officeBridge, demoScenario],
  );
  useEffect(() => () => bridge.dispose?.(), [bridge]);
  function save(next: LocalProfileState) {
    profileController.saveLocal(next);
  }
  function switchHost(next: HostKind) {
    if (next === host) {
      setPage('work');
      return;
    }
    setHost(next);
    setDemoDocument('');
    setPage('work');
  }
  return (
    <div className={`app-shell ${officeBridge ? 'office-shell' : ''}`}>
      <a className="skip-link" href="#main">
        跳到工作区
      </a>
      <aside className="sidebar" aria-label="应用导航">
        <a href="#main" className="brand" aria-label="TrioBridge AI 工作区">
          <img src={`${import.meta.env.BASE_URL}icon.svg`} alt="" width="35" height="35" />
          <div>
            <strong>
              TrioBridge <span>AI</span>
            </strong>
            <span className="brand-subtitle">三桥 · 你的办公 AI</span>
          </div>
        </a>
        <div className="sidebar-main">
          <div className="app-picker" aria-label="选择 Office 应用">
            {(Object.keys(HOST_NAMES) as HostKind[]).map((kind) => {
              const Icon = ICONS[kind];
              return (
                <button
                  key={kind}
                  className={`app-nav ${host === kind ? 'active' : ''}`}
                  aria-pressed={host === kind}
                  disabled={busy || Boolean(officeBridge && kind !== host)}
                  onClick={() => switchHost(kind)}
                >
                  <Icon size={19} />
                  <span>{HOST_NAMES[kind]}</span>
                  <ChevronRight size={15} className="nav-chevron" />
                </button>
              );
            })}
          </div>
          <nav className="utility-nav">
            <button
              className={page === 'work' ? 'active' : ''}
              onClick={() => setPage('work')}
              disabled={busy}
            >
              <LayoutDashboard size={18} />
              工作台
            </button>
            <button
              className={page === 'settings' ? 'active' : ''}
              onClick={() => setPage('settings')}
              disabled={busy}
            >
              <Settings2 size={18} />
              模型设置
            </button>
            <button
              className={page === 'help' ? 'active' : ''}
              onClick={() => setPage('help')}
              disabled={busy}
            >
              <BookOpen size={18} />
              使用指南
            </button>
          </nav>
        </div>
        <div className="sidebar-foot">
          <div className="bridge-mark">
            <i />
            <i />
            <i />
          </div>
          <p>
            你的文档。你的模型。
            <br />
            把选择留在你手里。
          </p>
          <span>TrioBridge AI · 0.6.0</span>
        </div>
      </aside>
      <div className="main-shell">
        <header className="topbar">
          <span>
            {page === 'settings'
              ? '模型设置'
              : page === 'help'
                ? '使用指南'
                : `${HOST_NAMES[host]} / 工作台`}
          </span>
          <span>
            <Monitor size={14} />
            {nativeWorkbench() ? '本机工作台' : officeBridge ? 'Office 加载项' : '交互演示'}
          </span>
        </header>
        <main id="main" className="main-content">
          {profiles.warning && (
            <div className="notice error" role="alert">
              {profiles.warning}
            </div>
          )}
          {page === 'settings' ? (
            <ModelSettings
              state={profiles}
              onSave={save}
              sharedActions={profileController.actions}
              onBack={() => setPage('work')}
            />
          ) : page === 'help' ? (
            <section className="help-page">
              <h1>把 AI 接入你的工作</h1>
              <p>
                支持 Mac 与 Windows，不要求订阅 Microsoft 365。永久授权版 Office
                的可用功能由版本和运行环境决定。
              </p>
              <ol className="help-steps">
                <li>
                  <h2>连接模型</h2>
                  <p>
                    在模型设置中填写服务地址、模型 ID 和 API Key，测试连接后保存。使用 OpenAI 兼容的
                    Chat Completions 接口。
                  </p>
                </li>
                <li>
                  <h2>提供上下文</h2>
                  <p>
                    Word / Excel 可读取选区。PowerPoint
                    可选择本地素材目录，或读取选中页面、整份演示稿。素材先在本机解析，点击生成时才发送勾选内容；图片理解需另外开启。
                  </p>
                </li>
                <li>
                  <h2>检查后写入</h2>
                  <p>
                    选择任务并生成预览。可以编辑结果，核对目标位置后确认写入。Word
                    报告支持章节和表格；Excel 可写入单个公式，或创建独立的数据工作表。PPT
                    支持生成可编辑整稿、逐项应用局部改动、插入原生图表；建议先下载原稿备份。
                  </p>
                </li>
              </ol>
              <h2>不同版本，按能力适配</h2>
              <p>
                基础文本使用通用 Office API；Word 和 Excel 单公式使用 1.1 基础 API，Excel
                新建格式化数据表需要 ExcelApi 1.2。PPT 对象编辑需要 PowerPointApi 1.4，选中页需要
                1.5，重排需要 1.8；插入整页需要 1.2 及可读取页面尺寸的 1.10。不支持时仍可下载
                PPTX。旧式 IE / EdgeHTML WebView 不在支持范围内。
              </p>
              <h2>本地模型与隐私</h2>
              <p>
                在线版直接连接模型，服务需允许当前来源的
                CORS。本机共享版由本机服务调用已保存的模型，
                不要求模型服务允许浏览器跨域访问，密钥保存在系统安全存储中。应用不保存文档或生成历史。
              </p>
              <p>
                本机共享版需要运行本机服务并安装本机专用清单。在同一台电脑、同一系统账号下，
                只需配置一次，Word、Excel、PowerPoint 即可共用多个模型并同步切换。
                在线安装包和纯静态网页使用各应用独立的配置；导出配置不包含密钥。
              </p>
              <a
                className="doc-link"
                href="https://github.com/kobe8bryant24lakers/triobridge-ai#readme"
                target="_blank"
                rel="noreferrer"
              >
                安装说明与兼容性文档 <ChevronRight size={16} />
              </a>
            </section>
          ) : null}
          <div hidden={page !== 'work'}>
            <Workbench
              key={`${host}-${demoScenario}`}
              bridge={bridge}
              profiles={profiles}
              onChooseModel={(id) => profileController.activate(id)}
              onSettings={() => setPage('settings')}
              onBusyChange={setBusy}
              demoScenario={demoScenario}
              presentationMode={presentationMode}
              onPresentationModeChange={setPresentationMode}
              onDemoScenarioChange={(next) => {
                setDemoScenario(next);
                setDemoDocument('');
              }}
            />
            {demoDocument && (
              <details className="demo-document">
                <summary>查看模拟文档 · 已应用的内容</summary>
                <pre>{demoDocument}</pre>
              </details>
            )}
          </div>
        </main>
        <footer className="main-footer">
          <span>TrioBridge AI · 三桥 AI</span>
          <span>内容由你审核，改动由你确认。</span>
        </footer>
      </div>
    </div>
  );
}
