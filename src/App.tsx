import { isUiLanguage, t, useUiLanguage } from './i18n';
import { uiMessage } from './i18n-errors';
import { nativeWorkbench } from './offline/transport';
import { useEffect, useMemo, useState, useSyncExternalStore } from 'react';
import {
  BookOpen,
  ChevronRight,
  FileText,
  Grid2X2,
  LayoutDashboard,
  Languages,
  Monitor,
  Presentation,
  Settings2,
} from 'lucide-react';
import { ModelSettings } from './components/ModelSettings';
import { Workbench } from './components/Workbench';
import { WordTextPreview } from './components/WordTextPreview';
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
  const language = useUiLanguage();
  const [savingLanguage, setSavingLanguage] = useState(false);
  const [languageSaveFailed, setLanguageSaveFailed] = useState(false);
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
  useEffect(() => {
    document.documentElement.lang = language;
    document.title = officeBridge
      ? t(HOST_NAMES[host] + ' 工作台 · 三桥 AI', HOST_NAMES[host] + ' Workbench · TrioBridge AI')
      : t('三桥 AI · 办公工作台', 'TrioBridge AI · Office Workbench');
  }, [language, host, officeBridge]);
  async function chooseLanguage(value: string) {
    if (!isUiLanguage(value)) return;
    setSavingLanguage(true);
    setLanguageSaveFailed(false);
    try {
      await profileController.chooseLanguage(value);
    } catch {
      setLanguageSaveFailed(true);
    } finally {
      setSavingLanguage(false);
    }
  }
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
        {t('跳到工作区', 'Skip to workspace')}
      </a>
      <aside className="sidebar" aria-label={t('应用导航', 'App navigation')}>
        <div className="brand-header">
          <a
            href="#main"
            className="brand"
            aria-label={t('TrioBridge AI 工作区', 'TrioBridge AI workspace')}
          >
            <img src={`${import.meta.env.BASE_URL}icon.svg`} alt="" width="35" height="35" />
            <div>
              <strong>
                TrioBridge <span>AI</span>
              </strong>
              <span className="brand-subtitle">
                {t('三桥 · 你的办公 AI', 'Your Office workbench')}
              </span>
            </div>
          </a>
          <label className="language-control">
            <Languages size={16} aria-hidden="true" />
            <select
              aria-label={t('界面语言 / Interface language', 'Interface language')}
              value={language}
              disabled={savingLanguage}
              onChange={(event) => void chooseLanguage(event.target.value)}
            >
              <option value="zh-CN" lang="zh-CN">
                简体中文
              </option>
              <option value="en" lang="en">
                English
              </option>
            </select>
          </label>
        </div>
        <div className="sidebar-main">
          <div className="app-picker" aria-label={t('选择 Office 应用', 'Choose an Office app')}>
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
              {t('工作台', 'Workbench')}
            </button>
            <button
              className={page === 'settings' ? 'active' : ''}
              onClick={() => setPage('settings')}
              disabled={busy}
            >
              <Settings2 size={18} />
              {t('模型设置', 'Model Settings')}
            </button>
            <button
              className={page === 'help' ? 'active' : ''}
              onClick={() => setPage('help')}
              disabled={busy}
            >
              <BookOpen size={18} />
              {t('使用指南', 'User Guide')}
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
            {t('你的文档。你的模型。', 'Your documents. Your models.')}
            <br />
            {t('把选择留在你手里。', 'You choose what to share.')}
          </p>
          <span>TrioBridge AI · 0.9.0</span>
        </div>
      </aside>
      <div className="main-shell">
        <header className="topbar">
          <span>
            {page === 'settings'
              ? t('模型设置', 'Model Settings')
              : page === 'help'
                ? t('使用指南', 'User Guide')
                : `${HOST_NAMES[host]} / ${t('工作台', 'Workbench')}`}
          </span>
          <span>
            <Monitor size={14} />
            {nativeWorkbench()
              ? t('本机工作台', 'Desktop workbench')
              : officeBridge
                ? t('Office 加载项', 'Office add-in')
                : t('交互演示', 'Interactive demo')}
          </span>
        </header>
        <main id="main" className="main-content">
          {languageSaveFailed && (
            <div role="status" className="notice">
              {t(
                '此窗口已切换语言。未能保存到本机共享设置，请检查本机服务。',
                'The language changed in this window. Check the local service to save it for other Office windows.',
              )}
            </div>
          )}
          {profiles.warning && (
            <div className="notice error" role="alert">
              {uiMessage(profiles.warning)}
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
              <h1>{t('把 AI 接入你的工作', 'Using TrioBridge AI in Office')}</h1>
              <p>
                {t(
                  '支持 Mac 与 Windows，不要求订阅 Microsoft 365。永久授权版 Office 的可用功能由版本和运行环境决定。',
                  'Available for desktop Office on Windows and Mac. Available Office features depend on the installed version and runtime. A Microsoft 365 subscription is optional.',
                )}
              </p>
              <ol className="help-steps">
                <li>
                  <h2>{t('连接模型', 'Connect a model')}</h2>
                  <p>
                    {t(
                      '在模型设置中填写服务地址和模型 ID。使用本机 API 时，按推理服务选择 OpenAI（Chat Completions）或 Claude（Anthropic Messages）。例如，基础地址可以填 http://127.0.0.1:8001/v1。按需填写 API Key，然后点击“连接并保存”。',
                      'In Model Settings, enter the service URL and model ID. Choose OpenAI (Chat Completions) or Claude (Anthropic Messages) to match your service. A local base URL might be http://127.0.0.1:8001/v1. Enter an API key if required, then choose Connect and Save.',
                    )}
                  </p>
                </li>
                <li>
                  <h2>{t('提供上下文', 'Choose the source')}</h2>
                  <p>
                    {t(
                      'Word / Excel 可读取选区。PowerPoint 可选择本地素材目录，或读取选中页面、整份演示稿。素材先在本机解析，点击生成时才发送勾选内容；图片理解需另外开启。',
                      'In Word or Excel, read the selection. In PowerPoint, choose source files or read selected slides or the entire deck. Files are parsed locally. Generation sends the checked sources to your model. Turn on image analysis separately to include image content.',
                    )}
                  </p>
                </li>
                <li>
                  <h2>{t('检查后写入', 'Review and write')}</h2>
                  <p>
                    {t(
                      '选择任务并生成预览。可以编辑结果，核对目标位置后确认写入。Word 报告支持章节和表格；Excel 可写入单个公式，或创建独立的数据工作表。PPT 支持生成带图表的可编辑整稿，也能调整已有页面的文字、样式和位置。',
                      'Choose a task and generate a preview. Review and edit the result, then confirm the destination. Word reports can contain sections and tables. Excel can write a formula or create a worksheet. PowerPoint can generate editable decks with charts or change slide text, style and position.',
                    )}
                  </p>
                </li>
              </ol>
              <h2>{t('从桌面 Office 打开', 'Open from desktop Office')}</h2>
              <p>
                {t(
                  'Windows 在“三桥 AI”功能区选择“打开侧边栏”或“打开独立窗口”。Mac 在当前应用的“服务”菜单打开独立窗口，启用内嵌窗格后也可从“我的加载项”打开。 首次读写文档时，macOS 可能请求允许工作台控制该应用。',
                  'On Windows, use the TrioBridge AI ribbon to open the pane or a separate window. On Mac, use the Services menu for a standalone window, or My Add-ins after enabling the task pane. macOS may ask permission to control the Office app on first use.',
                )}
              </p>
              <h2>{t('本地模型与隐私', 'Local models and privacy')}</h2>
              <p>
                {t(
                  '工作台从安装包读取页面。模型请求发往你设置的 API，本机推理需要先运行你自己的模型。 Word、Excel 和 PowerPoint 共用当前系统账户下的模型列表和选择。',
                  'The workbench loads its interface from the installed package. Model requests go to the API you configure. Start your local model server before using it. The three Office apps share the current account’s model list and selection.',
                )}
              </p>
              <p>
                {t(
                  '密钥保存在 macOS 钥匙串或 Windows 凭据管理器中。导出配置只包含模型设置。',
                  'Keys are stored in macOS Keychain or Windows Credential Manager. Exported profiles contain model settings.',
                )}
              </p>
              {!nativeWorkbench() && (
                <a
                  className="doc-link"
                  href="https://github.com/kobe8bryant24lakers/triobridge-ai#readme"
                  target="_blank"
                  rel="noreferrer"
                >
                  {t('安装说明与兼容性文档', 'Installation and compatibility')}{' '}
                  <ChevronRight size={16} />
                </a>
              )}
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
                <summary>
                  {t('查看模拟文档 · 已应用的内容', 'Demo document · applied changes')}
                </summary>
                {host === 'word' ? (
                  <WordTextPreview text={demoDocument} />
                ) : (
                  <pre>{demoDocument}</pre>
                )}
              </details>
            )}
          </div>
        </main>
        <footer className="main-footer">
          <span>{t('TrioBridge AI · 三桥 AI', 'TrioBridge AI')}</span>
          <span>
            {t('内容由你审核，改动由你确认。', 'Review the result before applying changes.')}
          </span>
        </footer>
      </div>
    </div>
  );
}
