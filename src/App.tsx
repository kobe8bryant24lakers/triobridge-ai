import { useEffect, useMemo, useState } from 'react';
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
import { loadProfiles, saveProfiles, type ProfileState } from './core/profile-store';
import type { HostBridge, HostKind } from './core/types';
import { createDemoBridge } from './office/demo';

const ICONS = { word: FileText, excel: Grid2X2, powerpoint: Presentation };
interface Props {
  officeBridge?: HostBridge;
}

export default function App({ officeBridge }: Props) {
  const [host, setHost] = useState<HostKind>(officeBridge?.host || 'word');
  const [page, setPage] = useState<'work' | 'settings' | 'help'>('work');
  const [profiles, setProfiles] = useState(loadProfiles);
  const [busy, setBusy] = useState(false);
  const [demoDocument, setDemoDocument] = useState('');
  const bridge = useMemo(
    () => officeBridge || createDemoBridge(host, setDemoDocument),
    [host, officeBridge],
  );
  useEffect(() => () => bridge.dispose?.(), [bridge]);
  function save(next: ProfileState) {
    saveProfiles(next);
    setProfiles(next);
  }
  function switchHost(next: HostKind) {
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
          <span>TrioBridge AI · 0.2.0</span>
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
            {officeBridge ? 'Office 加载项' : '交互演示'}
          </span>
        </header>
        <main id="main" className="main-content">
          {profiles.warning && (
            <div className="notice error" role="alert">
              {profiles.warning}
            </div>
          )}
          {page === 'settings' ? (
            <ModelSettings state={profiles} onSave={save} onBack={() => setPage('work')} />
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
                    在 Office 中选中内容，再点击「读取选区」。PowerPoint
                    需要进入文字编辑状态。只有你读取或粘贴的内容会发送给模型。
                  </p>
                </li>
                <li>
                  <h2>检查后写入</h2>
                  <p>
                    选择任务并生成预览。可以编辑结果，核对目标位置后确认写入。Excel
                    公式每次写入一个单元格。
                  </p>
                </li>
              </ol>
              <h2>不同版本，按能力适配</h2>
              <p>
                基础文本使用通用 Office API；Word 和 Excel 使用 1.1 基础 API。PowerPoint
                自动建页需要 PowerPointApi 1.4，不满足时仍可生成大纲并复制。旧式 IE / EdgeHTML
                WebView 不在支持范围内。
              </p>
              <h2>本地模型与隐私</h2>
              <p>
                模型服务需要允许当前来源的 CORS。若服务不支持，或本地模型只提供 HTTP，可按 README
                配置本机代理。所有请求直接前往你配置的服务；应用不保存文档或生成历史。
              </p>
              <p>配置和密钥默认不跨 Office 应用自动同步。可导出不含密钥的配置，在其他应用导入。</p>
              <a
                className="doc-link"
                href="https://github.com/kobe8bryant24lakers/triobridge-ai#readme"
                target="_blank"
                rel="noreferrer"
              >
                安装说明与兼容性文档 <ChevronRight size={16} />
              </a>
            </section>
          ) : (
            <>
              <Workbench
                key={host}
                bridge={bridge}
                profiles={profiles}
                onChooseModel={(id) => save({ ...profiles, activeId: id })}
                onSettings={() => setPage('settings')}
                onBusyChange={setBusy}
              />
              {demoDocument && (
                <details className="demo-document">
                  <summary>查看模拟文档 · 已应用的内容</summary>
                  <pre>{demoDocument}</pre>
                </details>
              )}
            </>
          )}
        </main>
        <footer className="main-footer">
          <span>TrioBridge AI · 三桥 AI</span>
          <span>内容由你审核，改动由你确认。</span>
        </footer>
      </div>
    </div>
  );
}
