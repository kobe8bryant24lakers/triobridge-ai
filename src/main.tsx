import { Component, type ReactNode } from 'react';
import { createRoot } from 'react-dom/client';
import App from './App';
import { initializeOffice } from './office/runtime';
import { bootstrapLocalCapability } from './core/local-capability';
import './styles.css';

bootstrapLocalCapability();

class ErrorBoundary extends Component<{ children: ReactNode }, { failed: boolean }> {
  state = { failed: false };
  static getDerivedStateFromError() {
    return { failed: true };
  }
  render() {
    return this.state.failed ? (
      <div className="boot-screen">
        <h1>界面暂时无法显示</h1>
        <p>请重新加载。此操作不会修改 Office 文档。</p>
        <button className="primary" onClick={() => location.reload()}>
          重新加载
        </button>
      </div>
    ) : (
      this.props.children
    );
  }
}

const root = createRoot(document.getElementById('root')!);
async function start() {
  if (new URLSearchParams(location.search).get('office') !== '1') {
    root.render(
      <ErrorBoundary>
        <App />
      </ErrorBoundary>,
    );
    return;
  }
  root.render(
    <div className="boot-screen" role="status">
      <img src={`${import.meta.env.BASE_URL}icon.svg`} width="44" height="44" alt="" />
      <h1>正在连接 Office</h1>
      <p>首次加载需要访问 Office.js 服务。</p>
    </div>,
  );
  try {
    const bridge = await initializeOffice();
    root.render(
      <ErrorBoundary>
        <App officeBridge={bridge} />
      </ErrorBoundary>,
    );
  } catch (error) {
    root.render(
      <div className="boot-screen">
        <h1>未能连接 Office</h1>
        <p role="alert">{error instanceof Error ? error.message : '请重新打开加载项。'}</p>
        <button className="primary" onClick={() => location.reload()}>
          重新连接
        </button>
      </div>,
    );
  }
}
void start();
