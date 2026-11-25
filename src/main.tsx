import { Component, type ReactNode } from 'react';
import { createRoot } from 'react-dom/client';
import App from './App';
import { initializeOffice } from './office/runtime';
import { bootstrapLocalCapability } from './core/local-capability';
import { getUiLanguage, t } from './i18n';
import { uiMessage } from './i18n-errors';
import './styles.css';

document.documentElement.lang = getUiLanguage();

bootstrapLocalCapability();

class ErrorBoundary extends Component<{ children: ReactNode }, { failed: boolean }> {
  state = { failed: false };
  static getDerivedStateFromError() {
    return { failed: true };
  }
  render() {
    return this.state.failed ? (
      <div className="boot-screen">
        <h1>{t('界面暂时无法显示', 'The workbench could not be displayed')}</h1>
        <p>
          {t(
            '请重新加载。此操作不会修改 Office 文档。',
            'Reload to try again. Your Office document is unchanged.',
          )}
        </p>
        <button className="primary" onClick={() => location.reload()}>
          {t('重新加载', 'Reload')}
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
      <h1>{t('正在连接 Office', 'Connecting to Office')}</h1>
      <p>
        {t('首次加载需要访问 Office.js 服务。', 'The first load requires access to Office.js.')}
      </p>
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
        <h1>{t('未能连接 Office', 'Could not connect to Office')}</h1>
        <p role="alert">
          {error instanceof Error
            ? uiMessage(error.message)
            : t('请重新打开加载项。', 'Reopen the add-in.')}
        </p>
        <button className="primary" onClick={() => location.reload()}>
          {t('重新连接', 'Reconnect')}
        </button>
      </div>,
    );
  }
}
void start();
