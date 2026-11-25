import { createRoot } from 'react-dom/client';
import App from '../App';
import { getUiLanguage, t } from '../i18n';
import { HOST_NAMES } from '../core/tasks';
import { createNativeBridge } from './office';
import { nativeWorkbench } from './transport';
import '../styles.css';

document.documentElement.lang = getUiLanguage();
const native = nativeWorkbench();
if (native)
  document.title = HOST_NAMES[native.host] + t(' 工作台 · 三桥 AI', ' Workbench · TrioBridge AI');
createRoot(document.getElementById('root')!).render(
  native ? (
    <App officeBridge={createNativeBridge(native)} />
  ) : (
    <main className="boot-screen">
      <h1>{t('请从 Office 打开三桥 AI', 'Open TrioBridge AI from Office')}</h1>
      <p>
        {t(
          'Windows 使用“三桥 AI”功能区，Mac 使用应用菜单中的“服务”。',
          'On Windows, use the TrioBridge AI ribbon. On Mac, open Services in the application menu.',
        )}
      </p>
    </main>
  ),
);
