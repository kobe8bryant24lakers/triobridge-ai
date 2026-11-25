import { createRoot } from 'react-dom/client';
import App from '../App';
import { HOST_NAMES } from '../core/tasks';
import { createNativeBridge } from './office';
import { nativeWorkbench } from './transport';
import '../styles.css';

const native = nativeWorkbench();
if (native) document.title = HOST_NAMES[native.host] + ' 工作台 · 三桥 AI';
createRoot(document.getElementById('root')!).render(
  native ? (
    <App officeBridge={createNativeBridge(native)} />
  ) : (
    <main className="boot-screen">
      <h1>请从 Office 打开三桥 AI</h1>
      <p>Windows 使用“三桥 AI”功能区，Mac 使用应用菜单中的“服务”。</p>
    </main>
  ),
);
