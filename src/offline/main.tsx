import { createRoot } from 'react-dom/client';
import App from '../App';
import { createNativeBridge } from './office';
import { nativeWorkbench } from './transport';
import '../styles.css';

const native = nativeWorkbench();
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
