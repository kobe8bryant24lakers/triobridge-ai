import { createOfficeBridge } from './bridge';
import type { HostBridge, HostKind } from '../core/types';

const OFFICE_SCRIPT = 'https://appsforoffice.microsoft.com/lib/1/hosted/office.js';

export async function initializeOffice(): Promise<HostBridge> {
  await new Promise<void>((resolve, reject) => {
    if (typeof Office !== 'undefined') {
      resolve();
      return;
    }
    const script = document.createElement('script');
    const timer = setTimeout(() => {
      script.remove();
      reject(
        new Error(
          'Office.js 加载超时。请检查网络是否能访问 appsforoffice.microsoft.com，然后重试。',
        ),
      );
    }, 15000);
    script.src = OFFICE_SCRIPT;
    script.onload = () => {
      clearTimeout(timer);
      resolve();
    };
    script.onerror = () => {
      clearTimeout(timer);
      script.remove();
      reject(new Error('Office.js 加载失败，请检查网络或代理设置。'));
    };
    document.head.append(script);
  });
  const info = await new Promise<{
    host: Office.HostType;
    platform: Office.PlatformType;
  }>((resolve, reject) => {
    const timer = setTimeout(
      () =>
        reject(
          new Error(
            '未连接到 Office。请从 Word、Excel 或 PowerPoint 的加载项入口打开。浏览器体验请移除地址中的 ?office=1。',
          ),
        ),
      15000,
    );
    Office.onReady().then(
      (value) => {
        clearTimeout(timer);
        resolve(value);
      },
      () => {
        clearTimeout(timer);
        reject(new Error('Office 初始化失败，请重新打开加载项。'));
      },
    );
  });
  const hosts: Record<string, HostKind> = {
    Word: 'word',
    Excel: 'excel',
    PowerPoint: 'powerpoint',
  };
  const host = hosts[info.host];
  if (!host) throw new Error('当前宿主不受支持，请在 Word、Excel 或 PowerPoint 中打开此加载项。');
  return createOfficeBridge(host);
}
