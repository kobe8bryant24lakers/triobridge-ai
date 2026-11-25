// Only these fixed values may cross the installer diagnostic boundary.
const details = {
  CATALOG_BUSY: '共享配置文件暂时被占用。请待占用释放后重试。',
  DIRECTORY_ACCESS: '当前账号无法访问共享配置目录。',
  CAPABILITY_ACCESS:
    '后台服务无法读取本次安装的共享访问凭据。请用平时打开 Office 的账号运行 Install。',
  ACL_INSPECT: '后台服务无法检查共享配置的 Windows 访问权限。',
  ACL_OWNER: '共享文件的所有者与运行服务的账号不匹配。请使用原安装账号。',
  ACL_APPLY: '后台服务无法设置共享文件的专用权限。请检查当前账号的目录权限。',
  ACL_VERIFY: '共享文件权限校验未通过。请检查现有文件所有者和访问权限。',
  CATALOG_ACCESS: '后台服务无法读取 profiles.json。请检查当前账号的文件访问权限。',
  CATALOG_UNSAFE: 'profiles.json 的文件类型或所有者不符合要求。请检查原文件。',
  CATALOG_TOO_LARGE: 'profiles.json 超出共享配置大小限制。请保留原文件并检查备份。',
  CATALOG_JSON: 'profiles.json 无法解析为 JSON。请保留原文件并核对可信备份。',
  CATALOG_SCHEMA: 'profiles.json 的结构或版本无法识别。请保留原文件并核对可信备份。',
  STORE_INIT: '后台服务无法初始化共享配置组件。',
};

export const storageDetailHeader = 'x-triobridge-storage-detail';

export function storageDetailMessage(value) {
  return typeof value === 'string' && Object.hasOwn(details, value) ? details[value] : undefined;
}

export function storageFailureDetail(error, fallback) {
  if (storageDetailMessage(error?.storageDetail)) return error.storageDetail;
  if (['EBUSY', 'EAGAIN', 'ETXTBSY'].includes(error?.code)) return 'CATALOG_BUSY';
  return fallback;
}
