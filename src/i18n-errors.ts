import { getUiLanguage } from './i18n';
import { translateModelMessage } from './core/model-ui-messages';
import { officeUiMessages } from './office/ui-messages';

const messages: Readonly<Record<string, string>> = {
  '请先选择模型服务。': 'Select a model service first.',
  '请填写所选服务的 API Key。': 'Enter the API key for this service.',
  '配置文件过大。': 'The configuration file is too large.',
  '服务地址或接口格式已更改，请填写新密钥或选择不使用密钥。':
    'The endpoint or API format changed. Enter a new key or select No key.',
  '请填写新密钥；无需密钥的服务请选择「不使用密钥」。':
    'Enter a new key, or select No key if the service does not require one.',
};

export function uiMessage(message: string): string {
  if (getUiLanguage() !== 'en') return message;
  if (Object.prototype.hasOwnProperty.call(messages, message)) return messages[message];
  if (Object.prototype.hasOwnProperty.call(officeUiMessages, message))
    return officeUiMessages[message];
  return translateModelMessage(message) ?? message;
}
