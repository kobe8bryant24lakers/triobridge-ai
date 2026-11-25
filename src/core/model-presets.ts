import type { ProfileDetails } from './types';

// Defaults verified against the providers' API documentation on 2026-11-21.
// These are editable starting points; access and browser CORS still depend on the provider.
export const MODEL_PRESETS = [
  {
    id: 'deepseek',
    name: 'DeepSeek',
    baseUrl: 'https://api.deepseek.com',
    model: 'deepseek-flash',
    docsUrl: 'https://api-docs.deepseek.com/',
  },
  {
    id: 'openai',
    name: 'OpenAI',
    baseUrl: 'https://api.openai.com/v1',
    model: 'gpt-4.1-mini',
    docsUrl: 'https://developers.openai.com/api/docs/models/gpt-4.1-mini',
  },
  {
    id: 'qwen',
    name: '通义千问（北京）',
    baseUrl: 'https://dashscope.aliyuncs.com/compatible-mode/v1',
    model: 'qwen-plus',
    docsUrl: 'https://help.aliyun.com/zh/model-studio/compatibility-of-openai-with-dashscope',
  },
] as const;

export function presetFor(profile: ProfileDetails): string {
  if (!profile.name && !profile.model) return '';
  if (profile.apiFormat === 'anthropic') return 'custom';
  return (
    MODEL_PRESETS.find(
      (preset) => preset.baseUrl === profile.baseUrl && preset.model === profile.model,
    )?.id || 'custom'
  );
}

export function customFormatSuffix(profile: ProfileDetails): string {
  if (presetFor(profile) !== 'custom') return '';
  return profile.apiFormat === 'anthropic' ? ' · Claude 格式' : ' · OpenAI 格式';
}
