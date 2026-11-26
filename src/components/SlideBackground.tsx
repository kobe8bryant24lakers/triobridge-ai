import { useState, type CSSProperties } from 'react';
import { readMaterials } from '../presentation/materials';
import type { PresentationRuntime } from './presentation-runtime';
import { t } from '../i18n';
import type {
  BackgroundFill,
  Material,
  MaterialContext,
  SlideBackground,
} from '../presentation/types';

export type PickBackgroundPicture = (file: File, selected: (id: string) => void) => void;
export function useBackgroundPictures(run: PresentationRuntime['run']) {
  const [pictures, setPictures] = useState<Material[]>([]);
  function withPictures(context: MaterialContext): MaterialContext {
    return { ...context, materials: [...context.materials, ...pictures] };
  }
  const pick: PickBackgroundPicture = (file, selected) => {
    void run(
      t('读取本地背景图片', 'Reading background picture'),
      async (signal) => {
        const result = await readMaterials([file], signal);
        const material = result.materials[0];
        if (!material?.image)
          throw Error(
            result.issues[0]?.message ||
              t('请选择 PNG、JPEG 或 WebP 图片。', 'Choose a PNG, JPEG or WebP picture.'),
          );
        if (signal.aborted) return;
        const id = 'bg-' + crypto.randomUUID().slice(0, 20);
        setPictures((old) => [...old, { ...material, id }]);
        selected(id);
      },
      true,
    );
  };
  return { withPictures, pick };
}
export function fillStyle(
  fill: SlideBackground | undefined,
  image: (id: string) => string | undefined,
): CSSProperties {
  if (fill?.kind === 'solid') return { backgroundColor: fill.color };
  if (fill?.kind === 'gradient')
    return { backgroundImage: `linear-gradient(90deg, ${fill.from}, ${fill.to})` };
  if (fill?.kind === 'image') {
    const url = image(fill.imageId),
      alpha = fill.transparency / 100;
    if (url)
      return {
        backgroundColor: 'white',
        backgroundImage: `linear-gradient(rgb(255 255 255 / ${alpha}),rgb(255 255 255 / ${alpha})),url("${url}")`,
        backgroundSize: '100% 100%',
      };
  }
  return {};
}
function ColorField({
  label,
  value,
  onChange,
}: {
  label: string;
  value: string;
  onChange: (v: string) => void;
}) {
  return (
    <label>
      {label}
      <span className="ppt-color-field">
        <input
          type="color"
          aria-label={label + t('色板', ' swatch')}
          value={/^#[a-f\d]{6}$/i.test(value) ? value : '#ffffff'}
          onChange={(e) => onChange(e.target.value.toUpperCase())}
        />
        <input
          aria-label={label + ' HEX'}
          maxLength={7}
          value={value}
          placeholder="#F5F1E8"
          onChange={(e) => onChange(e.target.value.toUpperCase())}
        />
      </span>
    </label>
  );
}
export function BackgroundPicker({
  fill,
  onChange,
  context,
  pickPicture,
  allowDefault = false,
  defaultColor = '#F5F1E8',
  kinds = ['solid', 'gradient', 'image', 'master'],
}: {
  fill?: BackgroundFill;
  onChange: (f: BackgroundFill | undefined) => void;
  context: MaterialContext;
  pickPicture: PickBackgroundPicture;
  allowDefault?: boolean;
  defaultColor?: string;
  kinds?: BackgroundFill['kind'][];
}) {
  const pictures = context.materials.filter((m) => m.image);
  const initial: Record<BackgroundFill['kind'], BackgroundFill> = {
    solid: { kind: 'solid', color: defaultColor },
    gradient: { kind: 'gradient', from: defaultColor, to: '#B7CBDD' },
    image: { kind: 'image', imageId: pictures[0]?.id || '', transparency: 0 },
    master: { kind: 'master' },
  };
  const labels = {
    solid: t('纯色', 'Solid color'),
    gradient: t('双色渐变', 'Two-color gradient'),
    image: t('本地图片', 'Local picture'),
    master: t('跟随母版', 'Follow master'),
  };
  return (
    <div className="ppt-background-fields">
      <label>
        {t('背景样式', 'Background style')}
        <select
          value={fill?.kind || ''}
          onChange={(e) => onChange(initial[e.target.value as BackgroundFill['kind']])}
        >
          {allowDefault && <option value="">{t('使用整稿默认背景', 'Use deck default')}</option>}
          {kinds.map((k) => (
            <option key={k} value={k}>
              {labels[k]}
            </option>
          ))}
        </select>
      </label>
      {fill?.kind === 'solid' && (
        <ColorField
          label={t('背景色', 'Background color')}
          value={fill.color}
          onChange={(color) => onChange({ ...fill, color })}
        />
      )}
      {fill?.kind === 'gradient' && (
        <div className="ppt-field-grid">
          <ColorField
            label={t('起始色', 'Start color')}
            value={fill.from}
            onChange={(from) => onChange({ ...fill, from })}
          />
          <ColorField
            label={t('结束色', 'End color')}
            value={fill.to}
            onChange={(to) => onChange({ ...fill, to })}
          />
        </div>
      )}
      {fill?.kind === 'image' && (
        <>
          <label>
            {t('背景图片', 'Background picture')}
            <select
              value={fill.imageId}
              onChange={(e) => onChange({ ...fill, imageId: e.target.value })}
            >
              <option value="">{t('选择图片素材', 'Select a picture source')}</option>
              {pictures.map((m) => (
                <option key={m.id} value={m.id}>
                  {m.name}
                </option>
              ))}
            </select>
          </label>
          <label className="secondary ppt-file-button">
            {t('选择本地背景图片', 'Choose a local background picture')}
            <input
              type="file"
              accept="image/png,image/jpeg,image/webp"
              onChange={(e) => {
                const file = e.target.files?.[0];
                e.target.value = '';
                if (file) pickPicture(file, (id) => onChange({ ...fill, imageId: id }));
              }}
            />
          </label>
          <label>
            {t('图片透明度', 'Picture transparency')} · {fill.transparency}%
            <input
              type="range"
              min={0}
              max={100}
              value={fill.transparency}
              onChange={(e) => onChange({ ...fill, transparency: Number(e.target.value) })}
            />
          </label>
        </>
      )}
    </div>
  );
}
