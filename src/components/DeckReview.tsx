import { useState } from 'react';
import { ListChecks } from 'lucide-react';
import { t } from '../i18n';
import type { DeckContent } from '../presentation/types';
export function deckFindings(deck: DeckContent) {
  const seen = new Set<string>();
  return deck.slides.flatMap((slide, index) => {
    const reasons: string[] = [];
    if (!slide.sources.length)
      reasons.push(t('补上依据或核实来源', 'Check the source for this claim'));
    if (slide.title.length > 42)
      reasons.push(t('标题偏长，可提炼成一句结论', 'Condense the long title into one claim'));
    const normalized = slide.title.trim().toLocaleLowerCase();
    if (seen.has(normalized))
      reasons.push(
        t('与前页标题重复，明确这页的作用', 'Differentiate this slide from the earlier one'),
      );
    seen.add(normalized);
    if (slide.chart && !slide.chart.unit)
      reasons.push(t('确认图表是否需要标明单位', 'Check the chart unit'));
    const bullets = [...(slide.bullets || []), ...(slide.columns || []).flatMap((c) => c.bullets)];
    if (bullets.some((b) => b.length > 55))
      reasons.push(t('要点较长，可精简或重排', 'Tighten or reflow the longer bullets'));
    if (!slide.notes.trim()) reasons.push(t('补充演讲备注', 'Add speaker notes'));
    return reasons.length ? [{ index, title: slide.title, reasons }] : [];
  });
}
export function DeckReview({
  deck,
  busy,
  onSelect,
}: {
  deck: DeckContent;
  busy: boolean;
  onSelect: (n: number) => void;
}) {
  const [open, setOpen] = useState(false),
    items = deckFindings(deck);
  return (
    <div className="deck-review">
      <button
        type="button"
        className="secondary"
        disabled={busy}
        aria-expanded={open}
        onClick={() => setOpen(!open)}
      >
        <ListChecks size={16} />
        {t('检查成稿', 'Review deck')}
      </button>
      {open && (
        <div className="theme-sheet" aria-live="polite">
          {items.length ? (
            <>
              <p className="field-help">
                {t(
                  '点页标题查看原页，再决定如何调整。',
                  'Select a title to inspect the slide and decide what to change.',
                )}
              </p>
              <ol>
                {items.map((v) => (
                  <li key={v.index}>
                    <button
                      className="text-button"
                      disabled={busy}
                      onClick={() => {
                        onSelect(v.index);
                        setOpen(false);
                      }}
                    >
                      {v.index + 1}. {v.title}
                    </button>
                    <div className="field-help">{v.reasons.join(' · ')}</div>
                  </li>
                ))}
              </ol>
            </>
          ) : (
            <p className="field-help">
              {t(
                '标题、来源、图表单位和备注检查完成，未检出上述问题。',
                'Checked titles, sources, chart units and notes. No issues found in these checks.',
              )}
            </p>
          )}
        </div>
      )}
    </div>
  );
}
