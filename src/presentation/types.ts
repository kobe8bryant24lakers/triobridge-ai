export type ChartKind = 'column' | 'bar' | 'line' | 'pie' | 'doughnut';
export interface ChartData {
  categories: string[];
  series: { name: string; values: (number | null)[] }[];
}
export type BackgroundFill =
  | { kind: 'master' }
  | { kind: 'solid'; color: string }
  | { kind: 'gradient'; from: string; to: string }
  | { kind: 'image'; imageId: string; transparency: number };
export type SlideBackground = BackgroundFill | { kind: 'other'; label: string };
export type PageEdit =
  | {
      kind: 'background';
      slideId: string;
      fill: BackgroundFill | { kind: 'reference'; sourceId: string };
    }
  | { kind: 'chart'; slideId: string; shapeId: string; data: ChartData }
  | {
      kind: 'insert';
      slideId: string;
      title: string;
      body: string;
      chart?: { kind: ChartKind; data: ChartData; unit: string };
    }
  | { kind: 'split'; slideId: string; shapeId: string; at: number; title: string };
export interface MaterialTable {
  id: string;
  name: string;
  headers: string[];
  rows: (string | number | boolean | null)[][];
}
export interface Material {
  id: string;
  name: string;
  kind: 'text' | 'document' | 'spreadsheet' | 'presentation' | 'pdf' | 'image';
  text: string;
  tables: MaterialTable[];
  image?: { data: string; modelData: string; width: number; height: number };
  warning?: string;
}
export interface MaterialIssue {
  name: string;
  message: string;
}
export interface MaterialContext {
  materials: Material[];
  includeImages: boolean;
}

export type SlideLayout =
  | 'cover'
  | 'bullets'
  | 'twoColumn'
  | 'metrics'
  | 'process'
  | 'chart'
  | 'image'
  | 'table'
  | 'closing';
export interface DeckSlide {
  background?: BackgroundFill;
  title: string;
  layout: SlideLayout;
  subtitle?: string;
  bullets?: string[];
  columns?: { heading: string; bullets: string[] }[];
  metrics?: { label: string; value: string; detail: string }[];
  steps?: { title: string; detail: string }[];
  chart?: { type: ChartKind; tableId: string; category: number; series: number[]; unit?: string };
  imageId?: string;
  tableId?: string;
  notes: string;
  sources: string[];
}
export interface DeckContent {
  title: string;
  slides: DeckSlide[];
}
export interface PresentationTheme {
  font: string;
  accent: string;
  ink: string;
  background: string;
  fill?: BackgroundFill;
}
export interface PageSize {
  width: number;
  height: number;
}
export interface ShapeState {
  id: string;
  name: string;
  type: string;
  left: number;
  top: number;
  width: number;
  height: number;
  text?: string;
  role?: 'title' | 'body' | 'footer';
  font?: { name: string | null; size: number | null; color: string | null; bold: boolean | null };
  alignment?: string | null;
  chart?: { kind: ChartKind; data: ChartData };
}
export interface SlideState {
  id: string;
  index: number;
  title: string;
  shapes: ShapeState[];
  background?: SlideBackground;
  effectiveBackground?: SlideBackground;
  masterBackground?: SlideBackground;
  backgroundStamp?: string;
  comments?: { id: string; text: string; shapeId?: string }[];
}
export interface PresentationSnapshot {
  scope: 'selected' | 'all';
  slides: SlideState[];
  slideIds: string[];
  selectedShapeIds: string[];
  size: PageSize | null;
  skipped: number;
}
/** Appending pages depends on the deck and its size, not the current selection. */
export interface PresentationInsertTarget {
  slideIds: string[];
  size: PageSize | null;
}
export interface ShapePatch {
  slideId: string;
  shapeId: string;
  text?: string;
  fontName?: string;
  fontSize?: number;
  fontColor?: string;
  bold?: boolean;
  alignment?: 'Left' | 'Center' | 'Right';
  left?: number;
  top?: number;
  width?: number;
  height?: number;
}
export interface EditProposal {
  summary: string;
  changes: ShapePatch[];
  actions?: PageEdit[];
  /** A permutation of the scoped slide IDs. Unselected slide positions stay fixed. */
  slideOrder?: string[];
}
export interface PresentationAccess {
  capabilities: {
    read: boolean;
    selected: boolean;
    reorder: boolean;
    insert: boolean;
    create: boolean;
    thumbnail: boolean;
    backup: boolean;
    backgrounds?: boolean;
    gradient?: boolean;
    pictureBackground?: boolean;
    structure?: boolean;
    charts?: boolean;
  };
  read: (scope: 'selected' | 'all') => Promise<PresentationSnapshot>;
  apply: (
    snapshot: PresentationSnapshot,
    proposal: EditProposal,
    images?: Record<string, string>,
  ) => Promise<void>;
  focus?: (slideId: string, shapeId?: string) => Promise<void>;
  prepareInsert: () => Promise<PresentationInsertTarget>;
  insert: (
    base64: string,
    count: number,
    target: PresentationInsertTarget,
    afterSlideId?: string,
  ) => Promise<void>;
  create: (base64: string) => Promise<void>;
  thumbnail: (id: string) => Promise<string>;
  backup: () => Promise<Blob>;
}
