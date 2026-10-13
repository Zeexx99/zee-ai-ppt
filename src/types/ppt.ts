export interface SlideSummary {
  id: string;
  index: number; // 1-based display index
  title: string;
  contentSnippet: string;
  shapeCount: number;
  notes?: string;
  isSelected?: boolean;
}

export interface PresentationOverview {
  slideCount: number;
  activeSlideIndex: number;
  slides: SlideSummary[];
  selectedText?: string;
}

export interface RequirementSetStatus {
  hasApi11: boolean;
  hasApi12: boolean;
  hasApi13: boolean;
  hasApi14: boolean;
  hasApi15: boolean;
  canFormatShapes: boolean;
  canAddTables: boolean;
  canManageNotes: boolean;
  canInsertSlidesBase64: boolean;
  highestVersion: string;
}

export interface TextFormatOptions {
  fontName?: string;
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  color?: string; // hex like #D24726
  align?: 'Left' | 'Center' | 'Right' | 'Justify';
  isRtl?: boolean;
}

export interface TableInsertOptions {
  rows: number;
  columns: number;
  headers?: string[];
  data?: string[][];
  title?: string;
}
