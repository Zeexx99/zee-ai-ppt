export interface PowerPointApiRequirements {
  isPowerPointHost: boolean;
  pptApi1_1: boolean; // Basic slide creation
  pptApi1_2: boolean; // insertSlidesFromBase64, delete slide
  pptApi1_3: boolean; // Slide Masters, Custom Layouts, add slide with options
  pptApi1_4: boolean; // Shapes, TextFrame, TextRange, formatting, geometric shapes, lines
  pptApi1_5: boolean; // Selection of slides, shapes, text ranges
  pptApi1_6: boolean; // Hyperlinks, slide reordering/move, background
  pptApi1_7: boolean; // Speaker notes (Slide.notesSlide), document properties
  pptApi1_8: boolean; // Tables API / advanced shape features
  selectedTextCommonApi: boolean; // Office.context.document.getSelectedDataAsync / setSelectedDataAsync
  imageCoercionSupported: boolean; // Office.CoercionType.Image
  tableCoercionSupported: boolean; // Office.CoercionType.Table
}

export interface ShapeInfo {
  id: string;
  name: string;
  type: string;
  text: string;
  left: number;
  top: number;
  width: number;
  height: number;
  fontName?: string;
  fontSize?: number;
  fontColor?: string;
  isTitle?: boolean;
}

export interface SlideSummary {
  index: number; // 1-based slide number
  id: string;
  title: string;
  textPreview: string;
  shapeCount: number;
  layoutName?: string;
  hasNotes?: boolean;
}

export interface SlideDetail {
  index: number; // 1-based slide number
  id: string;
  title: string;
  layoutName?: string;
  shapes: ShapeInfo[];
  speakerNotes: string;
  fullText: string;
}

export interface PresentationSummary {
  documentId: string;
  slideCount: number;
  activeSlideIndex: number;
  slides: SlideSummary[];
  activeSlideDetail?: SlideDetail;
  selectedText?: string;
  availableLayouts: string[];
}

export interface DeckOutlineSlide {
  slideNumber: number;
  title: string;
  titleArabic?: string;
  layout: "Title Slide" | "Title and Content" | "Two Content" | "Section Header" | "Summary Table";
  bullets: string[];
  bulletsArabic?: string[];
  table?: {
    headers: string[];
    rows: string[][];
  };
  speakerNotes: string;
}

export interface DeckOutlinePlan {
  deckTitle: string;
  deckSubtitle?: string;
  sourceSummary: string;
  slides: DeckOutlineSlide[];
}

export type PetroleumResultCategory = "decline_curve" | "esp_analysis" | "production_qc" | "custom";

export interface ExcelResultsPayload {
  category: PetroleumResultCategory;
  wellOrFieldName: string;
  rawInputText: string;
  slideTitle: string;
  kpis: { label: string; value: string; status?: "normal" | "warning" | "critical" }[];
  headers: string[];
  rows: string[][];
  takeaways: string[];
  speakerNotes: string;
  chartImageBase64?: string;
}

export interface ConsistencyIssue {
  id: string;
  slideIndex: number;
  slideId: string;
  shapeId?: string;
  category: "font_family" | "font_size" | "title_case" | "units" | "protected_term";
  severity: "warning" | "info";
  description: string;
  currentValue: string;
  suggestedValue: string;
  selected: boolean;
}

export interface TextFitIssue {
  id: string;
  slideIndex: number;
  slideId: string;
  shapeId: string;
  shapeName: string;
  originalText: string;
  lineCount: number;
  wordCount: number;
  boxHeightPt: number;
  estimatedOverflowRatio: number;
  shortenedText?: string;
  recommendedFontSize?: number;
  selected: boolean;
}

export interface WellBookRow {
  wellName: string;
  metrics: Record<string, string>;
  status?: string;
  action?: string;
}

export interface CrossSlideContradiction {
  id: string;
  entity: string; // e.g. "WELL-04" or "Liquid Rate"
  unit: string;
  firstSlideIndex: number;
  firstValue: string;
  secondSlideIndex: number;
  secondValue: string;
  description: string;
}

export interface DeckQnaItem {
  id: string;
  question: string;
  recommendedAnswer: string;
  backupSlideTitle: string;
  backupSlideBullets: string[];
  selected: boolean;
}

export interface DeckReviewReport {
  overallScore: number;
  executiveSummary: string;
  contradictions: CrossSlideContradiction[];
  slideCritiques: { slideIndex: number; title: string; feedback: string }[];
  toughQuestions: DeckQnaItem[];
}

export type SlideBeautifyStyle = "two_column" | "three_cards" | "align_grid";

export type EngineeringDiagramKind =
  | "wellbore_schematic"
  | "process_flowchart"
  | "workover_timeline";

export interface AgendaSectionItem {
  id: string;
  sectionNumber: number;
  title: string;
  subtitle: string;
  startSlideIndex: number;
  selected: boolean;
}

export interface ActionMatrixItem {
  id: string;
  assetOrWell: string;
  action: string;
  owner: string;
  targetDate: string;
  priority: "High" | "Medium" | "Low";
  status: string;
  selected: boolean;
}

export interface AgendaAndActionsPlan {
  presentationTitle: string;
  sections: AgendaSectionItem[];
  actions: ActionMatrixItem[];
}

