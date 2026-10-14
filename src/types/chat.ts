export type AttachmentKind = "image" | "pdf" | "text";

export interface ChatAttachment {
  id: string;
  name: string;
  mimeType: string;
  kind: AttachmentKind;
  base64Data: string; // raw base64 without data: prefix
  sizeBytes: number;
  textContent?: string;
}

export type AgentStepStatus =
  | "pending_approval"
  | "approved"
  | "running"
  | "completed"
  | "failed"
  | "rejected"
  | "undone";

export interface SlideUndoSnapshot {
  stepId: string;
  toolName: string;
  timestamp: number;
  slideIndex: number;
  slideId?: string;
  actionType: "modify" | "add_slide" | "delete_slide" | "reorder_slide";
  previousTitle?: string;
  previousBodyTexts?: { shapeId: string; name: string; text: string }[];
  previousNotes?: string;
  createdSlideId?: string;
  deletedSlideState?: {
    index: number;
    title: string;
    bodyTexts: string[];
    notes: string;
    layoutName?: string;
  };
  reorderFromIndex?: number;
  reorderToIndex?: number;
}

export interface AgentToolStep {
  id: string;
  stepNumber: number;
  toolName: string;
  args: Record<string, unknown>;
  description: string;
  requiresApproval: boolean;
  status: AgentStepStatus;
  resultSummary?: string;
  errorMessage?: string;
  retryCount: number;
  undoSnapshot?: SlideUndoSnapshot;
}

export interface GroundingCitation {
  title: string;
  uri: string;
}

export interface ChatMessage {
  id: string;
  role: "user" | "assistant" | "system";
  content: string;
  timestamp: number;
  isStreaming?: boolean;
  attachments?: ChatAttachment[];
  agentSteps?: AgentToolStep[];
  awaitingContinue?: boolean;
  citations?: GroundingCitation[];
  codeExecutionOutput?: string;
  estimatedTokens?: number;
  direction?: "rtl" | "ltr";
}

export interface SavedPrompt {
  id: string;
  title: string;
  prompt: string;
  category: "slide" | "arabic" | "petroleum" | "custom";
  isBuiltIn?: boolean;
}

export type QuickActionType =
  | "rewrite"
  | "translate_ar"
  | "translate_en"
  | "translate_bilingual"
  | "summarize"
  | "shorten_fit"
  | "speaker_notes"
  | "convert_units";
