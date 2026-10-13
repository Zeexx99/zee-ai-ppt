export type MessageRole = 'user' | 'assistant' | 'system';

export interface Attachment {
  id: string;
  name: string;
  type: 'image' | 'pdf';
  mimeType: string;
  base64Data: string; // raw base64 without prefix for Gemini inlineData
  size: number;
}

export type StepStatus = 'pending' | 'approving' | 'running' | 'success' | 'failed';

export interface AgentStep {
  id: string;
  stepNumber: number;
  toolName: string;
  description: string;
  status: StepStatus;
  params?: Record<string, unknown>;
  result?: string;
  error?: string;
  timestamp: number;
}

export interface ToolApproval {
  stepId: string;
  toolName: string;
  description: string;
  params: Record<string, unknown>;
  isPending: boolean;
  approved?: boolean;
}

export interface ChatMessage {
  id: string;
  role: MessageRole;
  content: string;
  timestamp: number;
  attachments?: Attachment[];
  agentSteps?: AgentStep[];
  toolApproval?: ToolApproval;
  isStreaming?: boolean;
  error?: string;
  tokenCount?: number;
  isBilingual?: boolean;
}

export interface UndoRecord {
  id: string;
  stepId: string;
  description: string;
  timestamp: number;
  actionType: string;
  snapshotData: unknown;
}

export interface SavedPrompt {
  id: string;
  title: string;
  prompt: string;
  category: 'generate' | 'edit' | 'notes' | 'petroleum' | 'arabic';
}
