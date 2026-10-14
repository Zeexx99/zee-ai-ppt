import { ZeeSettings } from "../types/settings";
import { ChatAttachment, ChatMessage, GroundingCitation } from "../types/chat";
import { buildGeminiEndpoint } from "./authService";
import { buildDirectionAndLanguageInstruction } from "./arabicService";

export interface GeminiFunctionDeclaration {
  name: string;
  description: string;
  parameters: {
    type: "OBJECT";
    properties: Record<
      string,
      {
        type: "STRING" | "NUMBER" | "INTEGER" | "BOOLEAN" | "ARRAY" | "OBJECT";
        description: string;
        items?: { type: string };
        enum?: string[];
      }
    >;
    required?: string[];
  };
}

export interface GeminiFunctionCall {
  name: string;
  args: Record<string, unknown>;
}

export interface GeminiPart {
  text?: string;
  inlineData?: {
    mimeType: string;
    data: string;
  };
  functionCall?: GeminiFunctionCall;
  functionResponse?: {
    name: string;
    response: Record<string, unknown>;
  };
  executableCode?: {
    language: string;
    code: string;
  };
  codeExecutionResult?: {
    outcome: string;
    output: string;
  };
}

export interface GeminiTurnContent {
  role: "user" | "model";
  parts: GeminiPart[];
}

export interface CostGuardEvaluation {
  estimatedTokens: number;
  limit: number;
  exceedsLimit: boolean;
}

/**
 * Estimates prompt token count for text, conversation history, and attachments
 * (images ~258 tokens each, PDF pages/bytes ~1 token per 12 bytes of base64).
 */
export function estimatePromptTokens(
  prompt: string,
  history: ChatMessage[] = [],
  attachments: ChatAttachment[] = [],
  systemContext: string = ""
): number {
  let chars = (prompt?.length || 0) + (systemContext?.length || 0);
  for (const msg of history.slice(-12)) {
    chars += msg.content?.length || 0;
  }

  let tokens = Math.ceil(chars / 3.8);

  for (const att of attachments) {
    if (att.kind === "image") {
      tokens += 260;
    } else if (att.kind === "pdf") {
      tokens += Math.max(300, Math.ceil((att.sizeBytes || att.base64Data.length) / 14));
    } else if (att.textContent) {
      tokens += Math.ceil(att.textContent.length / 3.8);
    }
  }

  return tokens;
}

export function evaluateCostGuard(
  settings: ZeeSettings,
  prompt: string,
  history: ChatMessage[] = [],
  attachments: ChatAttachment[] = [],
  systemContext: string = ""
): CostGuardEvaluation {
  const estimatedTokens = estimatePromptTokens(prompt, history, attachments, systemContext);
  const limit = Math.max(500, settings.costGuardTokenLimit || 12000);
  return {
    estimatedTokens,
    limit,
    exceedsLimit: Boolean(settings.costGuardEnabled && estimatedTokens > limit),
  };
}

export function buildSystemInstruction(
  settings: ZeeSettings,
  presentationContextSummary?: string
): string {
  const lines: string[] = [
    "You are Zee AI for PowerPoint, an expert AI presentation assistant, slide designer, and petroleum/reservoir engineering analyst embedded inside Microsoft PowerPoint.",
    "Always write crisp, executive-ready slide content with clear structure and accurate engineering figures.",
    buildDirectionAndLanguageInstruction(settings),
  ];

  if (settings.companyTemplate?.enabled) {
    const tpl = settings.companyTemplate;
    lines.push(
      `COMPANY TEMPLATE MODE IS ACTIVE: Use layout "${tpl.layoutName}", heading font "${tpl.headingFont}", body font "${tpl.bodyFont}", primary color ${tpl.primaryColor}, accent color ${tpl.accentColor}, and footer "${tpl.footerText}".`
    );
  }

  if (presentationContextSummary) {
    lines.push("CURRENT PRESENTATION CONTEXT (compact summary):\n" + presentationContextSummary);
  }

  return lines.filter(Boolean).join("\n\n");
}

export function convertAttachmentsToGeminiParts(attachments: ChatAttachment[] = []): GeminiPart[] {
  const parts: GeminiPart[] = [];
  for (const att of attachments) {
    if (att.kind === "text" && att.textContent) {
      parts.push({
        text: `[Attached File: ${att.name}]\n${att.textContent}`,
      });
    } else if (att.base64Data) {
      parts.push({
        inlineData: {
          mimeType: att.mimeType || (att.kind === "pdf" ? "application/pdf" : "image/png"),
          data: att.base64Data,
        },
      });
    }
  }
  return parts;
}

export function buildGeminiContents(
  history: ChatMessage[],
  userPrompt: string,
  attachments: ChatAttachment[] = []
): GeminiTurnContent[] {
  const contents: GeminiTurnContent[] = [];

  for (const msg of history.slice(-14)) {
    if (!msg.content?.trim() || msg.role === "system") continue;
    contents.push({
      role: msg.role === "assistant" ? "model" : "user",
      parts: [{ text: msg.content }],
    });
  }

  const userParts: GeminiPart[] = [
    ...convertAttachmentsToGeminiParts(attachments),
    { text: userPrompt },
  ];

  contents.push({
    role: "user",
    parts: userParts,
  });

  return contents;
}

export function buildToolsPayload(
  settings: ZeeSettings,
  functionDeclarations?: GeminiFunctionDeclaration[]
): Record<string, unknown>[] | undefined {
  // When Agent function declarations are provided, prioritize function calling
  if (functionDeclarations && functionDeclarations.length > 0) {
    return [{ functionDeclarations }];
  }

  const tools: Record<string, unknown>[] = [];
  if (settings.webSearchEnabled) {
    tools.push({ google_search: {} });
  }
  if (settings.codeExecutionEnabled) {
    tools.push({ code_execution: {} });
  }
  return tools.length > 0 ? tools : undefined;
}

export interface GeminiGenerationResult {
  text: string;
  functionCalls: GeminiFunctionCall[];
  citations: GroundingCitation[];
  codeExecutionOutput?: string;
  rawModelParts: GeminiPart[];
}

export function extractResultFromCandidate(candidate: Record<string, any>): GeminiGenerationResult {
  const parts: GeminiPart[] = candidate?.content?.parts || [];
  let text = "";
  const functionCalls: GeminiFunctionCall[] = [];
  const codeBlocks: string[] = [];

  for (const part of parts) {
    if (typeof part.text === "string") {
      text += part.text;
    }
    if (part.functionCall && part.functionCall.name) {
      functionCalls.push({
        name: part.functionCall.name,
        args: (part.functionCall.args as Record<string, unknown>) || {},
      });
    }
    if (part.executableCode?.code) {
      codeBlocks.push(
        `\`\`\`${(part.executableCode.language || "python").toLowerCase()}\n${part.executableCode.code}\n\`\`\``
      );
    }
    if (part.codeExecutionResult?.output) {
      codeBlocks.push(`**Execution Output:**\n\`\`\`text\n${part.codeExecutionResult.output}\n\`\`\``);
    }
  }

  const citations: GroundingCitation[] = [];
  const chunks = candidate?.groundingMetadata?.groundingChunks;
  if (Array.isArray(chunks)) {
    for (const chunk of chunks) {
      if (chunk?.web?.uri) {
        citations.push({
          title: chunk.web.title || chunk.web.uri,
          uri: chunk.web.uri,
        });
      }
    }
  }

  return {
    text,
    functionCalls,
    citations,
    codeExecutionOutput: codeBlocks.length > 0 ? codeBlocks.join("\n\n") : undefined,
    rawModelParts: parts,
  };
}

export async function countTokensWithGemini(
  settings: ZeeSettings,
  contents: GeminiTurnContent[]
): Promise<number> {
  const { url, headers } = buildGeminiEndpoint(settings, "countTokens", false);
  const res = await fetch(url, {
    method: "POST",
    headers,
    body: JSON.stringify({ contents }),
  });
  if (!res.ok) {
    const errText = await res.text();
    throw new Error(`countTokens failed (${res.status}): ${errText}`);
  }
  const data = await res.json();
  return Number(data.totalTokens) || 0;
}

export async function generateGeminiTurn(
  settings: ZeeSettings,
  contents: GeminiTurnContent[],
  options: {
    systemInstruction?: string;
    functionDeclarations?: GeminiFunctionDeclaration[];
    signal?: AbortSignal;
  } = {}
): Promise<GeminiGenerationResult> {
  const { url, headers } = buildGeminiEndpoint(settings, "generateContent", false);
  const tools = buildToolsPayload(settings, options.functionDeclarations);

  const body: Record<string, unknown> = {
    contents,
    generationConfig: {
      temperature: settings.temperature ?? 0.3,
    },
  };

  if (options.systemInstruction) {
    body.systemInstruction = {
      parts: [{ text: options.systemInstruction }],
    };
  }
  if (tools) {
    body.tools = tools;
  }

  const res = await fetch(url, {
    method: "POST",
    headers,
    body: JSON.stringify(body),
    signal: options.signal,
  });

  if (!res.ok) {
    const errBody = await res.text();
    throw new Error(`Gemini API Error (${res.status}): ${errBody}`);
  }

  const json = await res.json();
  const candidate = json?.candidates?.[0] || {};
  return extractResultFromCandidate(candidate);
}

/**
 * Streams a chat response from Gemini using SSE (?alt=sse) with automatic
 * fallback to standard JSON parsing if the proxy returns a non-SSE response.
 */
export async function streamGeminiChat(
  settings: ZeeSettings,
  history: ChatMessage[],
  userPrompt: string,
  options: {
    attachments?: ChatAttachment[];
    presentationContextSummary?: string;
    onChunk: (accumulatedText: string, citations: GroundingCitation[], codeOutput?: string) => void;
    signal?: AbortSignal;
  }
): Promise<GeminiGenerationResult> {
  const { url, headers } = buildGeminiEndpoint(settings, "streamGenerateContent", true);
  const contents = buildGeminiContents(history, userPrompt, options.attachments);
  const systemInstruction = buildSystemInstruction(settings, options.presentationContextSummary);
  const tools = buildToolsPayload(settings);

  const body: Record<string, unknown> = {
    contents,
    systemInstruction: {
      parts: [{ text: systemInstruction }],
    },
    generationConfig: {
      temperature: settings.temperature ?? 0.3,
    },
  };

  if (tools) {
    body.tools = tools;
  }

  const res = await fetch(url, {
    method: "POST",
    headers,
    body: JSON.stringify(body),
    signal: options.signal,
  });

  if (!res.ok) {
    const errBody = await res.text();
    throw new Error(`Gemini Streaming Error (${res.status}): ${errBody}`);
  }

  if (!res.body) {
    const fallbackJson = await res.json();
    const result = extractResultFromCandidate(fallbackJson?.candidates?.[0] || {});
    options.onChunk(result.text, result.citations, result.codeExecutionOutput);
    return result;
  }

  const reader = res.body.getReader();
  const decoder = new TextDecoder("utf-8");

  let buffer = "";
  let accumulatedText = "";
  let allCitations: GroundingCitation[] = [];
  let accumulatedCodeOutput = "";
  const allParts: GeminiPart[] = [];

  const processDataObject = (parsed: any) => {
    const candidate = parsed?.candidates?.[0];
    if (!candidate) return;
    const partial = extractResultFromCandidate(candidate);
    if (partial.text) {
      accumulatedText += partial.text;
    }
    if (partial.citations.length > 0) {
      const seen = new Set(allCitations.map((c) => c.uri));
      for (const c of partial.citations) {
        if (!seen.has(c.uri)) {
          allCitations.push(c);
          seen.add(c.uri);
        }
      }
    }
    if (partial.codeExecutionOutput) {
      accumulatedCodeOutput = accumulatedCodeOutput
        ? `${accumulatedCodeOutput}\n\n${partial.codeExecutionOutput}`
        : partial.codeExecutionOutput;
    }
    allParts.push(...partial.rawModelParts);
    options.onChunk(
      accumulatedText,
      allCitations,
      accumulatedCodeOutput || undefined
    );
  };

  while (true) {
    const { value, done } = await reader.read();
    if (done) break;
    buffer += decoder.decode(value, { stream: true });

    const lines = buffer.split(/\r?\n/);
    buffer = lines.pop() || "";

    for (const line of lines) {
      const trimmed = line.trim();
      if (!trimmed.startsWith("data:")) continue;
      const jsonStr = trimmed.slice(5).trim();
      if (!jsonStr || jsonStr === "[DONE]") continue;
      try {
        const parsed = JSON.parse(jsonStr);
        processDataObject(parsed);
      } catch {
        // Ignore incomplete JSON line
      }
    }
  }

  if (buffer.trim().startsWith("data:")) {
    const jsonStr = buffer.trim().slice(5).trim();
    if (jsonStr && jsonStr !== "[DONE]") {
      try {
        processDataObject(JSON.parse(jsonStr));
      } catch {
        // Ignore
      }
    }
  } else if (!accumulatedText && buffer.trim().startsWith("{")) {
    try {
      processDataObject(JSON.parse(buffer.trim()));
    } catch {
      // Ignore
    }
  }

  return {
    text: accumulatedText,
    functionCalls: [],
    citations: allCitations,
    codeExecutionOutput: accumulatedCodeOutput || undefined,
    rawModelParts: allParts,
  };
}
