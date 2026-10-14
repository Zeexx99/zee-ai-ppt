import {
  APP_VERSION,
  BUILD_DATE,
  DEFAULT_SETTINGS,
  STORAGE_PREFIX,
  ZeeSettings,
} from "../types/settings";
import { ChatMessage, SavedPrompt } from "../types/chat";

export const SETTINGS_STORAGE_KEY = `${STORAGE_PREFIX}settings`;
export const PROMPTS_STORAGE_KEY = `${STORAGE_PREFIX}saved-prompts`;
export const HISTORY_INDEX_KEY = `${STORAGE_PREFIX}history-index`;
export const BUILD_VERSION_KEY = `${STORAGE_PREFIX}build-version`;
export const RELOAD_GUARD_KEY = `${STORAGE_PREFIX}reloaded-for-version`;

const MAX_MESSAGES_PER_DOC = 60;
const MAX_HISTORY_BYTES_PER_DOC = 180_000; // ~180 KB per deck

export const BUILT_IN_PROMPTS: SavedPrompt[] = [
  {
    id: "builtin-exec-summary",
    title: "Executive Summary Slide",
    prompt:
      "Summarize the current slide or selected text into 4-5 concise executive bullet points with clear action-oriented takeaways.",
    category: "slide",
    isBuiltIn: true,
  },
  {
    id: "builtin-speaker-notes",
    title: "Generate Speaker Notes",
    prompt:
      "Write natural, well-paced speaker notes (60-90 seconds) for the active slide, highlighting key numbers and transitions.",
    category: "slide",
    isBuiltIn: true,
  },
  {
    id: "builtin-bilingual-slide",
    title: "Bilingual English + Arabic Slide",
    prompt:
      "Rewrite this slide content in Bilingual mode (English followed by professional Arabic translation), keeping all protected engineering terms (ESP, TDH, PIP, GOR, BHP, bbl/d, well names) unchanged.",
    category: "arabic",
    isBuiltIn: true,
  },
  {
    id: "builtin-esp-review",
    title: "ESP Performance & Diagnosis Slide",
    prompt:
      "Build a structured ESP Well Performance & Troubleshooting slide including PIP, TDH, BHP, GOR, pump intake conditions, and recommended workover/operating actions.",
    category: "petroleum",
    isBuiltIn: true,
  },
  {
    id: "builtin-dca-slide",
    title: "Decline Curve Analysis (Arps) Summary",
    prompt:
      "Create a Decline Curve Analysis (DCA) summary slide showing Qi, Di, b-factor, EUR, remaining reserves (bbl / MSTB), and forecast assumptions.",
    category: "petroleum",
    isBuiltIn: true,
  },
  {
    id: "builtin-prod-qc",
    title: "Production QC & Well Test Slide",
    prompt:
      "Generate a Production QC slide summarizing liquid rate (bbl/d), oil rate (stb/d), water cut (WC%), GOR (scf/stb), WHP/BHP anomalies, and data quality flags.",
    category: "petroleum",
    isBuiltIn: true,
  },
];

export function getStorageKey(suffix: string): string {
  if (suffix.startsWith(STORAGE_PREFIX)) {
    return suffix;
  }
  return `${STORAGE_PREFIX}${suffix}`;
}

export function loadSettings(): ZeeSettings {
  try {
    const raw = localStorage.getItem(SETTINGS_STORAGE_KEY);
    if (!raw) {
      return {
        ...DEFAULT_SETTINGS,
        companyTemplate: { ...DEFAULT_SETTINGS.companyTemplate },
      };
    }
    const parsed = JSON.parse(raw) as Partial<ZeeSettings>;
    const savedModel = (parsed.model || "").trim();
    const isRemovedLegacyModel =
      !savedModel ||
      savedModel.startsWith("gemini-1.") ||
      savedModel.startsWith("gemini-2.0") ||
      savedModel.startsWith("gemini-2.5");

    return {
      ...DEFAULT_SETTINGS,
      ...parsed,
      model: isRemovedLegacyModel ? DEFAULT_SETTINGS.model : savedModel,
      customProtectedTerms: Array.isArray(parsed.customProtectedTerms)
        ? parsed.customProtectedTerms
        : DEFAULT_SETTINGS.customProtectedTerms,
      companyTemplate: {
        ...DEFAULT_SETTINGS.companyTemplate,
        ...(parsed.companyTemplate || {}),
      },
    };
  } catch {
    return {
      ...DEFAULT_SETTINGS,
      companyTemplate: { ...DEFAULT_SETTINGS.companyTemplate },
    };
  }
}

export function saveSettings(settings: ZeeSettings): void {
  try {
    localStorage.setItem(SETTINGS_STORAGE_KEY, JSON.stringify(settings));
  } catch (err) {
    console.warn("Failed to save Zee AI settings:", err);
  }
}

export function normalizeDocHistoryKey(docIdentifier?: string): string {
  const cleaned = (docIdentifier || "default-presentation")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9._-]+/g, "_")
    .slice(-80);
  return `${STORAGE_PREFIX}history:${cleaned || "default-presentation"}`;
}

/**
 * Trims oldest messages (preserving pairs where possible) so the serialized
 * history fits within maxMessages and maxBytes instead of wiping all messages.
 */
export function trimChatHistory(
  messages: ChatMessage[],
  maxMessages: number = MAX_MESSAGES_PER_DOC,
  maxBytes: number = MAX_HISTORY_BYTES_PER_DOC
): { messages: ChatMessage[]; trimmedCount: number } {
  let working = messages.map((m) => ({
    ...m,
    isStreaming: false,
    // Strip heavy base64 payloads from persisted history while keeping metadata
    attachments: m.attachments?.map((att) => ({
      ...att,
      base64Data: att.base64Data.length > 4096 ? "" : att.base64Data,
    })),
  }));

  let trimmedCount = 0;

  if (working.length > maxMessages) {
    const excess = working.length - maxMessages;
    working = working.slice(excess);
    trimmedCount += excess;
  }

  while (working.length > 1 && JSON.stringify(working).length > maxBytes) {
    working = working.slice(1);
    trimmedCount += 1;
  }

  return { messages: working, trimmedCount };
}

export function loadChatHistory(docIdentifier?: string): ChatMessage[] {
  try {
    const key = normalizeDocHistoryKey(docIdentifier);
    const raw = localStorage.getItem(key);
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) return [];
    return parsed as ChatMessage[];
  } catch {
    return [];
  }
}

export function saveChatHistory(
  messages: ChatMessage[],
  docIdentifier?: string
): { savedCount: number; trimmedCount: number } {
  const key = normalizeDocHistoryKey(docIdentifier);
  const { messages: trimmed, trimmedCount } = trimChatHistory(messages);
  let current = [...trimmed];
  let extraTrimmed = 0;

  while (current.length > 0) {
    try {
      localStorage.setItem(key, JSON.stringify(current));
      return { savedCount: current.length, trimmedCount: trimmedCount + extraTrimmed };
    } catch {
      // Quota exceeded: drop oldest message and retry instead of clearing everything
      if (current.length === 1) {
        break;
      }
      current = current.slice(Math.max(1, Math.floor(current.length * 0.25)));
      extraTrimmed += 1;
    }
  }

  return { savedCount: 0, trimmedCount: messages.length };
}

export function clearChatHistory(docIdentifier?: string): void {
  try {
    const key = normalizeDocHistoryKey(docIdentifier);
    localStorage.removeItem(key);
  } catch {
    // ignore
  }
}

export function loadSavedPrompts(): SavedPrompt[] {
  try {
    const raw = localStorage.getItem(PROMPTS_STORAGE_KEY);
    if (!raw) return [...BUILT_IN_PROMPTS];
    const custom = JSON.parse(raw) as SavedPrompt[];
    if (!Array.isArray(custom)) return [...BUILT_IN_PROMPTS];
    return [...BUILT_IN_PROMPTS, ...custom.filter((p) => !p.isBuiltIn)];
  } catch {
    return [...BUILT_IN_PROMPTS];
  }
}

export function saveCustomPrompts(prompts: SavedPrompt[]): void {
  try {
    const customOnly = prompts.filter((p) => !p.isBuiltIn);
    localStorage.setItem(PROMPTS_STORAGE_KEY, JSON.stringify(customOnly));
  } catch (err) {
    console.warn("Failed to save custom prompts:", err);
  }
}

/**
 * Stale-cache detector and one-time reload guard.
 * Prevents crashes after deployment when Office WebView caches an old index.html or chunk.
 */
export function checkAndHandleStaleCache(): boolean {
  const currentStamp = `${APP_VERSION}_${BUILD_DATE}`;
  try {
    const previousStamp = localStorage.getItem(BUILD_VERSION_KEY);
    localStorage.setItem(BUILD_VERSION_KEY, currentStamp);

    if (previousStamp && previousStamp !== currentStamp) {
      const alreadyReloaded = sessionStorage.getItem(RELOAD_GUARD_KEY);
      if (alreadyReloaded !== currentStamp) {
        sessionStorage.setItem(RELOAD_GUARD_KEY, currentStamp);
        const url = new URL(window.location.href);
        url.searchParams.set("v", currentStamp);
        window.location.replace(url.toString());
        return true;
      }
    }
  } catch {
    // Ignore storage access errors in restricted frames
  }
  return false;
}
