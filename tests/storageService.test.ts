import { beforeEach, describe, expect, it } from "vitest";
import {
  getStorageKey,
  loadChatHistory,
  loadSavedPrompts,
  loadSettings,
  normalizeDocHistoryKey,
  saveChatHistory,
  saveCustomPrompts,
  saveSettings,
  trimChatHistory,
} from "../src/services/storageService";
import { DEFAULT_SETTINGS, STORAGE_PREFIX } from "../src/types/settings";
import { ChatMessage } from "../src/types/chat";

function setupMockLocalStorage() {
  const store = new Map<string, string>();
  const mock: Storage = {
    get length() {
      return store.size;
    },
    clear() {
      store.clear();
    },
    getItem(key: string) {
      return store.has(key) ? store.get(key)! : null;
    },
    key(index: number) {
      return Array.from(store.keys())[index] ?? null;
    },
    removeItem(key: string) {
      store.delete(key);
    },
    setItem(key: string, value: string) {
      store.set(key, String(value));
    },
  };
  (globalThis as any).localStorage = mock;
}

describe("storageService", () => {
  beforeEach(() => {
    setupMockLocalStorage();
  });

  it("enforces the zee-ppt: storage key prefix so nothing collides with Excel/Word", () => {
    expect(STORAGE_PREFIX).toBe("zee-ppt:");
    expect(getStorageKey("settings")).toBe("zee-ppt:settings");
    expect(normalizeDocHistoryKey("North Field Q3 Review.pptx")).toMatch(
      /^zee-ppt:history:/
    );
  });

  it("saves and loads settings with company template defaults", () => {
    const initial = loadSettings();
    expect(initial.model).toBe(DEFAULT_SETTINGS.model);

    saveSettings({
      ...initial,
      geminiApiKey: "test-key-123",
      bilingualMode: true,
      arabicFont: "Sakkal Majalla",
    });

    const loaded = loadSettings();
    expect(loaded.geminiApiKey).toBe("test-key-123");
    expect(loaded.bilingualMode).toBe(true);
    expect(loaded.arabicFont).toBe("Sakkal Majalla");
  });

  it("trims oldest chat history messages when count or byte size exceeds limit instead of wiping all", () => {
    const messages: ChatMessage[] = Array.from({ length: 15 }, (_, idx) => ({
      id: `msg-${idx + 1}`,
      role: idx % 2 === 0 ? "user" : "assistant",
      content: `Message #${idx + 1} — ${"X".repeat(200)}`,
      timestamp: Date.now() + idx,
    }));

    const { messages: trimmedByCount, trimmedCount } = trimChatHistory(
      messages,
      10,
      100_000
    );
    expect(trimmedByCount.length).toBe(10);
    expect(trimmedCount).toBe(5);
    expect(trimmedByCount[0].id).toBe("msg-6");
    expect(trimmedByCount[9].id).toBe("msg-15");

    // Byte size trimming keeps newest messages
    const { messages: trimmedByBytes } = trimChatHistory(messages, 50, 1200);
    expect(trimmedByBytes.length).toBeGreaterThan(0);
    expect(trimmedByBytes.length).toBeLessThan(15);
    expect(trimmedByBytes[trimmedByBytes.length - 1].id).toBe("msg-15");

    saveChatHistory(messages, "deck-1.pptx");
    const restored = loadChatHistory("deck-1.pptx");
    expect(restored.length).toBe(15);
  });

  it("persists custom saved prompts alongside built-in prompts", () => {
    const initial = loadSavedPrompts();
    expect(initial.some((p) => p.isBuiltIn)).toBe(true);

    saveCustomPrompts([
      ...initial,
      {
        id: "custom-1",
        title: "Custom Well Prompt",
        prompt: "Summarize WELL-09 test",
        category: "custom",
      },
    ]);

    const after = loadSavedPrompts();
    expect(after.some((p) => p.id === "custom-1")).toBe(true);
  });
});
