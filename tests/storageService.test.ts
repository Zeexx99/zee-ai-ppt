import { describe, it, expect, beforeEach } from 'vitest';
import { storageService } from '../src/services/storageService';
import { DEFAULT_SETTINGS } from '../src/types/settings';
import { ChatMessage } from '../src/types/chat';

describe('storageService', () => {
  beforeEach(() => {
    localStorage.clear();
  });

  it('retrieves default settings when nothing is stored', () => {
    const s = storageService.getSettings();
    expect(s.model).toBe('gemini-2.5-flash');
    expect(s.protectedTerms).toContain('ESP');
    expect(s.protectedTerms).toContain('TDH');
  });

  it('saves and retrieves updated settings using zee-ppt: prefix', () => {
    const custom = {
      ...DEFAULT_SETTINGS,
      geminiApiKey: 'test-api-key-12345',
      autoRtl: false
    };

    storageService.saveSettings(custom);
    expect(localStorage.getItem('zee-ppt:settings')).toContain('test-api-key-12345');

    const loaded = storageService.getSettings();
    expect(loaded.geminiApiKey).toBe('test-api-key-12345');
    expect(loaded.autoRtl).toBe(false);
  });

  it('retrieves default saved prompts', () => {
    const prompts = storageService.getSavedPrompts();
    expect(prompts.length).toBeGreaterThanOrEqual(4);
    expect(prompts[0].title).toBe('Deck from Outline');
  });

  it('saves and clears chat history', () => {
    const msgs: ChatMessage[] = [
      { id: '1', role: 'user', content: 'Hello', timestamp: 1000 },
      { id: '2', role: 'assistant', content: 'Hi there', timestamp: 1001 }
    ];

    storageService.saveChatHistory(msgs, 'file1');
    const loaded = storageService.getChatHistory('file1');
    expect(loaded).toHaveLength(2);
    expect(loaded[0].content).toBe('Hello');

    storageService.clearChatHistory('file1');
    expect(storageService.getChatHistory('file1')).toHaveLength(0);
  });
});
