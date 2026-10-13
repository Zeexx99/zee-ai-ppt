import { AppSettings, DEFAULT_SETTINGS } from '../types/settings';
import { ChatMessage, SavedPrompt } from '../types/chat';

const PREFIX = 'zee-ppt:';
const SETTINGS_KEY = `${PREFIX}settings`;
const PROMPTS_KEY = `${PREFIX}saved-prompts`;
const HISTORY_PREFIX = `${PREFIX}history:`;

export const DEFAULT_SAVED_PROMPTS: SavedPrompt[] = [
  {
    id: 'p-1',
    title: 'Deck from Outline',
    prompt: 'Transform this outline into a structured 5-slide presentation with clear titles, 3-4 concise bullet points per slide, and speaker notes.',
    category: 'generate'
  },
  {
    id: 'p-2',
    title: 'Speaker Notes Generator',
    prompt: 'Analyze each bullet point on this slide and write natural, professional speaker notes that elaborate on key takeaways without reading the slide verbatim.',
    category: 'notes'
  },
  {
    id: 'p-3',
    title: 'Consistency Audit',
    prompt: 'Audit all slides for consistency: verify font styles, title capitalization, punctuation in bullet lists, and proper capitalization of protected technical terms (ESP, TDH, PIP, GOR, BHP).',
    category: 'edit'
  },
  {
    id: 'p-4',
    title: 'Shorten & Fit Text',
    prompt: 'Condense the text on the active slide to fit comfortably without shrinking font size. Remove fluff while retaining all technical parameters and quantitative metrics.',
    category: 'edit'
  },
  {
    id: 'p-5',
    title: 'ESP Analysis to Results Slide',
    prompt: 'Format the attached ESP pump data (frequency, PIP, PDP, TDH, flow rate, motor temp) into an executive summary slide with a KPI table and bulleted observations.',
    category: 'petroleum'
  },
  {
    id: 'p-6',
    title: 'Bilingual Technical Summary',
    prompt: 'Create a bilingual slide summary (English and Arabic) with RTL formatting for Arabic paragraphs and preserving all engineering units and acronyms.',
    category: 'arabic'
  }
];

export const storageService = {
  getSettings(): AppSettings {
    try {
      const raw = localStorage.getItem(SETTINGS_KEY);
      if (!raw) return DEFAULT_SETTINGS;
      const parsed = JSON.parse(raw);
      return { ...DEFAULT_SETTINGS, ...parsed };
    } catch (e) {
      console.warn('Failed to load settings from storage, using defaults:', e);
      return DEFAULT_SETTINGS;
    }
  },

  saveSettings(settings: AppSettings): void {
    try {
      localStorage.setItem(SETTINGS_KEY, JSON.stringify(settings));
    } catch (e) {
      console.error('Failed to save settings:', e);
    }
  },

  getSavedPrompts(): SavedPrompt[] {
    try {
      const raw = localStorage.getItem(PROMPTS_KEY);
      if (!raw) return DEFAULT_SAVED_PROMPTS;
      const parsed = JSON.parse(raw);
      return Array.isArray(parsed) ? parsed : DEFAULT_SAVED_PROMPTS;
    } catch {
      return DEFAULT_SAVED_PROMPTS;
    }
  },

  saveSavedPrompts(prompts: SavedPrompt[]): void {
    try {
      localStorage.setItem(PROMPTS_KEY, JSON.stringify(prompts));
    } catch (e) {
      console.error('Failed to save prompts:', e);
    }
  },

  getChatHistory(fileKey = 'global'): ChatMessage[] {
    try {
      const raw = localStorage.getItem(`${HISTORY_PREFIX}${fileKey}`);
      if (!raw) return [];
      const parsed = JSON.parse(raw);
      return Array.isArray(parsed) ? parsed : [];
    } catch {
      return [];
    }
  },

  /**
   * Saves chat history with automatic oldest-message trimming if quota is approached
   */
  saveChatHistory(messages: ChatMessage[], fileKey = 'global'): void {
    const key = `${HISTORY_PREFIX}${fileKey}`;
    let list = [...messages];

    while (list.length > 0) {
      try {
        localStorage.setItem(key, JSON.stringify(list));
        return; // Success
      } catch (e) {
        // QuotaExceededError: Trim oldest 2 messages and retry
        console.warn('Storage quota reached, trimming oldest chat history...');
        if (list.length <= 2) {
          localStorage.removeItem(key);
          return;
        }
        list = list.slice(2);
      }
    }
  },

  clearChatHistory(fileKey = 'global'): void {
    try {
      localStorage.removeItem(`${HISTORY_PREFIX}${fileKey}`);
    } catch (e) {
      console.error('Failed to clear chat history:', e);
    }
  }
};
