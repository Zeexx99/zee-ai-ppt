import { ArabicFontChoice } from '../types/settings';

// Unicode ranges for Arabic, Persian, Urdu scripts
const ARABIC_REGEX = /[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]/;

export interface DirectionAnalysis {
  isRtl: boolean;
  alignment: 'Left' | 'Center' | 'Right';
  suggestedFont?: string;
  arabicRatio: number;
}

export const arabicService = {
  /**
   * Tests if string contains Arabic characters
   */
  hasArabic(text: string): boolean {
    return ARABIC_REGEX.test(text);
  },

  /**
   * Calculates the proportion of Arabic characters in text
   */
  getArabicRatio(text: string): number {
    if (!text || text.length === 0) return 0;
    const clean = text.replace(/[\s\d\p{P}]/gu, '');
    if (!clean) return 0;
    let arabicCount = 0;
    for (const char of clean) {
      if (ARABIC_REGEX.test(char)) {
        arabicCount++;
      }
    }
    return arabicCount / clean.length;
  },

  /**
   * Analyzes text and recommends direction, alignment, and font
   */
  analyzeDirection(text: string, fontSetting: ArabicFontChoice = 'Same as document', isTitle = false): DirectionAnalysis {
    const ratio = this.getArabicRatio(text);
    const isRtl = ratio >= 0.25;

    let alignment: 'Left' | 'Center' | 'Right' = 'Left';
    if (isTitle) {
      alignment = 'Center';
    } else if (isRtl) {
      alignment = 'Right';
    } else {
      alignment = 'Left';
    }

    let suggestedFont: string | undefined = undefined;
    if (isRtl && fontSetting !== 'Same as document') {
      suggestedFont = fontSetting;
    }

    return {
      isRtl,
      alignment,
      suggestedFont,
      arabicRatio: ratio
    };
  },

  /**
   * Wraps protected technical terms into prompt protection constraints
   */
  buildProtectedTermsPrompt(terms: string[]): string {
    if (!terms || terms.length === 0) return '';
    const formatted = terms.map(t => `"${t}"`).join(', ');
    return `CRITICAL INSTRUCTION - PROTECTED TERMS:
Do NOT translate, alter, spell out, or modify the casing of the following petroleum engineering abbreviations, well names, or units:
${formatted}.
Keep them exactly as written in English even when translating or formatting in Arabic.`;
  },

  /**
   * Scans text to check if any protected terms were altered or omitted
   */
  checkProtectedTermsIntegrity(original: string, processed: string, terms: string[]): { intact: boolean; missingTerms: string[] } {
    const missing: string[] = [];
    for (const term of terms) {
      const regex = new RegExp(`\\b${term}\\b`, 'i');
      if (regex.test(original)) {
        // Must appear in processed text (case-sensitive or insensitive)
        const targetRegex = new RegExp(`\\b${term}\\b`);
        if (!targetRegex.test(processed)) {
          missing.push(term);
        }
      }
    }
    return {
      intact: missing.length === 0,
      missingTerms: missing
    };
  },

  /**
   * Prepares bilingual prompt instruction
   */
  buildBilingualInstruction(enabled: boolean): string {
    if (!enabled) return '';
    return `BILINGUAL MODE ACTIVE:
For each bullet point, heading, or title, provide both English and Arabic versions:
- English first (LTR)
- Followed by Arabic (RTL)
Example:
• ESP Pump Frequency Optimization | تحسين تردد مضخة الـ ESP
• Target Rate: 4,500 bbl/d | المعدل المستهدف: 4,500 bbl/d`;
  }
};
