import { describe, it, expect } from 'vitest';
import { arabicService } from '../src/services/arabicService';

describe('arabicService', () => {
  it('detects Arabic script correctly', () => {
    expect(arabicService.hasArabic('Hello World')).toBe(false);
    expect(arabicService.hasArabic('مرحبا بالعالم')).toBe(true);
    expect(arabicService.hasArabic('ESP pump analysis: تحليل مضخة')).toBe(true);
  });

  it('calculates Arabic character ratio', () => {
    const pureEnglish = 'Total Dynamic Head calculation';
    expect(arabicService.getArabicRatio(pureEnglish)).toBe(0);

    const pureArabic = 'تقرير أداء المضخة الكهربائية الغاطسة';
    expect(arabicService.getArabicRatio(pureArabic)).toBeGreaterThan(0.9);
  });

  it('analyzes text direction and recommends RTL and fonts', () => {
    const arabicHeading = 'تقرير الآبار الربع سنوي';
    const analysis = arabicService.analyzeDirection(arabicHeading, 'Cairo', true);

    expect(analysis.isRtl).toBe(true);
    expect(analysis.alignment).toBe('Center'); // Titles centered
    expect(analysis.suggestedFont).toBe('Cairo');

    const englishBody = 'The current pump intake pressure is 1,200 psi';
    const enAnalysis = arabicService.analyzeDirection(englishBody, 'Cairo', false);
    expect(enAnalysis.isRtl).toBe(false);
    expect(enAnalysis.alignment).toBe('Left');
    expect(enAnalysis.suggestedFont).toBeUndefined();
  });

  it('generates protected terms prompt correctly', () => {
    const terms = ['ESP', 'TDH', 'PIP', 'GOR', 'BHP', 'bbl/d'];
    const prompt = arabicService.buildProtectedTermsPrompt(terms);

    expect(prompt).toContain('CRITICAL INSTRUCTION - PROTECTED TERMS');
    expect(prompt).toContain('"ESP"');
    expect(prompt).toContain('"bbl/d"');
  });

  it('detects missing or altered protected terms', () => {
    const original = 'Verify ESP motor temperature and TDH at 60 Hz.';
    const goodOutput = 'تأكد من حرارة محرك الـ ESP وقيمة TDH عند 60 Hz.';
    const badOutput = 'تأكد من حرارة المحرك وقيمة الرأس الهيدروليكي.';

    const terms = ['ESP', 'TDH', 'PIP'];

    const checkGood = arabicService.checkProtectedTermsIntegrity(original, goodOutput, terms);
    expect(checkGood.intact).toBe(true);
    expect(checkGood.missingTerms).toHaveLength(0);

    const checkBad = arabicService.checkProtectedTermsIntegrity(original, badOutput, terms);
    expect(checkBad.intact).toBe(false);
    expect(checkBad.missingTerms).toContain('ESP');
    expect(checkBad.missingTerms).toContain('TDH');
  });

  it('generates bilingual instructions', () => {
    const enabled = arabicService.buildBilingualInstruction(true);
    expect(enabled).toContain('BILINGUAL MODE ACTIVE');
    expect(enabled).toContain('English first');

    const disabled = arabicService.buildBilingualInstruction(false);
    expect(disabled).toBe('');
  });
});
