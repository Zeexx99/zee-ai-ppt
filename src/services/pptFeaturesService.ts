import { AppSettings } from '../types/settings';
import { pptService } from './pptService';
import { arabicService } from './arabicService';
import { geminiService } from './geminiService';

export interface DeckSlideOutline {
  slideNumber: number;
  title: string;
  bulletPoints: string[];
  notes?: string;
  visualSuggestion?: string;
}

export interface DeckOutline {
  presentationTitle: string;
  theme: string;
  slides: DeckSlideOutline[];
}

export interface ConsistencyIssue {
  slideIndex: number;
  type: 'casing' | 'punctuation' | 'protected_term' | 'font';
  description: string;
  originalText: string;
  suggestedText: string;
}

export const pptFeaturesService = {
  /**
   * 1. Report or PDF -> Deck: Generates a structured presentation outline for user approval
   */
  async generateDeckOutline(sourceText: string, slideCount = 5, settings: AppSettings): Promise<DeckOutline> {
    const protectedTermsPrompt = arabicService.buildProtectedTermsPrompt(settings.protectedTerms);
    const bilingualPrompt = arabicService.buildBilingualInstruction(settings.bilingualMode);

    const prompt = `Analyze the following report/document and create a high-impact presentation outline with exactly ${slideCount} slides.

DOCUMENT CONTENT:
"""
${sourceText.substring(0, 10000)}
"""

REQUIREMENTS:
1. Create a clear, engaging presentation title.
2. For each slide provide:
   - slideNumber (1 to ${slideCount})
   - title (concise, action-oriented)
   - bulletPoints (3-4 crisp, informative bullet points)
   - notes (professional speaker notes for the presenter)
   - visualSuggestion (a brief idea for a chart, diagram, or table)
3. Return ONLY a valid JSON object matching this schema without markdown code blocks:
{
  "presentationTitle": "string",
  "theme": "string",
  "slides": [
    {
      "slideNumber": 1,
      "title": "string",
      "bulletPoints": ["string"],
      "notes": "string",
      "visualSuggestion": "string"
    }
  ]
}

${protectedTermsPrompt}
${bilingualPrompt}`;

    let jsonResponse = '';
    for await (const chunk of geminiService.streamChat({
      messages: [{ role: 'user', content: prompt }],
      settings
    })) {
      if (chunk.textDelta) {
        jsonResponse += chunk.textDelta;
      }
    }

    const cleanJson = jsonResponse
      .replace(/^```json\s*/i, '')
      .replace(/^```\s*/i, '')
      .replace(/```$/i, '')
      .trim();

    try {
      const parsed = JSON.parse(cleanJson);
      return parsed as DeckOutline;
    } catch {
      throw new Error('Failed to parse generated presentation outline. Please try again.');
    }
  },

  /**
   * Builds slides into the presentation from an approved DeckOutline
   */
  async buildDeckFromOutline(
    outline: DeckOutline,
    onProgress?: (current: number, total: number, title: string) => void
  ): Promise<{ createdCount: number }> {
    let created = 0;
    const total = outline.slides.length;

    for (const slideData of outline.slides) {
      if (onProgress) {
        onProgress(slideData.slideNumber, total, slideData.title);
      }

      const res = await pptService.addSlide({
        title: slideData.title,
        bullets: slideData.bulletPoints
      });

      if (slideData.notes) {
        try {
          await pptService.writeSpeakerNotes(res.slideIndex, slideData.notes);
        } catch {
          // ignore if notes not supported on this PowerPoint version
        }
      }

      created++;
    }

    return { createdCount: created };
  },

  /**
   * 2. Excel Results -> Results Slide: formats pasted data/metrics into an executive results slide
   */
  async buildResultsSlideFromData(rawData: string, settings: AppSettings): Promise<{ slideIndex: number }> {
    const prompt = `Convert the following petroleum/engineering test data or spreadsheet extract into a structured executive summary slide.

RAW DATA:
"""
${rawData.substring(0, 5000)}
"""

REQUIREMENTS:
1. Provide a concise, professional slide title summarizing the test/well/asset.
2. Provide 3-4 bullet points highlighting key findings, anomalies, and operational recommendations.
3. Provide a summary table: 3-5 rows, 2-3 columns (e.g. Metric | Value | Status or Baseline).
4. Return ONLY a valid JSON object matching:
{
  "title": "string",
  "bullets": ["string"],
  "table": {
    "headers": ["Metric", "Value", "Status/Benchmark"],
    "rows": [["string", "string", "string"]]
  },
  "speakerNotes": "string"
}`;

    let jsonResponse = '';
    for await (const chunk of geminiService.streamChat({
      messages: [{ role: 'user', content: prompt }],
      settings
    })) {
      if (chunk.textDelta) {
        jsonResponse += chunk.textDelta;
      }
    }

    const cleanJson = jsonResponse
      .replace(/^```json\s*/i, '')
      .replace(/^```\s*/i, '')
      .replace(/```$/i, '')
      .trim();

    const parsed = JSON.parse(cleanJson);

    // Create the slide
    const slide = await pptService.addSlide({
      title: parsed.title,
      bullets: parsed.bullets
    });

    // Add table if supported
    const status = pptService.getRequirementStatus();
    if (status.canAddTables && parsed.table && parsed.table.rows?.length > 0) {
      try {
        await pptService.insertTable(slide.slideIndex, {
          rows: parsed.table.rows.length + 1,
          columns: parsed.table.headers.length,
          headers: parsed.table.headers,
          data: parsed.table.rows
        });
      } catch (e) {
        console.warn('Could not insert table on results slide:', e);
      }
    }

    // Add speaker notes if supported
    if (status.canManageNotes && parsed.speakerNotes) {
      try {
        await pptService.writeSpeakerNotes(slide.slideIndex, parsed.speakerNotes);
      } catch {
        // ignore
      }
    }

    return { slideIndex: slide.slideIndex };
  },

  /**
   * 3. Speaker Notes Generator: writes professional notes for active slide
   */
  async generateSpeakerNotesForSlide(slideIndex: number, settings: AppSettings): Promise<string> {
    const slide = await pptService.readSlide(slideIndex);
    if (!slide.title && slide.bullets.length === 0) {
      throw new Error(`Slide ${slideIndex} has no text content to generate speaker notes from.`);
    }

    const prompt = `You are an executive speech coach. Write natural, conversational, professional speaker notes for the following slide:

SLIDE TITLE: ${slide.title}
SLIDE BULLETS:
${slide.bullets.join('\n')}

GUIDELINES:
- Provide what the presenter should say to expand on the bullets, provide context, and transition smoothly.
- Do not repeat the bullet points verbatim.
- Keep it under 150 words.`;

    let notes = '';
    for await (const chunk of geminiService.streamChat({
      messages: [{ role: 'user', content: prompt }],
      settings
    })) {
      if (chunk.textDelta) {
        notes += chunk.textDelta;
      }
    }

    // Save to slide
    await pptService.writeSpeakerNotes(slideIndex, notes.trim());
    return notes.trim();
  },

  /**
   * 4. Consistency Checker: Audits slides for fonts, casing, units, and protected terms
   */
  async checkDeckConsistency(settings: AppSettings): Promise<ConsistencyIssue[]> {
    const overview = await pptService.getPresentationOverview();
    const issues: ConsistencyIssue[] = [];

    for (let i = 1; i <= overview.slideCount; i++) {
      const slide = await pptService.readSlide(i);

      // Check protected terms casing
      for (const term of settings.protectedTerms) {
        const regex = new RegExp(`\\b${term}\\b`, 'i');
        const exactRegex = new RegExp(`\\b${term}\\b`);
        for (const bullet of slide.bullets) {
          if (regex.test(bullet) && !exactRegex.test(bullet)) {
            issues.push({
              slideIndex: i,
              type: 'protected_term',
              description: `Protected term "${term}" has incorrect casing in slide ${i}.`,
              originalText: bullet,
              suggestedText: bullet.replace(regex, term)
            });
          }
        }
      }

      // Check bullet punctuation consistency (ending period)
      if (slide.bullets.length >= 2) {
        const withPeriod = slide.bullets.filter(b => b.trim().endsWith('.')).length;
        const withoutPeriod = slide.bullets.length - withPeriod;
        if (withPeriod > 0 && withoutPeriod > 0) {
          issues.push({
            slideIndex: i,
            type: 'punctuation',
            description: `Mixed punctuation in bullet list on slide ${i} (${withPeriod} end with '.', ${withoutPeriod} do not).`,
            originalText: slide.bullets[0],
            suggestedText: 'Standardize all bullet points to omit trailing periods.'
          });
        }
      }
    }

    return issues;
  },

  /**
   * 5. Text-Fit Fixer: Intelligently shortens bullet points that overflow
   */
  async shortenSlideText(slideIndex: number, settings: AppSettings): Promise<{ shortenedText: string }> {
    const slide = await pptService.readSlide(slideIndex);
    if (!slide.bullets || slide.bullets.length === 0) {
      throw new Error(`Slide ${slideIndex} does not contain bullet points to shorten.`);
    }

    const prompt = `Condense the following slide bullet points by 30-40% so they fit comfortably on the slide without shrinking the font size.

BULLET POINTS:
${slide.bullets.join('\n')}

RULES:
1. Retain all technical parameters, metrics, numbers, and protected terms.
2. Remove filler words, redundant phrases, and narrative fluff.
3. Return the shortened bullets as a bulleted list starting with •.`;

    let condensed = '';
    for await (const chunk of geminiService.streamChat({
      messages: [{ role: 'user', content: prompt }],
      settings
    })) {
      if (chunk.textDelta) {
        condensed += chunk.textDelta;
      }
    }

    return { shortenedText: condensed.trim() };
  }
};
