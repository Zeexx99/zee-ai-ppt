import {
  PresentationOverview,
  RequirementSetStatus,
  SlideSummary,
  TableInsertOptions,
  TextFormatOptions
} from '../types/ppt';

/* global Office, PowerPoint */

export const pptService = {
  /**
   * Detects supported PowerPointApi requirement sets
   */
  getRequirementStatus(): RequirementSetStatus {
    const isSupported = (ver: string) => {
      try {
        return (
          typeof Office !== 'undefined' &&
          Office.context &&
          Office.context.requirements &&
          Office.context.requirements.isSetSupported('PowerPointApi', ver)
        );
      } catch {
        return false;
      }
    };

    const hasApi11 = isSupported('1.1');
    const hasApi12 = isSupported('1.2');
    const hasApi13 = isSupported('1.3');
    const hasApi14 = isSupported('1.4');
    const hasApi15 = isSupported('1.5');

    let highest = '1.0';
    if (hasApi15) highest = '1.5';
    else if (hasApi14) highest = '1.4';
    else if (hasApi13) highest = '1.3';
    else if (hasApi12) highest = '1.2';
    else if (hasApi11) highest = '1.1';

    return {
      hasApi11,
      hasApi12,
      hasApi13,
      hasApi14,
      hasApi15,
      canFormatShapes: hasApi13,
      canAddTables: hasApi14,
      canManageNotes: hasApi15,
      canInsertSlidesBase64: hasApi12,
      highestVersion: highest
    };
  },

  /**
   * Promisified selected text retrieval
   */
  async getSelectedText(): Promise<string> {
    return new Promise(resolve => {
      try {
        if (typeof Office === 'undefined' || !Office.context || !Office.context.document) {
          resolve('');
          return;
        }
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, result => {
          if (result.status === Office.AsyncResultStatus.Succeeded && typeof result.value === 'string') {
            resolve(result.value.trim());
          } else {
            resolve('');
          }
        });
      } catch {
        resolve('');
      }
    });
  },

  /**
   * Promisified replacement of selected text
   */
  async setSelectedText(text: string): Promise<boolean> {
    return new Promise(resolve => {
      try {
        if (typeof Office === 'undefined' || !Office.context || !Office.context.document) {
          resolve(false);
          return;
        }
        Office.context.document.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text }, result => {
          resolve(result.status === Office.AsyncResultStatus.Succeeded);
        });
      } catch {
        resolve(false);
      }
    });
  },

  /**
   * Retrieves compact presentation overview (slides, count, titles, active slide)
   */
  async getPresentationOverview(): Promise<PresentationOverview> {
    const selectedText = await this.getSelectedText();

    if (typeof PowerPoint === 'undefined' || !PowerPoint.run) {
      return {
        slideCount: 0,
        activeSlideIndex: 1,
        slides: [],
        selectedText
      };
    }

    try {
      return await PowerPoint.run(async (context: PowerPoint.RequestContext) => {
        const presentation = context.presentation;
        const slides = presentation.slides;
        slides.load('items');
        await context.sync();

        const slideCount = slides.items.length;
        const slideSummaries: SlideSummary[] = [];

        // Load shapes for each slide
        for (let i = 0; i < slides.items.length; i++) {
          const slide = slides.items[i];
          const shapes = slide.shapes;
          shapes.load('items');
        }
        await context.sync();

        for (let i = 0; i < slides.items.length; i++) {
          const slide = slides.items[i];
          const shapes = slide.shapes.items;

          // Preload textframes
          for (const shape of shapes) {
            shape.load(['name', 'type']);
            if (shape.textFrame) {
              shape.textFrame.load('hasText');
              shape.textFrame.textRange.load('text');
            }
          }
        }
        await context.sync();

        for (let i = 0; i < slides.items.length; i++) {
          const slide = slides.items[i];
          const shapes = slide.shapes.items;
          let foundTitle = '';
          const contentParts: string[] = [];

          for (const shape of shapes) {
            if (shape.textFrame && shape.textFrame.hasText) {
              const text = shape.textFrame.textRange.text.trim();
              if (text) {
                if (!foundTitle && (shape.name?.toLowerCase().includes('title') || text.length < 80)) {
                  foundTitle = text;
                } else {
                  contentParts.push(text);
                }
              }
            }
          }

          slideSummaries.push({
            id: slide.id || `slide-${i + 1}`,
            index: i + 1,
            title: foundTitle || `Slide ${i + 1}`,
            contentSnippet: contentParts.slice(0, 3).join(' | ').substring(0, 200),
            shapeCount: shapes.length
          });
        }

        return {
          slideCount,
          activeSlideIndex: 1,
          slides: slideSummaries,
          selectedText
        };
      });
    } catch (e) {
      console.warn('Could not read presentation overview:', e);
      return {
        slideCount: 0,
        activeSlideIndex: 1,
        slides: [],
        selectedText
      };
    }
  },

  /**
   * Reads full slide text and shapes
   */
  async readSlide(slideIndex: number): Promise<{ title: string; bullets: string[]; allText: string }> {
    return await PowerPoint.run(async (context: PowerPoint.RequestContext) => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();

      const idx = Math.max(0, Math.min(slideIndex - 1, slides.items.length - 1));
      if (slides.items.length === 0) {
        return { title: '', bullets: [], allText: '' };
      }

      const slide = slides.items[idx];
      const shapes = slide.shapes;
      shapes.load('items');
      await context.sync();

      for (const shape of shapes.items) {
        if (shape.textFrame) {
          shape.textFrame.load('hasText');
          shape.textFrame.textRange.load('text');
        }
      }
      await context.sync();

      let title = '';
      const bullets: string[] = [];
      const allParts: string[] = [];

      for (const shape of shapes.items) {
        if (shape.textFrame && shape.textFrame.hasText) {
          const text = shape.textFrame.textRange.text.trim();
          if (text) {
            allParts.push(text);
            if (!title) {
              title = text;
            } else {
              bullets.push(text);
            }
          }
        }
      }

      return {
        title,
        bullets,
        allText: allParts.join('\n\n')
      };
    });
  },

  /**
   * Adds a new slide with optional title and content text box
   */
  async addSlide(options?: { title?: string; bullets?: string[]; index?: number }): Promise<{ slideIndex: number }> {
    return await PowerPoint.run(async (context: PowerPoint.RequestContext) => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();

      // In PowerPointApi, slides.add() returns void; reload slides collection
      slides.add();
      await context.sync();

      slides.load('items');
      await context.sync();

      const newSlideIndex = slides.items.length;
      const newSlide = slides.items[newSlideIndex - 1];

      // Add title if requested
      if (options?.title) {
        const titleBox = newSlide.shapes.addTextBox(options.title, {
          left: 50,
          top: 40,
          width: 620,
          height: 60
        });
        titleBox.textFrame.textRange.font.size = 28;
        titleBox.textFrame.textRange.font.bold = true;
        titleBox.textFrame.textRange.font.color = '#D24726';
      }

      // Add bullets/body if requested
      if (options?.bullets && options.bullets.length > 0) {
        const bodyText = options.bullets.map(b => (b.startsWith('•') ? b : `•  ${b}`)).join('\n\n');
        const contentBox = newSlide.shapes.addTextBox(bodyText, {
          left: 50,
          top: 120,
          width: 620,
          height: 380
        });
        contentBox.textFrame.textRange.font.size = 18;
      }

      await context.sync();
      return { slideIndex: newSlideIndex };
    });
  },

  /**
   * Deletes a slide at specified 1-based index
   */
  async deleteSlide(slideIndex: number): Promise<boolean> {
    return await PowerPoint.run(async (context: PowerPoint.RequestContext) => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();

      const idx = slideIndex - 1;
      if (idx < 0 || idx >= slides.items.length) {
        throw new Error(`Slide index ${slideIndex} out of range (total: ${slides.items.length})`);
      }

      const slide = slides.items[idx];
      slide.delete();
      await context.sync();
      return true;
    });
  },

  /**
   * Inserts a text box on a slide with optional formatting and RTL support
   */
  async insertTextBox(
    slideIndex: number,
    text: string,
    options?: TextFormatOptions & { left?: number; top?: number; width?: number; height?: number }
  ): Promise<boolean> {
    return await PowerPoint.run(async (context: PowerPoint.RequestContext) => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();

      const idx = Math.max(0, Math.min(slideIndex - 1, slides.items.length - 1));
      const slide = slides.items[idx];

      const left = options?.left ?? 50;
      const top = options?.top ?? 150;
      const width = options?.width ?? 620;
      const height = options?.height ?? 200;

      const textBox = slide.shapes.addTextBox(text, { left, top, width, height });
      const range = textBox.textFrame.textRange;

      if (options?.fontSize) range.font.size = options.fontSize;
      if (options?.fontName) range.font.name = options.fontName;
      if (options?.bold !== undefined) range.font.bold = options.bold;
      if (options?.italic !== undefined) range.font.italic = options.italic;
      if (options?.color) range.font.color = options.color;

      await context.sync();
      return true;
    });
  },

  /**
   * Inserts a formatted table on the slide (Requirement set 1.4+)
   */
  async insertTable(slideIndex: number, options: TableInsertOptions): Promise<boolean> {
    const status = this.getRequirementStatus();
    if (!status.canAddTables) {
      throw new Error('Table creation requires PowerPointApi 1.4 or newer.');
    }

    return await PowerPoint.run(async (context: PowerPoint.RequestContext) => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();

      const idx = Math.max(0, Math.min(slideIndex - 1, slides.items.length - 1));
      const slide = slides.items[idx];

      const rows = Math.max(1, options.rows);
      const cols = Math.max(1, options.columns);

      // Add table shape
      const tableShape = slide.shapes.addTable(rows, cols, {
        left: 50,
        top: options.title ? 110 : 80,
        width: 620,
        height: Math.min(400, rows * 40 + 30)
      });

      const table = tableShape.getTable();

      // Populate headers if provided
      if (options.headers && options.headers.length > 0) {
        for (let c = 0; c < Math.min(cols, options.headers.length); c++) {
          const cell = table.getCellOrNullObject(0, c);
          cell.text = options.headers[c];
          cell.font.bold = true;
          cell.font.size = 14;
        }
      }

      // Populate data rows if provided
      if (options.data && options.data.length > 0) {
        const startRow = options.headers ? 1 : 0;
        for (let r = 0; r < options.data.length; r++) {
          const rowIdx = startRow + r;
          if (rowIdx >= rows) break;
          const rowData = options.data[r];
          for (let c = 0; c < Math.min(cols, rowData.length); c++) {
            const cell = table.getCellOrNullObject(rowIdx, c);
            cell.text = rowData[c];
            cell.font.size = 13;
          }
        }
      }

      await context.sync();
      return true;
    });
  },

  /**
   * Inserts an image using Office.CoercionType.Image
   */
  async insertImage(
    slideIndex: number,
    base64Data: string,
    options?: { left?: number; top?: number; width?: number; height?: number }
  ): Promise<boolean> {
    const cleanBase64 = base64Data.replace(/^data:image\/[a-z]+;base64,/, '');

    return new Promise(resolve => {
      try {
        if (typeof Office === 'undefined' || !Office.context || !Office.context.document) {
          resolve(false);
          return;
        }

        Office.context.document.setSelectedDataAsync(
          cleanBase64,
          {
            coercionType: Office.CoercionType.Image,
            imageLeft: options?.left ?? 100,
            imageTop: options?.top ?? 120,
            imageWidth: options?.width ?? 520,
            imageHeight: options?.height ?? 300
          },
          result => {
            resolve(result.status === Office.AsyncResultStatus.Succeeded);
          }
        );
      } catch {
        resolve(false);
      }
    });
  },

  /**
   * Reads speaker notes for a slide
   */
  async readSpeakerNotes(slideIndex: number): Promise<string> {
    const status = this.getRequirementStatus();
    if (!status.canManageNotes) {
      return '';
    }

    try {
      return await PowerPoint.run(async (context: PowerPoint.RequestContext) => {
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();

        const idx = Math.max(0, Math.min(slideIndex - 1, slides.items.length - 1));
        const slide = slides.items[idx];

        // @ts-expect-error notesSlide in 1.5+
        const notesSlide = slide.notesSlide;
        if (!notesSlide) return '';

        notesSlide.load('textFrame');
        await context.sync();

        if (notesSlide.textFrame) {
          notesSlide.textFrame.textRange.load('text');
          await context.sync();
          return notesSlide.textFrame.textRange.text || '';
        }
        return '';
      });
    } catch {
      return '';
    }
  },

  /**
   * Writes or appends speaker notes for a slide
   */
  async writeSpeakerNotes(slideIndex: number, notesText: string, append = false): Promise<boolean> {
    const status = this.getRequirementStatus();
    if (!status.canManageNotes) {
      throw new Error('Speaker notes editing requires PowerPointApi 1.5 or newer.');
    }

    return await PowerPoint.run(async (context: PowerPoint.RequestContext) => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();

      const idx = Math.max(0, Math.min(slideIndex - 1, slides.items.length - 1));
      const slide = slides.items[idx];

      // @ts-expect-error notesSlide in 1.5+
      const notesSlide = slide.notesSlide;
      notesSlide.load('textFrame');
      await context.sync();

      if (notesSlide.textFrame) {
        notesSlide.textFrame.textRange.load('text');
        await context.sync();
        const current = notesSlide.textFrame.textRange.text || '';
        notesSlide.textFrame.textRange.text = append && current ? `${current}\n\n${notesText}` : notesText;
      }

      await context.sync();
      return true;
    });
  },

  /**
   * Replaces text in all shapes of a slide or across all slides
   */
  async replaceText(search: string, replacement: string, slideIndex?: number): Promise<{ replacementsCount: number }> {
    return await PowerPoint.run(async (context: PowerPoint.RequestContext) => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();

      let targetSlides = slides.items;
      if (slideIndex !== undefined) {
        const idx = Math.max(0, Math.min(slideIndex - 1, slides.items.length - 1));
        targetSlides = [slides.items[idx]];
      }

      for (const slide of targetSlides) {
        slide.shapes.load('items');
      }
      await context.sync();

      for (const slide of targetSlides) {
        for (const shape of slide.shapes.items) {
          if (shape.textFrame) {
            shape.textFrame.load('hasText');
            shape.textFrame.textRange.load('text');
          }
        }
      }
      await context.sync();

      let count = 0;
      for (const slide of targetSlides) {
        for (const shape of slide.shapes.items) {
          if (shape.textFrame && shape.textFrame.hasText) {
            const current = shape.textFrame.textRange.text;
            if (current.includes(search)) {
              shape.textFrame.textRange.text = current.replaceAll(search, replacement);
              count++;
            }
          }
        }
      }

      await context.sync();
      return { replacementsCount: count };
    });
  }
};
