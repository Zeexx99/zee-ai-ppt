import { AgentStep, ToolApproval, UndoRecord } from '../types/chat';
import { AppSettings } from '../types/settings';
import { pptService } from './pptService';
import { arabicService } from './arabicService';
import { FunctionDeclaration, geminiService, ToolCall } from './geminiService';

export const AGENT_TOOLS: FunctionDeclaration[] = [
  {
    name: 'list_slides',
    description: 'Lists all slides in the presentation with their indices and titles (Read-only).',
    parameters: {
      type: 'object',
      properties: {}
    }
  },
  {
    name: 'read_slide',
    description: 'Reads the full text, titles, bullet points, and speaker notes of a specific slide (Read-only).',
    parameters: {
      type: 'object',
      properties: {
        slideIndex: { type: 'number', description: '1-based slide index to read' }
      },
      required: ['slideIndex']
    }
  },
  {
    name: 'read_selected_text',
    description: 'Reads text currently selected by the user in PowerPoint (Read-only).',
    parameters: {
      type: 'object',
      properties: {}
    }
  },
  {
    name: 'read_speaker_notes',
    description: 'Reads the presenter speaker notes for a slide (Read-only).',
    parameters: {
      type: 'object',
      properties: {
        slideIndex: { type: 'number', description: '1-based slide index' }
      },
      required: ['slideIndex']
    }
  },
  {
    name: 'add_slide',
    description: 'Creates a new slide with an optional title and bullet points.',
    parameters: {
      type: 'object',
      properties: {
        title: { type: 'string', description: 'Headline title for the new slide' },
        bullets: {
          type: 'array',
          items: { type: 'string' },
          description: 'Key bullet points for the slide body'
        }
      }
    }
  },
  {
    name: 'delete_slide',
    description: 'Permanently deletes a slide at the specified 1-based index.',
    parameters: {
      type: 'object',
      properties: {
        slideIndex: { type: 'number', description: '1-based slide index to delete' }
      },
      required: ['slideIndex']
    }
  },
  {
    name: 'insert_text_box',
    description: 'Inserts a new text box onto a slide with optional formatting (font size, bold, color, RTL).',
    parameters: {
      type: 'object',
      properties: {
        slideIndex: { type: 'number', description: '1-based slide index' },
        text: { type: 'string', description: 'Text content to insert' },
        fontSize: { type: 'number', description: 'Font size in points (e.g. 18)' },
        bold: { type: 'boolean', description: 'Whether text is bold' },
        color: { type: 'string', description: 'Hex color code, e.g. #D24726' },
        isRtl: { type: 'boolean', description: 'Set true if text is Arabic / right-to-left' }
      },
      required: ['slideIndex', 'text']
    }
  },
  {
    name: 'insert_table',
    description: 'Inserts a formatted grid table onto a slide with headers and data rows.',
    parameters: {
      type: 'object',
      properties: {
        slideIndex: { type: 'number', description: '1-based slide index' },
        rows: { type: 'number', description: 'Number of rows' },
        columns: { type: 'number', description: 'Number of columns' },
        headers: {
          type: 'array',
          items: { type: 'string' },
          description: 'Header column titles'
        },
        data: {
          type: 'array',
          items: { type: 'array', items: { type: 'string' } },
          description: '2D array of table row values'
        }
      },
      required: ['slideIndex', 'rows', 'columns']
    }
  },
  {
    name: 'write_speaker_notes',
    description: 'Writes or appends presenter speaker notes to a slide.',
    parameters: {
      type: 'object',
      properties: {
        slideIndex: { type: 'number', description: '1-based slide index' },
        notesText: { type: 'string', description: 'Speaker notes text' },
        append: { type: 'boolean', description: 'Whether to append or overwrite' }
      },
      required: ['slideIndex', 'notesText']
    }
  },
  {
    name: 'replace_text',
    description: 'Finds and replaces text across a slide or the entire presentation.',
    parameters: {
      type: 'object',
      properties: {
        search: { type: 'string', description: 'Exact string to find' },
        replacement: { type: 'string', description: 'Replacement string' },
        slideIndex: { type: 'number', description: 'Optional 1-based slide index to restrict replacement' }
      },
      required: ['search', 'replacement']
    }
  }
];

export const agentService = {
  /**
   * Identifies if a tool call is read-only or mutates the document
   */
  isReadOnlyTool(name: string): boolean {
    return ['list_slides', 'read_slide', 'read_selected_text', 'read_speaker_notes'].includes(name);
  },

  /**
   * Formats a human-readable description for an approval card
   */
  formatApprovalDescription(toolName: string, args: Record<string, unknown>): string {
    switch (toolName) {
      case 'add_slide':
        return `Add new slide${args.title ? `: "${args.title}"` : ''} with ${Array.isArray(args.bullets) ? args.bullets.length : 0} bullet points`;
      case 'delete_slide':
        return `Delete slide ${args.slideIndex}`;
      case 'insert_text_box':
        return `Insert text box on slide ${args.slideIndex}: "${String(args.text || '').substring(0, 40)}..."`;
      case 'insert_table':
        return `Insert ${args.rows}x${args.columns} table on slide ${args.slideIndex}`;
      case 'write_speaker_notes':
        return `Update speaker notes for slide ${args.slideIndex}`;
      case 'replace_text':
        return `Replace "${args.search}" with "${args.replacement}"`;
      default:
        return `Execute ${toolName}`;
    }
  },

  /**
   * Executes a tool with snapshot recording for undo
   */
  async executeTool(
    toolName: string,
    args: Record<string, unknown>,
    settings: AppSettings
  ): Promise<{ result: string; undoRecord?: UndoRecord }> {
    switch (toolName) {
      case 'list_slides': {
        const overview = await pptService.getPresentationOverview();
        const summary = overview.slides
          .map(s => `Slide ${s.index}: "${s.title}" (${s.shapeCount} shapes)`)
          .join('\n');
        return {
          result: `Total slides: ${overview.slideCount}\n${summary || 'No slides found'}`
        };
      }

      case 'read_slide': {
        const idx = Number(args.slideIndex) || 1;
        const data = await pptService.readSlide(idx);
        let notes = '';
        try {
          notes = await pptService.readSpeakerNotes(idx);
        } catch {
          // ignore
        }
        return {
          result: `Slide ${idx}:\nTitle: ${data.title}\nBullets:\n${data.bullets.join('\n')}\nNotes: ${notes || '(none)'}`
        };
      }

      case 'read_selected_text': {
        const text = await pptService.getSelectedText();
        return { result: text ? `Selected text: "${text}"` : 'No text is currently selected.' };
      }

      case 'read_speaker_notes': {
        const idx = Number(args.slideIndex) || 1;
        const notes = await pptService.readSpeakerNotes(idx);
        return { result: notes ? `Notes for slide ${idx}:\n${notes}` : `Slide ${idx} has no speaker notes.` };
      }

      case 'add_slide': {
        const title = args.title ? String(args.title) : undefined;
        const bullets = Array.isArray(args.bullets) ? (args.bullets as string[]) : undefined;

        const res = await pptService.addSlide({ title, bullets });
        const undo: UndoRecord = {
          id: `undo-${Date.now()}`,
          stepId: `step-${Date.now()}`,
          description: `Add slide ${res.slideIndex}`,
          timestamp: Date.now(),
          actionType: 'add_slide',
          snapshotData: { slideIndex: res.slideIndex }
        };

        return {
          result: `Successfully added slide ${res.slideIndex}`,
          undoRecord: undo
        };
      }

      case 'delete_slide': {
        const idx = Number(args.slideIndex);
        // Snapshot before deletion
        const priorData = await pptService.readSlide(idx);
        await pptService.deleteSlide(idx);

        const undo: UndoRecord = {
          id: `undo-${Date.now()}`,
          stepId: `step-${Date.now()}`,
          description: `Deleted slide ${idx}`,
          timestamp: Date.now(),
          actionType: 'delete_slide',
          snapshotData: { priorData, slideIndex: idx }
        };

        return {
          result: `Successfully deleted slide ${idx}`,
          undoRecord: undo
        };
      }

      case 'insert_text_box': {
        const idx = Number(args.slideIndex) || 1;
        const text = String(args.text || '');
        const isRtl = args.isRtl !== undefined ? Boolean(args.isRtl) : arabicService.hasArabic(text);

        await pptService.insertTextBox(idx, text, {
          fontSize: args.fontSize ? Number(args.fontSize) : 18,
          bold: args.bold !== undefined ? Boolean(args.bold) : false,
          color: args.color ? String(args.color) : settings.companyTemplate.textColor,
          isRtl
        });

        return {
          result: `Inserted text box on slide ${idx}`
        };
      }

      case 'insert_table': {
        const idx = Number(args.slideIndex) || 1;
        const rows = Number(args.rows) || 2;
        const columns = Number(args.columns) || 2;
        const headers = Array.isArray(args.headers) ? (args.headers as string[]) : undefined;
        const data = Array.isArray(args.data) ? (args.data as string[][]) : undefined;

        await pptService.insertTable(idx, { rows, columns, headers, data });
        return {
          result: `Inserted ${rows}x${columns} table on slide ${idx}`
        };
      }

      case 'write_speaker_notes': {
        const idx = Number(args.slideIndex) || 1;
        const text = String(args.notesText || '');
        const append = Boolean(args.append);

        // Snapshot prior notes
        const priorNotes = await pptService.readSpeakerNotes(idx);
        await pptService.writeSpeakerNotes(idx, text, append);

        const undo: UndoRecord = {
          id: `undo-${Date.now()}`,
          stepId: `step-${Date.now()}`,
          description: `Updated notes on slide ${idx}`,
          timestamp: Date.now(),
          actionType: 'write_speaker_notes',
          snapshotData: { slideIndex: idx, priorNotes }
        };

        return {
          result: `Updated speaker notes for slide ${idx}`,
          undoRecord: undo
        };
      }

      case 'replace_text': {
        const search = String(args.search || '');
        const replacement = String(args.replacement || '');
        const slideIndex = args.slideIndex ? Number(args.slideIndex) : undefined;

        const res = await pptService.replaceText(search, replacement, slideIndex);
        const undo: UndoRecord = {
          id: `undo-${Date.now()}`,
          stepId: `step-${Date.now()}`,
          description: `Replaced "${search}" with "${replacement}"`,
          timestamp: Date.now(),
          actionType: 'replace_text',
          snapshotData: { search: replacement, replacement: search, slideIndex }
        };

        return {
          result: `Replaced ${res.replacementsCount} occurrence(s) of "${search}"`,
          undoRecord: undo
        };
      }

      default:
        throw new Error(`Unknown tool: ${toolName}`);
    }
  },

  /**
   * Reverts an action using its recorded undo snapshot
   */
  async undoAction(record: UndoRecord): Promise<string> {
    switch (record.actionType) {
      case 'add_slide': {
        const data = record.snapshotData as { slideIndex: number };
        await pptService.deleteSlide(data.slideIndex);
        return `Deleted added slide ${data.slideIndex}`;
      }

      case 'delete_slide': {
        const data = record.snapshotData as { priorData: { title: string; bullets: string[] } };
        await pptService.addSlide({
          title: data.priorData.title,
          bullets: data.priorData.bullets
        });
        return `Restored deleted slide "${data.priorData.title}"`;
      }

      case 'write_speaker_notes': {
        const data = record.snapshotData as { slideIndex: number; priorNotes: string };
        await pptService.writeSpeakerNotes(data.slideIndex, data.priorNotes, false);
        return `Restored prior notes for slide ${data.slideIndex}`;
      }

      case 'replace_text': {
        const data = record.snapshotData as { search: string; replacement: string; slideIndex?: number };
        await pptService.replaceText(data.search, data.replacement, data.slideIndex);
        return `Reverted text replacement`;
      }

      default:
        return `Undo not supported for ${record.actionType}`;
    }
  },

  /**
   * Builds the agent system prompt incorporating presentation context, protected terms, and Arabic direction
   */
  buildSystemPrompt(overview: { slideCount: number; slides: { index: number; title: string }[] }, settings: AppSettings): string {
    const protectedTermsClause = arabicService.buildProtectedTermsPrompt(settings.protectedTerms);
    const bilingualClause = arabicService.buildBilingualInstruction(settings.bilingualMode);

    const slidesSummary = overview.slides.map(s => `Slide ${s.index}: "${s.title}"`).join('\n');

    return `You are Zee AI for PowerPoint, an expert AI presentation assistant.
You are helping the user build, edit, summarize, and polish PowerPoint presentations.

PRESENTATION OVERVIEW:
Total slides: ${overview.slideCount}
${slidesSummary || '(Presentation currently empty)'}

CORE BEHAVIOR & RULES:
1. Always be concise, clear, and professional.
2. When creating or revising slides, use clean, impactful titles and 3-5 punchy bullet points.
3. Call specific tools (like add_slide, insert_table, write_speaker_notes) to make edits.
4. For petroleum engineering data (ESP pumps, decline curves, GOR, BHP, etc.), present key metrics prominently in bullet points or tables.
5. Never translate or change protected petroleum terms or abbreviations.
6. If the user writes in Arabic, format text appropriately with RTL orientation.

${protectedTermsClause}

${bilingualClause}`;
  }
};
