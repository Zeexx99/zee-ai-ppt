import { AgentToolStep, ChatAttachment, ChatMessage } from "../types/chat";
import { PowerPointApiRequirements } from "../types/ppt";
import { ZeeSettings } from "../types/settings";
import {
  buildGeminiContents,
  buildSystemInstruction,
  GeminiFunctionCall,
  GeminiFunctionDeclaration,
  GeminiTurnContent,
  generateGeminiTurn,
} from "./geminiService";
import {
  addSlide,
  applySlideLayout,
  captureSlideSnapshot,
  deleteSlide,
  formatCompactPresentationSummary,
  formatSlideText,
  getActiveSlideIndex,
  getPresentationSummary,
  getSelectedText,
  insertImage,
  insertTable,
  insertTextBox,
  listSlides,
  readSlide,
  readSpeakerNotes,
  reorderSlide,
  replaceSlideText,
  restoreUndoSnapshot,
  setSlideTitle,
  writeSpeakerNotes,
} from "./pptService";

export const READ_ONLY_TOOLS = new Set<string>([
  "listSlides",
  "readSlide",
  "getSelectedText",
  "readSpeakerNotes",
]);

export const WRITE_TOOLS = new Set<string>([
  "writeSpeakerNotes",
  "addSlide",
  "deleteSlide",
  "reorderSlide",
  "setSlideTitle",
  "insertTextBox",
  "replaceSlideText",
  "insertTable",
  "insertImage",
  "formatSlideText",
  "applySlideLayout",
]);

export const ALL_AGENT_TOOL_DECLARATIONS: GeminiFunctionDeclaration[] = [
  {
    name: "listSlides",
    description:
      "List all slides in the presentation with their 1-based slide numbers, titles, shape counts, and text previews.",
    parameters: {
      type: "OBJECT",
      properties: {},
    },
  },
  {
    name: "readSlide",
    description:
      "Read full details of a specific slide (1-based slideNumber), including all text shapes, fonts, positions, and speaker notes.",
    parameters: {
      type: "OBJECT",
      properties: {
        slideNumber: {
          type: "INTEGER",
          description: "1-based slide number to read.",
        },
      },
      required: ["slideNumber"],
    },
  },
  {
    name: "getSelectedText",
    description: "Read the text or shape currently selected by the user in PowerPoint.",
    parameters: {
      type: "OBJECT",
      properties: {},
    },
  },
  {
    name: "readSpeakerNotes",
    description: "Read the speaker notes for a specific 1-based slideNumber.",
    parameters: {
      type: "OBJECT",
      properties: {
        slideNumber: {
          type: "INTEGER",
          description: "1-based slide number.",
        },
      },
      required: ["slideNumber"],
    },
  },
  {
    name: "writeSpeakerNotes",
    description: "Write or append speaker notes for a specific 1-based slideNumber.",
    parameters: {
      type: "OBJECT",
      properties: {
        slideNumber: {
          type: "INTEGER",
          description: "1-based slide number.",
        },
        notes: {
          type: "STRING",
          description: "Speaker notes text to write.",
        },
        append: {
          type: "BOOLEAN",
          description: "If true, append to existing notes instead of replacing.",
        },
      },
      required: ["slideNumber", "notes"],
    },
  },
  {
    name: "addSlide",
    description:
      "Add a new slide to the presentation with a title, bullet points or body text, optional layout name, and optional speaker notes.",
    parameters: {
      type: "OBJECT",
      properties: {
        title: {
          type: "STRING",
          description: "Slide title.",
        },
        bodyText: {
          type: "STRING",
          description: "Bullet points or body text separated by newlines.",
        },
        layoutName: {
          type: "STRING",
          description:
            "Layout name (e.g. 'Title and Content', 'Title Slide', 'Two Content', 'Section Header').",
        },
        notes: {
          type: "STRING",
          description: "Optional speaker notes for the new slide.",
        },
      },
      required: ["title"],
    },
  },
  {
    name: "deleteSlide",
    description: "Delete a slide at the specified 1-based slideNumber.",
    parameters: {
      type: "OBJECT",
      properties: {
        slideNumber: {
          type: "INTEGER",
          description: "1-based slide number to delete.",
        },
      },
      required: ["slideNumber"],
    },
  },
  {
    name: "reorderSlide",
    description: "Move a slide from fromSlideNumber (1-based) to toSlideNumber (1-based).",
    parameters: {
      type: "OBJECT",
      properties: {
        fromSlideNumber: {
          type: "INTEGER",
          description: "Current 1-based slide position.",
        },
        toSlideNumber: {
          type: "INTEGER",
          description: "Target 1-based slide position.",
        },
      },
      required: ["fromSlideNumber", "toSlideNumber"],
    },
  },
  {
    name: "setSlideTitle",
    description: "Insert or update the title of a specific slide (1-based slideNumber).",
    parameters: {
      type: "OBJECT",
      properties: {
        slideNumber: {
          type: "INTEGER",
          description: "1-based slide number.",
        },
        title: {
          type: "STRING",
          description: "New title text for the slide.",
        },
      },
      required: ["slideNumber", "title"],
    },
  },
  {
    name: "insertTextBox",
    description:
      "Insert a text box onto a slide with specified text, optional position/size (in points), and optional heading style.",
    parameters: {
      type: "OBJECT",
      properties: {
        slideNumber: {
          type: "INTEGER",
          description: "1-based slide number.",
        },
        text: {
          type: "STRING",
          description: "Text content to insert.",
        },
        left: { type: "NUMBER", description: "Left position in points (default 48)." },
        top: { type: "NUMBER", description: "Top position in points (default 110)." },
        width: { type: "NUMBER", description: "Width in points (default 624)." },
        height: { type: "NUMBER", description: "Height in points (default 260)." },
        isHeading: { type: "BOOLEAN", description: "Whether to style as a centered heading." },
      },
      required: ["slideNumber", "text"],
    },
  },
  {
    name: "replaceSlideText",
    description:
      "Find and replace text on a specific slide (or across all slides when slideNumber is 0 or omitted), respecting protected engineering terms.",
    parameters: {
      type: "OBJECT",
      properties: {
        slideNumber: {
          type: "INTEGER",
          description: "1-based slide number, or 0 for all slides.",
        },
        findText: {
          type: "STRING",
          description: "Text to find.",
        },
        replaceWith: {
          type: "STRING",
          description: "Replacement text.",
        },
      },
      required: ["findText", "replaceWith"],
    },
  },
  {
    name: "insertTable",
    description: "Insert a formatted data table onto a slide with column headers and row cells.",
    parameters: {
      type: "OBJECT",
      properties: {
        slideNumber: {
          type: "INTEGER",
          description: "1-based slide number.",
        },
        title: {
          type: "STRING",
          description: "Optional table caption/title.",
        },
        headers: {
          type: "ARRAY",
          description: "Array of column header strings.",
          items: { type: "STRING" },
        },
        rowsJson: {
          type: "STRING",
          description:
            "JSON 2D array of row strings, e.g. '[[\"WELL-01\",\"1,250 bbl/d\",\"1,180 psi\"]]'.",
        },
      },
      required: ["slideNumber", "headers", "rowsJson"],
    },
  },
  {
    name: "insertImage",
    description: "Insert a Base64-encoded image onto a slide.",
    parameters: {
      type: "OBJECT",
      properties: {
        slideNumber: {
          type: "INTEGER",
          description: "1-based slide number.",
        },
        base64Data: {
          type: "STRING",
          description: "Base64 image data or attachment reference.",
        },
        left: { type: "NUMBER", description: "Left position in points." },
        top: { type: "NUMBER", description: "Top position in points." },
        width: { type: "NUMBER", description: "Width in points." },
        height: { type: "NUMBER", description: "Height in points." },
      },
      required: ["slideNumber", "base64Data"],
    },
  },
  {
    name: "formatSlideText",
    description:
      "Format text on a slide (font family, font size, hex color, bold, italic, alignment Left/Center/Right, and RTL/LTR direction).",
    parameters: {
      type: "OBJECT",
      properties: {
        slideNumber: {
          type: "INTEGER",
          description: "1-based slide number.",
        },
        shapeId: {
          type: "STRING",
          description: "Optional specific shapeId to format (omit to format all text shapes on the slide).",
        },
        fontName: {
          type: "STRING",
          description: "Font family name (e.g. 'Calibri', 'Simplified Arabic').",
        },
        fontSize: {
          type: "NUMBER",
          description: "Font size in points.",
        },
        fontColor: {
          type: "STRING",
          description: "Hex color such as '#B73618' or '#242424'.",
        },
        bold: {
          type: "BOOLEAN",
          description: "Bold formatting.",
        },
        italic: {
          type: "BOOLEAN",
          description: "Italic formatting.",
        },
        alignment: {
          type: "STRING",
          description: "Paragraph alignment: 'Left', 'Center', or 'Right'.",
          enum: ["Left", "Center", "Right"],
        },
        direction: {
          type: "STRING",
          description: "Text direction: 'rtl' or 'ltr'.",
          enum: ["rtl", "ltr"],
        },
      },
      required: ["slideNumber"],
    },
  },
  {
    name: "applySlideLayout",
    description:
      "Apply a slide layout or company template theme styling (heading font, body font, primary color) to a slide.",
    parameters: {
      type: "OBJECT",
      properties: {
        slideNumber: {
          type: "INTEGER",
          description: "1-based slide number.",
        },
        layoutName: {
          type: "STRING",
          description: "Layout name (e.g. 'Title and Content', 'Two Content').",
        },
        primaryColor: {
          type: "STRING",
          description: "Hex color for slide headings.",
        },
        headingFont: {
          type: "STRING",
          description: "Font family for slide headings.",
        },
        bodyFont: {
          type: "STRING",
          description: "Font family for body text.",
        },
      },
      required: ["slideNumber"],
    },
  },
];

export function isReadOnlyTool(toolName: string): boolean {
  return READ_ONLY_TOOLS.has(toolName);
}

export function requiresToolApproval(toolName: string): boolean {
  return !isReadOnlyTool(toolName);
}

export function getSupportedAgentTools(
  reqs: PowerPointApiRequirements
): GeminiFunctionDeclaration[] {
  return ALL_AGENT_TOOL_DECLARATIONS.filter((tool) => {
    if (!reqs.isPowerPointHost) return true;
    if (tool.name === "reorderSlide" && !reqs.pptApi1_6) return false;
    if (tool.name === "deleteSlide" && !reqs.pptApi1_2) return false;
    if (tool.name === "insertImage" && !reqs.imageCoercionSupported) return false;
    return true;
  });
}

export function describeToolStep(
  toolName: string,
  args: Record<string, unknown>
): string {
  const slideNum = args.slideNumber ? `Slide ${args.slideNumber}` : "Presentation";
  switch (toolName) {
    case "listSlides":
      return "List all slides and titles in the presentation";
    case "readSlide":
      return `Read text, shapes, and notes on ${slideNum}`;
    case "getSelectedText":
      return "Read currently selected text in PowerPoint";
    case "readSpeakerNotes":
      return `Read speaker notes on ${slideNum}`;
    case "writeSpeakerNotes":
      return `Write speaker notes on ${slideNum}`;
    case "addSlide":
      return `Add new slide "${args.title || "Untitled"}"${args.layoutName ? ` (${args.layoutName})` : ""}`;
    case "deleteSlide":
      return `Delete ${slideNum}`;
    case "reorderSlide":
      return `Reorder Slide ${args.fromSlideNumber} → position ${args.toSlideNumber}`;
    case "setSlideTitle":
      return `Set ${slideNum} title to "${args.title || ""}"`;
    case "insertTextBox":
      return `Insert text box on ${slideNum}: "${String(args.text || "").slice(0, 50)}..."`;
    case "replaceSlideText":
      return `Replace "${args.findText}" with "${args.replaceWith}" on ${args.slideNumber ? slideNum : "all slides"}`;
    case "insertTable":
      return `Insert table${args.title ? ` "${args.title}"` : ""} on ${slideNum}`;
    case "insertImage":
      return `Insert image onto ${slideNum}`;
    case "formatSlideText":
      return `Format text on ${slideNum} (${[
        args.fontName,
        args.fontSize ? `${args.fontSize}pt` : "",
        args.alignment,
        args.direction ? String(args.direction).toUpperCase() : "",
      ]
        .filter(Boolean)
        .join(", ")})`;
    case "applySlideLayout":
      return `Apply layout "${args.layoutName || "Company Template"}" to ${slideNum}`;
    default:
      return `${toolName}(${JSON.stringify(args).slice(0, 60)})`;
  }
}

export function validateToolArgs(
  toolName: string,
  args: Record<string, unknown>
): { valid: boolean; error?: string } {
  switch (toolName) {
    case "readSlide":
    case "readSpeakerNotes":
    case "deleteSlide":
    case "formatSlideText":
    case "applySlideLayout": {
      const n = Number(args.slideNumber);
      if (!Number.isFinite(n) || n < 1) {
        return { valid: false, error: `${toolName} requires a valid 1-based slideNumber >= 1.` };
      }
      return { valid: true };
    }
    case "writeSpeakerNotes": {
      const n = Number(args.slideNumber);
      if (!Number.isFinite(n) || n < 1) {
        return { valid: false, error: "writeSpeakerNotes requires slideNumber >= 1." };
      }
      if (typeof args.notes !== "string") {
        return { valid: false, error: "writeSpeakerNotes requires a string 'notes'." };
      }
      return { valid: true };
    }
    case "addSlide": {
      if (!args.title || typeof args.title !== "string") {
        return { valid: false, error: "addSlide requires a non-empty string 'title'." };
      }
      return { valid: true };
    }
    case "reorderSlide": {
      const from = Number(args.fromSlideNumber);
      const to = Number(args.toSlideNumber);
      if (!Number.isFinite(from) || from < 1 || !Number.isFinite(to) || to < 1) {
        return {
          valid: false,
          error: "reorderSlide requires fromSlideNumber >= 1 and toSlideNumber >= 1.",
        };
      }
      return { valid: true };
    }
    case "setSlideTitle": {
      if (!Number.isFinite(Number(args.slideNumber)) || Number(args.slideNumber) < 1) {
        return { valid: false, error: "setSlideTitle requires slideNumber >= 1." };
      }
      if (typeof args.title !== "string") {
        return { valid: false, error: "setSlideTitle requires a string 'title'." };
      }
      return { valid: true };
    }
    case "insertTextBox": {
      if (!Number.isFinite(Number(args.slideNumber)) || Number(args.slideNumber) < 1) {
        return { valid: false, error: "insertTextBox requires slideNumber >= 1." };
      }
      if (typeof args.text !== "string" || !args.text.trim()) {
        return { valid: false, error: "insertTextBox requires non-empty 'text'." };
      }
      return { valid: true };
    }
    case "replaceSlideText": {
      if (typeof args.findText !== "string" || !args.findText) {
        return { valid: false, error: "replaceSlideText requires non-empty 'findText'." };
      }
      if (typeof args.replaceWith !== "string") {
        return { valid: false, error: "replaceSlideText requires 'replaceWith'." };
      }
      return { valid: true };
    }
    case "insertTable": {
      if (!Number.isFinite(Number(args.slideNumber)) || Number(args.slideNumber) < 1) {
        return { valid: false, error: "insertTable requires slideNumber >= 1." };
      }
      if (!Array.isArray(args.headers)) {
        return { valid: false, error: "insertTable requires an array 'headers'." };
      }
      return { valid: true };
    }
    default:
      return { valid: true };
  }
}

export async function executeAgentToolStep(
  step: AgentToolStep,
  settings: ZeeSettings,
  attachments: ChatAttachment[] = []
): Promise<AgentToolStep> {
  const validation = validateToolArgs(step.toolName, step.args);
  if (!validation.valid) {
    return {
      ...step,
      status: "failed",
      errorMessage: validation.error,
    };
  }

  let workingStep: AgentToolStep = {
    ...step,
    status: "running",
    errorMessage: undefined,
  };

  // Capture undo snapshot before mutating tools
  if (workingStep.requiresApproval) {
    try {
      const activeSlide = Number(workingStep.args.slideNumber) || (await getActiveSlideIndex());
      if (workingStep.toolName === "addSlide") {
        const slides = await listSlides();
        workingStep.undoSnapshot = {
          stepId: workingStep.id,
          toolName: workingStep.toolName,
          timestamp: Date.now(),
          slideIndex: slides.length + 1,
          actionType: "add_slide",
        };
      } else if (workingStep.toolName === "deleteSlide") {
        workingStep.undoSnapshot = await captureSlideSnapshot(
          workingStep.id,
          workingStep.toolName,
          activeSlide,
          "delete_slide"
        );
      } else if (workingStep.toolName === "reorderSlide") {
        workingStep.undoSnapshot = {
          stepId: workingStep.id,
          toolName: workingStep.toolName,
          timestamp: Date.now(),
          slideIndex: Number(workingStep.args.fromSlideNumber) || 1,
          actionType: "reorder_slide",
          reorderFromIndex: Number(workingStep.args.fromSlideNumber) || 1,
          reorderToIndex: Number(workingStep.args.toSlideNumber) || 1,
        };
      } else {
        workingStep.undoSnapshot = await captureSlideSnapshot(
          workingStep.id,
          workingStep.toolName,
          activeSlide,
          "modify"
        );
      }
    } catch {
      // Proceed even if snapshot fails
    }
  }

  const runToolOnce = async (): Promise<string> => {
    const args = workingStep.args;
    switch (workingStep.toolName) {
      case "listSlides": {
        const slides = await listSlides();
        return JSON.stringify(slides);
      }
      case "readSlide": {
        const detail = await readSlide(Number(args.slideNumber));
        return JSON.stringify(detail);
      }
      case "getSelectedText": {
        const text = await getSelectedText();
        return text || "(No text selected)";
      }
      case "readSpeakerNotes": {
        const notes = await readSpeakerNotes(Number(args.slideNumber));
        return notes || "(No speaker notes on this slide)";
      }
      case "writeSpeakerNotes": {
        return await writeSpeakerNotes(
          Number(args.slideNumber),
          String(args.notes || ""),
          Boolean(args.append)
        );
      }
      case "addSlide": {
        const res = await addSlide({
          title: String(args.title || "New Slide"),
          bodyText: args.bodyText ? String(args.bodyText) : undefined,
          layoutName: args.layoutName ? String(args.layoutName) : undefined,
          notes: args.notes ? String(args.notes) : undefined,
          settings,
        });
        if (workingStep.undoSnapshot) {
          workingStep.undoSnapshot.slideIndex = res.slideIndex;
          workingStep.undoSnapshot.createdSlideId = res.slideId;
        }
        return res.summary;
      }
      case "deleteSlide": {
        return await deleteSlide(Number(args.slideNumber));
      }
      case "reorderSlide": {
        return await reorderSlide(
          Number(args.fromSlideNumber),
          Number(args.toSlideNumber)
        );
      }
      case "setSlideTitle": {
        return await setSlideTitle(
          Number(args.slideNumber),
          String(args.title || ""),
          settings
        );
      }
      case "insertTextBox": {
        return await insertTextBox({
          slideNumber: Number(args.slideNumber),
          text: String(args.text || ""),
          left: args.left !== undefined ? Number(args.left) : undefined,
          top: args.top !== undefined ? Number(args.top) : undefined,
          width: args.width !== undefined ? Number(args.width) : undefined,
          height: args.height !== undefined ? Number(args.height) : undefined,
          isHeading: Boolean(args.isHeading),
          settings,
        });
      }
      case "replaceSlideText": {
        return await replaceSlideText({
          slideNumber: args.slideNumber ? Number(args.slideNumber) : undefined,
          findText: String(args.findText || ""),
          replaceWith: String(args.replaceWith ?? ""),
          settings,
        });
      }
      case "insertTable": {
        let parsedRows: string[][] = [];
        if (typeof args.rowsJson === "string") {
          try {
            parsedRows = JSON.parse(args.rowsJson);
          } catch {
            parsedRows = [[args.rowsJson]];
          }
        } else if (Array.isArray(args.rows)) {
          parsedRows = args.rows as string[][];
        }
        const headers = Array.isArray(args.headers) ? args.headers.map(String) : [];
        return await insertTable({
          slideNumber: Number(args.slideNumber),
          title: args.title ? String(args.title) : undefined,
          headers,
          rows: parsedRows,
          settings,
        });
      }
      case "insertImage": {
        let base64 = String(args.base64Data || "");
        if (!base64 || base64 === "attachment") {
          const imgAtt = attachments.find((a) => a.kind === "image");
          if (imgAtt) base64 = imgAtt.base64Data;
        }
        return await insertImage({
          slideNumber: Number(args.slideNumber),
          base64Data: base64,
          left: args.left !== undefined ? Number(args.left) : undefined,
          top: args.top !== undefined ? Number(args.top) : undefined,
          width: args.width !== undefined ? Number(args.width) : undefined,
          height: args.height !== undefined ? Number(args.height) : undefined,
        });
      }
      case "formatSlideText": {
        return await formatSlideText({
          slideNumber: Number(args.slideNumber),
          shapeId: args.shapeId ? String(args.shapeId) : undefined,
          fontName: args.fontName ? String(args.fontName) : undefined,
          fontSize: args.fontSize !== undefined ? Number(args.fontSize) : undefined,
          fontColor: args.fontColor ? String(args.fontColor) : undefined,
          bold: args.bold !== undefined ? Boolean(args.bold) : undefined,
          italic: args.italic !== undefined ? Boolean(args.italic) : undefined,
          alignment: args.alignment as "Left" | "Center" | "Right" | undefined,
          direction: args.direction as "rtl" | "ltr" | undefined,
          settings,
        });
      }
      case "applySlideLayout": {
        return await applySlideLayout({
          slideNumber: Number(args.slideNumber),
          layoutName: args.layoutName ? String(args.layoutName) : undefined,
          primaryColor: args.primaryColor ? String(args.primaryColor) : undefined,
          headingFont: args.headingFont ? String(args.headingFont) : undefined,
          bodyFont: args.bodyFont ? String(args.bodyFont) : undefined,
          settings,
        });
      }
      default:
        throw new Error(`Unknown agent tool: ${workingStep.toolName}`);
    }
  };

  // Automatic transient retry (up to 1 automatic retry)
  for (let attempt = 0; attempt < 2; attempt++) {
    try {
      const output = await runToolOnce();
      return {
        ...workingStep,
        status: "completed",
        resultSummary: output,
        errorMessage: undefined,
      };
    } catch (err) {
      workingStep.retryCount = attempt + 1;
      if (attempt === 1) {
        return {
          ...workingStep,
          status: "failed",
          errorMessage: err instanceof Error ? err.message : String(err),
        };
      }
      await new Promise((r) => setTimeout(r, 250));
    }
  }

  return workingStep;
}

export async function undoCompletedAgentStep(step: AgentToolStep): Promise<AgentToolStep> {
  if (!step.undoSnapshot) {
    throw new Error("No undo snapshot recorded for this step.");
  }
  const summary = await restoreUndoSnapshot(step.undoSnapshot);
  return {
    ...step,
    status: "undone",
    resultSummary: summary,
  };
}

export interface AgentLoopSession {
  contents: GeminiTurnContent[];
  steps: AgentToolStep[];
  finalText: string;
  awaitingApprovalStep?: AgentToolStep;
  awaitingContinue: boolean;
  stepLimitReached: boolean;
}

export async function runAgentLoop(options: {
  settings: ZeeSettings;
  history: ChatMessage[];
  userPrompt: string;
  attachments?: ChatAttachment[];
  requirements: PowerPointApiRequirements;
  existingSession?: AgentLoopSession;
  maxAdditionalSteps?: number;
  onUpdate: (session: AgentLoopSession) => void;
  requestApproval: (step: AgentToolStep) => Promise<boolean>;
  signal?: AbortSignal;
}): Promise<AgentLoopSession> {
  const presSummary = await getPresentationSummary();
  const compactSummary = formatCompactPresentationSummary(presSummary);
  const systemInstruction =
    buildSystemInstruction(options.settings, compactSummary) +
    "\n\nAGENT MODE RULES:\n" +
    "- Use small, specific slide tools to inspect or modify the presentation.\n" +
    "- Read-only tools (listSlides, readSlide, getSelectedText, readSpeakerNotes) run immediately.\n" +
    "- Write tools require user approval via an approval card.\n" +
    "- After completing the requested slide actions, provide a clear summary of what was done.";

  const declarations = getSupportedAgentTools(options.requirements);
  const maxStepsThisRun = options.maxAdditionalSteps || options.settings.agentMaxSteps || 8;

  const session: AgentLoopSession = options.existingSession
    ? {
        ...options.existingSession,
        awaitingContinue: false,
        stepLimitReached: false,
      }
    : {
        contents: buildGeminiContents(
          options.history,
          options.userPrompt,
          options.attachments
        ),
        steps: [],
        finalText: "",
        awaitingContinue: false,
        stepLimitReached: false,
      };

  let executedInRun = 0;

  while (executedInRun < maxStepsThisRun) {
    const turnResult = await generateGeminiTurn(options.settings, session.contents, {
      systemInstruction,
      functionDeclarations: declarations,
      signal: options.signal,
    });

    if (turnResult.text) {
      session.finalText = turnResult.text;
    }

    if (turnResult.functionCalls.length === 0) {
      options.onUpdate({ ...session });
      return session;
    }

    session.contents.push({
      role: "model",
      parts: turnResult.rawModelParts,
    });

    const responseParts: {
      functionResponse: { name: string; response: Record<string, unknown> };
    }[] = [];

    for (const call of turnResult.functionCalls) {
      if (executedInRun >= maxStepsThisRun) {
        session.awaitingContinue = true;
        session.stepLimitReached = true;
        options.onUpdate({ ...session });
        return session;
      }

      const stepNumber = session.steps.length + 1;
      const needsApproval = requiresToolApproval(call.name);

      let step: AgentToolStep = {
        id: `step-${Date.now()}-${stepNumber}`,
        stepNumber,
        toolName: call.name,
        args: call.args || {},
        description: describeToolStep(call.name, call.args || {}),
        requiresApproval: needsApproval,
        status: needsApproval ? "pending_approval" : "running",
        retryCount: 0,
      };

      session.steps.push(step);
      options.onUpdate({ ...session, steps: [...session.steps] });

      if (needsApproval) {
        const approved = await options.requestApproval(step);
        if (!approved) {
          step = {
            ...step,
            status: "rejected",
            resultSummary: "Skipped / rejected by user.",
          };
          session.steps[session.steps.length - 1] = step;
          options.onUpdate({ ...session, steps: [...session.steps] });
          responseParts.push({
            functionResponse: {
              name: call.name,
              response: { status: "rejected", message: "User declined this action." },
            },
          });
          executedInRun += 1;
          continue;
        }
        step.status = "approved";
      }

      const finishedStep = await executeAgentToolStep(
        step,
        options.settings,
        options.attachments
      );
      session.steps[session.steps.length - 1] = finishedStep;
      options.onUpdate({ ...session, steps: [...session.steps] });

      responseParts.push({
        functionResponse: {
          name: call.name,
          response:
            finishedStep.status === "completed"
              ? { status: "ok", result: finishedStep.resultSummary }
              : { status: "error", error: finishedStep.errorMessage },
        },
      });

      executedInRun += 1;
    }

    session.contents.push({
      role: "user",
      parts: responseParts,
    });
  }

  session.awaitingContinue = true;
  session.stepLimitReached = true;
  options.onUpdate({ ...session });
  return session;
}
