import {
  PowerPointApiRequirements,
  PresentationSummary,
  ShapeInfo,
  SlideDetail,
  SlideSummary,
} from "../types/ppt";
import { SlideUndoSnapshot } from "../types/chat";
import { DEFAULT_SETTINGS, ZeeSettings } from "../types/settings";
import {
  containsArabic,
  maskProtectedTerms,
  resolveArabicFontFamily,
  resolveSlideTextAlignment,
  restoreProtectedTerms,
} from "./arabicService";

// In-memory fallback presentation for browser preview / unit testing outside PowerPoint
const mockPresentationSlides: SlideDetail[] = [
  {
    index: 1,
    id: "mock-slide-1",
    title: "North Field ESP Surveillance & Optimization",
    layoutName: "Title and Content",
    speakerNotes:
      "Welcome the reservoir and production engineering team. Emphasize WELL-04 PIP stability and TDH reduction.",
    fullText:
      "North Field ESP Surveillance & Optimization\n• Average field liquid rate: 4,850 bbl/d (WC 28%)\n• WELL-04 operating at PIP 1,120 psi, TDH 5,400 ft, GOR 420 scf/stb\n• Target BHP drawdown optimization to gain +320 stb/d",
    shapes: [
      {
        id: "shape-title-1",
        name: "Title 1",
        type: "GeometricShape",
        text: "North Field ESP Surveillance & Optimization",
        left: 48,
        top: 32,
        width: 620,
        height: 56,
        fontName: "Calibri Light",
        fontSize: 28,
        fontColor: "#B73618",
        isTitle: true,
      },
      {
        id: "shape-body-1",
        name: "Content Placeholder 2",
        type: "GeometricShape",
        text: "• Average field liquid rate: 4,850 bbl/d (WC 28%)\n• WELL-04 operating at PIP 1,120 psi, TDH 5,400 ft, GOR 420 scf/stb\n• Target BHP drawdown optimization to gain +320 stb/d",
        left: 48,
        top: 104,
        width: 620,
        height: 320,
        fontName: "Calibri",
        fontSize: 18,
        fontColor: "#242424",
        isTitle: false,
      },
    ],
  },
  {
    index: 2,
    id: "mock-slide-2",
    title: "Decline Curve Analysis — WELL-04 & WELL-12",
    layoutName: "Title and Content",
    speakerNotes:
      "Review hyperbolic Arps parameters (b=0.45, Di=14%/yr) and remaining EUR.",
    fullText:
      "Decline Curve Analysis — WELL-04 & WELL-12\n• WELL-04: Qi = 1,450 stb/d, Di = 14.2%/yr, b = 0.45, EUR = 2.85 MMstb\n• WELL-12: Qi = 980 stb/d, BHP = 1,640 psia, GOR = 510 scf/stb\n• Recommended ESP frequency increase from 54 Hz to 58 Hz",
    shapes: [
      {
        id: "shape-title-2",
        name: "Title 1",
        type: "GeometricShape",
        text: "Decline Curve Analysis — WELL-04 & WELL-12",
        left: 48,
        top: 32,
        width: 620,
        height: 56,
        fontName: "Calibri Light",
        fontSize: 28,
        fontColor: "#B73618",
        isTitle: true,
      },
      {
        id: "shape-body-2",
        name: "Content Placeholder 2",
        type: "GeometricShape",
        text: "• WELL-04: Qi = 1,450 stb/d, Di = 14.2%/yr, b = 0.45, EUR = 2.85 MMstb\n• WELL-12: Qi = 980 stb/d, BHP = 1,640 psia, GOR = 510 scf/stb\n• Recommended ESP frequency increase from 54 Hz to 58 Hz",
        left: 48,
        top: 104,
        width: 620,
        height: 320,
        fontName: "Arial",
        fontSize: 16,
        fontColor: "#242424",
        isTitle: false,
      },
    ],
  },
];

let mockActiveSlideIndex = 1;
let mockSelectedText = "";

function reindexMockSlides(): void {
  mockPresentationSlides.forEach((s, idx) => {
    s.index = idx + 1;
    s.fullText = s.shapes.map((sh) => sh.text).filter(Boolean).join("\n");
  });
}

export function isRunningInPowerPoint(): boolean {
  try {
    if (typeof window === "undefined") return false;
    const officeAny = (window as any).Office;
    const pptAny = (window as any).PowerPoint;
    if (
      officeAny?.context?.host === officeAny?.HostType?.PowerPoint ||
      String(officeAny?.context?.host || "").toLowerCase() === "powerpoint"
    ) {
      return true;
    }
    if (pptAny && typeof pptAny.run === "function" && officeAny?.context?.document) {
      return true;
    }
  } catch {
    // ignore
  }
  return false;
}

export function detectPowerPointRequirements(): PowerPointApiRequirements {
  const inPpt = isRunningInPowerPoint();
  const checkSet = (ver: string): boolean => {
    try {
      if (
        typeof Office !== "undefined" &&
        Office.context &&
        Office.context.requirements &&
        typeof Office.context.requirements.isSetSupported === "function"
      ) {
        return Office.context.requirements.isSetSupported("PowerPointApi", ver);
      }
    } catch {
      // ignore
    }
    // In browser preview mode, expose features so the user can preview UI
    return !inPpt;
  };

  const hasCommonDoc =
    !inPpt ||
    Boolean(
      typeof Office !== "undefined" &&
        Office.context?.document &&
        typeof Office.context.document.getSelectedDataAsync === "function"
    );

  return {
    isPowerPointHost: inPpt,
    pptApi1_1: checkSet("1.1"),
    pptApi1_2: checkSet("1.2"),
    pptApi1_3: checkSet("1.3"),
    pptApi1_4: checkSet("1.4"),
    pptApi1_5: checkSet("1.5"),
    pptApi1_6: checkSet("1.6"),
    pptApi1_7: checkSet("1.7"),
    pptApi1_8: checkSet("1.8"),
    selectedTextCommonApi: hasCommonDoc,
    imageCoercionSupported: hasCommonDoc,
    tableCoercionSupported: hasCommonDoc,
  };
}

/**
 * Repeatedly checks Office.onReady and polls host status so the add-in never
 * gets stuck in "Preview mode" if Office finishes initializing after first render.
 */
export function startOfficeReadinessMonitor(
  onStatusChange: (reqs: PowerPointApiRequirements) => void
): () => void {
  let cancelled = false;
  let lastHostState: boolean | null = null;

  const notifyIfChanged = () => {
    if (cancelled) return;
    const reqs = detectPowerPointRequirements();
    if (lastHostState !== reqs.isPowerPointHost) {
      lastHostState = reqs.isPowerPointHost;
      onStatusChange(reqs);
    }
  };

  const tryOnReady = () => {
    if (cancelled) return;
    try {
      if (typeof Office !== "undefined" && typeof Office.onReady === "function") {
        Office.onReady()
          .then(() => {
            notifyIfChanged();
          })
          .catch(() => {
            notifyIfChanged();
          });
      }
    } catch {
      // ignore
    }
    notifyIfChanged();
  };

  tryOnReady();

  // Check repeatedly at staggered intervals (fixes Word/PPT "Preview mode" banner bug)
  const delays = [150, 400, 900, 1800, 3200, 5500, 9000];
  const timers = delays.map((ms) => window.setTimeout(tryOnReady, ms));
  const periodic = window.setInterval(() => {
    tryOnReady();
    if (isRunningInPowerPoint()) {
      window.clearInterval(periodic);
    }
  }, 3000);

  return () => {
    cancelled = true;
    timers.forEach((t) => window.clearTimeout(t));
    window.clearInterval(periodic);
  };
}

export function getDocumentIdentifier(): string {
  try {
    if (typeof Office !== "undefined" && Office.context?.document?.url) {
      return Office.context.document.url;
    }
  } catch {
    // ignore
  }
  return "powerpoint-active-deck";
}

function parseSlideTitleAndPreview(shapes: ShapeInfo[]): {
  title: string;
  textPreview: string;
} {
  let title = "";
  const allTexts: string[] = [];

  for (const sh of shapes) {
    const clean = (sh.text || "").trim();
    if (!clean) continue;
    allTexts.push(clean);
    if (!title && (sh.isTitle || /title/i.test(sh.name))) {
      title = clean.split(/\r?\n/)[0].slice(0, 120);
    }
  }

  if (!title && allTexts.length > 0) {
    title = allTexts[0].split(/\r?\n/)[0].slice(0, 120);
  }

  const combined = allTexts.join(" | ").replace(/\s+/g, " ").trim();
  return {
    title: title || "Untitled Slide",
    textPreview: combined.slice(0, 220),
  };
}

export async function listSlides(): Promise<SlideSummary[]> {
  if (!isRunningInPowerPoint()) {
    return mockPresentationSlides.map((s) => ({
      index: s.index,
      id: s.id,
      title: s.title,
      textPreview: s.fullText.replace(/\s+/g, " ").slice(0, 200),
      shapeCount: s.shapes.length,
      layoutName: s.layoutName,
      hasNotes: Boolean(s.speakerNotes?.trim()),
    }));
  }

  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id");
    await context.sync();

    const loadedShapes: {
      slideId: string;
      shapes: PowerPoint.ShapeCollection;
    }[] = [];

    for (const slide of slides.items) {
      slide.shapes.load("items/id,items/name,items/type,items/left,items/top,items/width,items/height");
      loadedShapes.push({ slideId: slide.id, shapes: slide.shapes });
    }
    await context.sync();

    // Load textFrame safely per shape
    const textRanges: {
      slideIdx: number;
      shape: PowerPoint.Shape;
      textRange?: PowerPoint.TextRange;
    }[] = [];

    loadedShapes.forEach((entry, sIdx) => {
      for (const sh of entry.shapes.items) {
        try {
          sh.textFrame.load("hasText");
          sh.textFrame.textRange.load("text");
          textRanges.push({ slideIdx: sIdx, shape: sh, textRange: sh.textFrame.textRange });
        } catch {
          textRanges.push({ slideIdx: sIdx, shape: sh });
        }
      }
    });

    try {
      await context.sync();
    } catch {
      // Some non-text shapes (pictures/groups) throw when accessing textFrame; load individually if needed
    }

    const summaries: SlideSummary[] = [];
    for (let i = 0; i < loadedShapes.length; i++) {
      const entry = loadedShapes[i];
      const shapeInfos: ShapeInfo[] = [];
      for (const sh of entry.shapes.items) {
        let text = "";
        try {
          text = sh.textFrame?.textRange?.text || "";
        } catch {
          text = "";
        }
        shapeInfos.push({
          id: sh.id,
          name: sh.name || "",
          type: String(sh.type || "Shape"),
          text,
          left: sh.left || 0,
          top: sh.top || 0,
          width: sh.width || 0,
          height: sh.height || 0,
          isTitle: /title/i.test(sh.name || ""),
        });
      }
      const { title, textPreview } = parseSlideTitleAndPreview(shapeInfos);
      summaries.push({
        index: i + 1,
        id: entry.slideId,
        title,
        textPreview,
        shapeCount: shapeInfos.length,
      });
    }

    return summaries;
  });
}

export async function readSlide(slideNumber: number): Promise<SlideDetail> {
  if (!isRunningInPowerPoint()) {
    const idx = Math.max(1, Math.min(mockPresentationSlides.length, Math.round(slideNumber || 1))) - 1;
    const target = mockPresentationSlides[idx];
    if (!target) {
      throw new Error(`Slide ${slideNumber} not found.`);
    }
    return JSON.parse(JSON.stringify(target)) as SlideDetail;
  }

  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id");
    await context.sync();

    const zeroIdx = Math.max(0, Math.min(slides.items.length - 1, Math.round(slideNumber) - 1));
    const slide = slides.items[zeroIdx];
    if (!slide) {
      throw new Error(`Slide ${slideNumber} does not exist (deck has ${slides.items.length} slides).`);
    }

    slide.shapes.load("items/id,items/name,items/type,items/left,items/top,items/width,items/height");
    await context.sync();

    const shapeInfos: ShapeInfo[] = [];
    let notesFromHiddenShape = "";

    for (const sh of slide.shapes.items) {
      let text = "";
      let fontName: string | undefined;
      let fontSize: number | undefined;
      let fontColor: string | undefined;
      try {
        sh.textFrame.textRange.load("text,font/name,font/size,font/color");
        await context.sync();
        text = sh.textFrame.textRange.text || "";
        fontName = sh.textFrame.textRange.font?.name ?? undefined;
        fontSize = sh.textFrame.textRange.font?.size ?? undefined;
        fontColor = sh.textFrame.textRange.font?.color ?? undefined;
      } catch {
        // Non-text shape
      }

      if (sh.name === "ZeeAI_SpeakerNotes") {
        notesFromHiddenShape = text;
        continue;
      }

      shapeInfos.push({
        id: sh.id,
        name: sh.name || "Shape",
        type: String(sh.type || "Shape"),
        text,
        left: Math.round(sh.left || 0),
        top: Math.round(sh.top || 0),
        width: Math.round(sh.width || 0),
        height: Math.round(sh.height || 0),
        fontName,
        fontSize,
        fontColor,
        isTitle: /title/i.test(sh.name || "") || (shapeInfos.length === 0 && (sh.top || 0) < 100),
      });
    }

    let speakerNotes = notesFromHiddenShape;
    try {
      const slideAny = slide as any;
      if (slideAny.notesSlide?.notesTextFrame?.textRange) {
        slideAny.notesSlide.notesTextFrame.textRange.load("text");
        await context.sync();
        speakerNotes = slideAny.notesSlide.notesTextFrame.textRange.text || speakerNotes;
      }
    } catch {
      // Notes API not available on this PowerPoint version; fallback used
    }

    const { title } = parseSlideTitleAndPreview(shapeInfos);
    const fullText = shapeInfos
      .map((s) => s.text.trim())
      .filter(Boolean)
      .join("\n");

    return {
      index: zeroIdx + 1,
      id: slide.id,
      title,
      shapes: shapeInfos,
      speakerNotes,
      fullText,
    };
  });
}

export async function getActiveSlideIndex(): Promise<number> {
  if (!isRunningInPowerPoint()) {
    return mockActiveSlideIndex;
  }

  try {
    return await PowerPoint.run(async (context) => {
      const allSlides = context.presentation.slides;
      allSlides.load("items/id");

      let selectedId = "";
      try {
        const selectedSlides = context.presentation.getSelectedSlides();
        selectedSlides.load("items/id");
        await context.sync();
        selectedId = selectedSlides.items[0]?.id || "";
      } catch {
        await context.sync();
      }

      if (selectedId) {
        const foundIdx = allSlides.items.findIndex((s) => s.id === selectedId);
        if (foundIdx >= 0) return foundIdx + 1;
      }
      return 1;
    });
  } catch {
    return 1;
  }
}

export async function getSelectedText(): Promise<string> {
  if (!isRunningInPowerPoint()) {
    if (mockSelectedText) return mockSelectedText;
    const active = await readSlide(mockActiveSlideIndex);
    return active.fullText;
  }

  // First try PowerPointApi 1.5 getSelectedTextRangeOrNullObject
  try {
    const pptText = await PowerPoint.run(async (context) => {
      const presAny = context.presentation as any;
      if (typeof presAny.getSelectedTextRangeOrNullObject === "function") {
        const tr = presAny.getSelectedTextRangeOrNullObject();
        tr.load("isNullObject,text");
        await context.sync();
        if (!tr.isNullObject && tr.text) {
          return String(tr.text);
        }
      }
      if (typeof presAny.getSelectedShapes === "function") {
        const selShapes = presAny.getSelectedShapes();
        selShapes.load("items/id");
        await context.sync();
        if (selShapes.items.length > 0) {
          const texts: string[] = [];
          for (const sh of selShapes.items) {
            try {
              sh.textFrame.textRange.load("text");
              await context.sync();
              if (sh.textFrame.textRange.text) {
                texts.push(sh.textFrame.textRange.text);
              }
            } catch {
              // ignore
            }
          }
          if (texts.length > 0) return texts.join("\n");
        }
      }
      return "";
    });
    if (pptText.trim()) return pptText;
  } catch {
    // Fall back to Common API getSelectedDataAsync
  }

  return new Promise<string>((resolve) => {
    try {
      if (
        typeof Office !== "undefined" &&
        Office.context?.document &&
        typeof Office.context.document.getSelectedDataAsync === "function"
      ) {
        Office.context.document.getSelectedDataAsync(
          Office.CoercionType.Text,
          (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded && result.value) {
              resolve(String(result.value));
            } else {
              resolve("");
            }
          }
        );
        return;
      }
    } catch {
      // ignore
    }
    resolve("");
  });
}

export async function replaceSelectedText(
  newText: string,
  settings: ZeeSettings = DEFAULT_SETTINGS
): Promise<void> {
  if (!isRunningInPowerPoint()) {
    mockSelectedText = newText;
    const slide = mockPresentationSlides[mockActiveSlideIndex - 1];
    if (slide && slide.shapes.length > 0) {
      const bodyShape = slide.shapes[slide.shapes.length - 1];
      bodyShape.text = newText;
      reindexMockSlides();
    }
    return;
  }

  // Try Common API setSelectedDataAsync first if text is currently selected
  const replacedViaCommon = await new Promise<boolean>((resolve) => {
    try {
      if (
        typeof Office !== "undefined" &&
        Office.context?.document &&
        typeof Office.context.document.setSelectedDataAsync === "function"
      ) {
        Office.context.document.setSelectedDataAsync(
          newText,
          { coercionType: Office.CoercionType.Text },
          (res) => {
            resolve(res.status === Office.AsyncResultStatus.Succeeded);
          }
        );
        return;
      }
    } catch {
      // ignore
    }
    resolve(false);
  });

  if (replacedViaCommon) return;

  // Otherwise insert a formatted text box on the active slide
  const activeIdx = await getActiveSlideIndex();
  await insertTextBox({
    slideNumber: activeIdx,
    text: newText,
    settings,
  });
}

export async function readSpeakerNotes(slideNumber?: number): Promise<string> {
  const targetSlide = slideNumber || (await getActiveSlideIndex());
  const detail = await readSlide(targetSlide);
  return detail.speakerNotes || "";
}

export async function captureSlideSnapshot(
  stepId: string,
  toolName: string,
  slideNumber: number,
  actionType: SlideUndoSnapshot["actionType"] = "modify"
): Promise<SlideUndoSnapshot> {
  try {
    const detail = await readSlide(slideNumber);
    return {
      stepId,
      toolName,
      timestamp: Date.now(),
      slideIndex: detail.index,
      slideId: detail.id,
      actionType,
      previousTitle: detail.title,
      previousBodyTexts: detail.shapes.map((s) => ({
        shapeId: s.id,
        name: s.name,
        text: s.text,
      })),
      previousNotes: detail.speakerNotes,
      deletedSlideState:
        actionType === "delete_slide"
          ? {
              index: detail.index,
              title: detail.title,
              bodyTexts: detail.shapes.filter((s) => !s.isTitle).map((s) => s.text),
              notes: detail.speakerNotes,
              layoutName: detail.layoutName,
            }
          : undefined,
    };
  } catch {
    return {
      stepId,
      toolName,
      timestamp: Date.now(),
      slideIndex: slideNumber,
      actionType,
    };
  }
}

export async function restoreUndoSnapshot(snapshot: SlideUndoSnapshot): Promise<string> {
  if (snapshot.actionType === "add_slide") {
    if (snapshot.createdSlideId && isRunningInPowerPoint()) {
      await PowerPoint.run(async (context) => {
        const slides = context.presentation.slides;
        slides.load("items/id");
        await context.sync();
        const found = slides.items.find((s) => s.id === snapshot.createdSlideId);
        if (found) {
          found.delete();
          await context.sync();
        }
      });
      return `Undid addSlide (removed slide ${snapshot.slideIndex}).`;
    }
    await deleteSlide(snapshot.slideIndex);
    return `Undid addSlide (removed slide ${snapshot.slideIndex}).`;
  }

  if (snapshot.actionType === "delete_slide" && snapshot.deletedSlideState) {
    const st = snapshot.deletedSlideState;
    await addSlide({
      title: st.title,
      bodyText: st.bodyTexts.join("\n"),
      notes: st.notes,
      layoutName: st.layoutName,
    });
    return `Restored deleted slide "${st.title}".`;
  }

  if (
    snapshot.actionType === "reorder_slide" &&
    snapshot.reorderFromIndex &&
    snapshot.reorderToIndex
  ) {
    await reorderSlide(snapshot.reorderToIndex, snapshot.reorderFromIndex);
    return `Restored slide order (${snapshot.reorderToIndex} -> ${snapshot.reorderFromIndex}).`;
  }

  // Default: "modify" -> restore previous shape texts and notes
  if (!isRunningInPowerPoint()) {
    const slide = mockPresentationSlides[snapshot.slideIndex - 1];
    if (slide) {
      if (snapshot.previousTitle) slide.title = snapshot.previousTitle;
      if (snapshot.previousNotes !== undefined) slide.speakerNotes = snapshot.previousNotes;
      if (snapshot.previousBodyTexts) {
        for (const prev of snapshot.previousBodyTexts) {
          const found = slide.shapes.find((sh) => sh.id === prev.shapeId);
          if (found) found.text = prev.text;
        }
      }
      reindexMockSlides();
    }
    return `Restored slide ${snapshot.slideIndex} to its state before ${snapshot.toolName}.`;
  }

  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id");
    await context.sync();

    const slide =
      (snapshot.slideId && slides.items.find((s) => s.id === snapshot.slideId)) ||
      slides.items[Math.max(0, snapshot.slideIndex - 1)];
    if (!slide) return;

    slide.shapes.load("items/id,items/name");
    await context.sync();

    if (snapshot.previousBodyTexts) {
      for (const prev of snapshot.previousBodyTexts) {
        const sh = slide.shapes.items.find((item) => item.id === prev.shapeId);
        if (sh) {
          try {
            sh.textFrame.textRange.text = prev.text;
          } catch {
            // ignore
          }
        }
      }
      await context.sync();
    }
  });

  if (snapshot.previousNotes !== undefined) {
    await writeSpeakerNotes(snapshot.slideIndex, snapshot.previousNotes, false);
  }

  return `Restored slide ${snapshot.slideIndex} to its state before ${snapshot.toolName}.`;
}

export async function writeSpeakerNotes(
  slideNumber: number,
  notes: string,
  append: boolean = false
): Promise<string> {
  if (!isRunningInPowerPoint()) {
    const idx = Math.max(0, Math.min(mockPresentationSlides.length - 1, slideNumber - 1));
    const slide = mockPresentationSlides[idx];
    if (!slide) throw new Error(`Slide ${slideNumber} not found.`);
    slide.speakerNotes = append && slide.speakerNotes
      ? `${slide.speakerNotes}\n\n${notes}`
      : notes;
    return `Updated speaker notes on Slide ${slide.index}.`;
  }

  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id");
    await context.sync();

    const idx = Math.max(0, Math.min(slides.items.length - 1, slideNumber - 1));
    const slide = slides.items[idx];
    if (!slide) throw new Error(`Slide ${slideNumber} not found.`);

    const slideAny = slide as any;
    let wroteNativeNotes = false;

    try {
      if (slideAny.notesSlide?.notesTextFrame?.textRange) {
        if (append) {
          slideAny.notesSlide.notesTextFrame.textRange.load("text");
          await context.sync();
          const existing = slideAny.notesSlide.notesTextFrame.textRange.text || "";
          slideAny.notesSlide.notesTextFrame.textRange.text = existing
            ? `${existing}\n\n${notes}`
            : notes;
        } else {
          slideAny.notesSlide.notesTextFrame.textRange.text = notes;
        }
        await context.sync();
        wroteNativeNotes = true;
      }
    } catch {
      wroteNativeNotes = false;
    }

    if (!wroteNativeNotes) {
      // Fallback for older PowerPoint desktop versions without Slide.notesSlide API:
      // store in a compact footer callout / speaker notes box at the bottom margin
      slide.shapes.load("items/id,items/name");
      await context.sync();
      const existingBox = slide.shapes.items.find((s) => s.name === "ZeeAI_SpeakerNotes");
      if (existingBox) {
        existingBox.textFrame.textRange.text = notes;
      } else {
        const box = slide.shapes.addTextBox(notes, {
          left: 36,
          top: 495,
          width: 648,
          height: 36,
        });
        box.name = "ZeeAI_SpeakerNotes";
        box.textFrame.textRange.font.size = 9;
        box.textFrame.textRange.font.color = "#605E5C";
      }
      await context.sync();
    }

    return `Speaker notes ${append ? "appended" : "written"} on Slide ${idx + 1}.`;
  });
}

function applyDirectionAndFontToTextRange(
  textRange: PowerPoint.TextRange,
  text: string,
  isHeading: boolean,
  settings: ZeeSettings = DEFAULT_SETTINGS
): void {
  const { alignment } = resolveSlideTextAlignment(text, {
    directionMode: settings.directionMode,
    isHeading,
    centerHeadings: settings.centerHeadings,
  });

  try {
    const paraFormat = (textRange as any).paragraphFormat;
    if (paraFormat) {
      if (alignment === "Center") {
        paraFormat.horizontalAlignment = "Center";
      } else if (alignment === "Right") {
        paraFormat.horizontalAlignment = "Right";
      } else {
        paraFormat.horizontalAlignment = "Left";
      }
    }
  } catch {
    // paragraphFormat may not be supported on older requirement sets
  }

  const fallbackFont = isHeading
    ? settings.companyTemplate?.enabled
      ? settings.companyTemplate.headingFont
      : undefined
    : settings.companyTemplate?.enabled
      ? settings.companyTemplate.bodyFont
      : undefined;

  const fontName = resolveArabicFontFamily(text, settings.arabicFont, fallbackFont);
  if (fontName) {
    try {
      textRange.font.name = fontName;
    } catch {
      // ignore
    }
  }
}

export async function addSlide(options: {
  title?: string;
  bodyText?: string;
  bullets?: string[];
  layoutName?: string;
  notes?: string;
  settings?: ZeeSettings;
}): Promise<{ slideIndex: number; slideId: string; summary: string }> {
  const settings = options.settings || DEFAULT_SETTINGS;
  const titleText = (options.title || "New Slide").trim();
  const bodyContent = options.bullets?.length
    ? options.bullets.map((b) => (b.startsWith("•") ? b : `• ${b}`)).join("\n")
    : (options.bodyText || "").trim();

  const preferredLayout =
    options.layoutName ||
    (settings.companyTemplate?.enabled ? settings.companyTemplate.layoutName : "Title and Content");

  if (!isRunningInPowerPoint()) {
    const newIndex = mockPresentationSlides.length + 1;
    const newId = `mock-slide-${Date.now()}`;
    const shapes: ShapeInfo[] = [
      {
        id: `shape-title-${Date.now()}`,
        name: "Title 1",
        type: "GeometricShape",
        text: titleText,
        left: 48,
        top: 32,
        width: 620,
        height: 56,
        fontName:
          resolveArabicFontFamily(
            titleText,
            settings.arabicFont,
            settings.companyTemplate?.enabled ? settings.companyTemplate.headingFont : "Calibri Light"
          ) || "Calibri Light",
        fontSize: 28,
        fontColor: settings.companyTemplate?.enabled
          ? settings.companyTemplate.primaryColor
          : "#B73618",
        isTitle: true,
      },
    ];
    if (bodyContent) {
      shapes.push({
        id: `shape-body-${Date.now()}`,
        name: "Content Placeholder 2",
        type: "GeometricShape",
        text: bodyContent,
        left: 48,
        top: 104,
        width: 620,
        height: 320,
        fontName:
          resolveArabicFontFamily(
            bodyContent,
            settings.arabicFont,
            settings.companyTemplate?.enabled ? settings.companyTemplate.bodyFont : "Calibri"
          ) || "Calibri",
        fontSize: 18,
        fontColor: "#242424",
        isTitle: false,
      });
    }
    mockPresentationSlides.push({
      index: newIndex,
      id: newId,
      title: titleText,
      layoutName: preferredLayout,
      speakerNotes: options.notes || "",
      fullText: [titleText, bodyContent].filter(Boolean).join("\n"),
      shapes,
    });
    mockActiveSlideIndex = newIndex;
    return {
      slideIndex: newIndex,
      slideId: newId,
      summary: `Added Slide ${newIndex}: "${titleText}" (${preferredLayout}).`,
    };
  }

  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;

    // If PowerPointApi 1.3 is supported, try matching the requested layout from SlideMaster
    let addedSlide: PowerPoint.Slide | undefined;
    try {
      const masters = context.presentation.slideMasters;
      masters.load("items/id,items/layouts/items/id,items/layouts/items/name");
      await context.sync();
      if (masters.items.length > 0) {
        const layouts = masters.items[0].layouts.items;
        const matchedLayout = layouts.find(
          (l) => l.name.toLowerCase() === preferredLayout.toLowerCase()
        );
        if (matchedLayout) {
          slides.add({ layoutId: matchedLayout.id });
          await context.sync();
        }
      }
    } catch {
      // Fallback to default slides.add()
    }

    if (!addedSlide) {
      slides.add();
      await context.sync();
    }

    slides.load("items/id");
    await context.sync();

    const newIndex = slides.items.length;
    const slide = slides.items[newIndex - 1];
    slide.shapes.load("items/id,items/name");
    await context.sync();

    // Populate or add Title box
    const titleBox = slide.shapes.addTextBox(titleText, {
      left: 48,
      top: 32,
      width: 624,
      height: 58,
    });
    titleBox.name = "Title 1";
    titleBox.textFrame.textRange.font.size = 26;
    titleBox.textFrame.textRange.font.bold = true;
    titleBox.textFrame.textRange.font.color = settings.companyTemplate?.enabled
      ? settings.companyTemplate.primaryColor
      : "#B73618";
    applyDirectionAndFontToTextRange(titleBox.textFrame.textRange, titleText, true, settings);

    if (bodyContent) {
      const bodyBox = slide.shapes.addTextBox(bodyContent, {
        left: 48,
        top: 104,
        width: 624,
        height: 340,
      });
      bodyBox.name = "Content Placeholder 2";
      bodyBox.textFrame.textRange.font.size = 16;
      bodyBox.textFrame.textRange.font.color = "#242424";
      applyDirectionAndFontToTextRange(bodyBox.textFrame.textRange, bodyContent, false, settings);
    }

    if (settings.companyTemplate?.enabled && settings.companyTemplate.footerText) {
      const footerBox = slide.shapes.addTextBox(settings.companyTemplate.footerText, {
        left: 48,
        top: 498,
        width: 624,
        height: 24,
      });
      footerBox.name = "Footer Placeholder";
      footerBox.textFrame.textRange.font.size = 9;
      footerBox.textFrame.textRange.font.color = "#605E5C";
    }

    await context.sync();

    if (options.notes) {
      await writeSpeakerNotes(newIndex, options.notes, false);
    }

    return {
      slideIndex: newIndex,
      slideId: slide.id,
      summary: `Added Slide ${newIndex}: "${titleText}"`,
    };
  });
}

export async function deleteSlide(slideNumber: number): Promise<string> {
  if (!isRunningInPowerPoint()) {
    const idx = slideNumber - 1;
    if (idx < 0 || idx >= mockPresentationSlides.length) {
      throw new Error(`Slide ${slideNumber} does not exist.`);
    }
    const removed = mockPresentationSlides.splice(idx, 1)[0];
    reindexMockSlides();
    mockActiveSlideIndex = Math.max(1, Math.min(mockActiveSlideIndex, mockPresentationSlides.length));
    return `Deleted Slide ${slideNumber} ("${removed.title}").`;
  }

  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id");
    await context.sync();

    const idx = slideNumber - 1;
    if (idx < 0 || idx >= slides.items.length) {
      throw new Error(`Slide ${slideNumber} out of range (1..${slides.items.length}).`);
    }

    slides.items[idx].delete();
    await context.sync();
    return `Deleted Slide ${slideNumber}.`;
  });
}

export async function reorderSlide(
  fromSlideNumber: number,
  toSlideNumber: number
): Promise<string> {
  if (!isRunningInPowerPoint()) {
    const fromIdx = fromSlideNumber - 1;
    const toIdx = Math.max(0, Math.min(mockPresentationSlides.length - 1, toSlideNumber - 1));
    if (fromIdx < 0 || fromIdx >= mockPresentationSlides.length) {
      throw new Error(`Slide ${fromSlideNumber} out of range.`);
    }
    const [moved] = mockPresentationSlides.splice(fromIdx, 1);
    mockPresentationSlides.splice(toIdx, 0, moved);
    reindexMockSlides();
    return `Moved Slide ${fromSlideNumber} ("${moved.title}") to position ${toIdx + 1}.`;
  }

  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id");
    await context.sync();

    const fromIdx = fromSlideNumber - 1;
    const toIdx = Math.max(0, Math.min(slides.items.length - 1, toSlideNumber - 1));
    const slide = slides.items[fromIdx] as any;
    if (!slide) {
      throw new Error(`Slide ${fromSlideNumber} not found.`);
    }

    if (typeof slide.moveTo === "function") {
      slide.moveTo(toIdx);
      await context.sync();
      return `Moved Slide ${fromSlideNumber} to position ${toIdx + 1}.`;
    }

    throw new Error("Slide reordering requires PowerPointApi 1.6 or newer.");
  });
}

export async function setSlideTitle(
  slideNumber: number,
  title: string,
  settings: ZeeSettings = DEFAULT_SETTINGS
): Promise<string> {
  if (!isRunningInPowerPoint()) {
    const slide = mockPresentationSlides[slideNumber - 1];
    if (!slide) throw new Error(`Slide ${slideNumber} not found.`);
    slide.title = title;
    const titleShape = slide.shapes.find((s) => s.isTitle || /title/i.test(s.name));
    if (titleShape) {
      titleShape.text = title;
    } else {
      slide.shapes.unshift({
        id: `shape-title-${Date.now()}`,
        name: "Title 1",
        type: "GeometricShape",
        text: title,
        left: 48,
        top: 32,
        width: 620,
        height: 56,
        isTitle: true,
      });
    }
    reindexMockSlides();
    return `Updated title on Slide ${slideNumber} to "${title}".`;
  }

  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id");
    await context.sync();

    const slide = slides.items[slideNumber - 1];
    if (!slide) throw new Error(`Slide ${slideNumber} not found.`);

    slide.shapes.load("items/id,items/name,items/top");
    await context.sync();

    const existingTitle = slide.shapes.items.find(
      (s) => /title/i.test(s.name || "") || (s.top !== undefined && s.top < 95)
    );

    if (existingTitle) {
      existingTitle.textFrame.textRange.text = title;
      applyDirectionAndFontToTextRange(existingTitle.textFrame.textRange, title, true, settings);
    } else {
      const box = slide.shapes.addTextBox(title, {
        left: 48,
        top: 32,
        width: 624,
        height: 58,
      });
      box.name = "Title 1";
      box.textFrame.textRange.font.size = 26;
      box.textFrame.textRange.font.bold = true;
      applyDirectionAndFontToTextRange(box.textFrame.textRange, title, true, settings);
    }

    await context.sync();
    return `Set Slide ${slideNumber} title to "${title}".`;
  });
}

export async function insertTextBox(options: {
  slideNumber: number;
  text: string;
  left?: number;
  top?: number;
  width?: number;
  height?: number;
  isHeading?: boolean;
  fontSize?: number;
  fontColor?: string;
  settings?: ZeeSettings;
}): Promise<string> {
  const settings = options.settings || DEFAULT_SETTINGS;
  const left = options.left ?? 48;
  const top = options.top ?? 110;
  const width = options.width ?? 624;
  const height = options.height ?? 260;

  if (!isRunningInPowerPoint()) {
    const slide = mockPresentationSlides[options.slideNumber - 1];
    if (!slide) throw new Error(`Slide ${options.slideNumber} not found.`);
    slide.shapes.push({
      id: `shape-tb-${Date.now()}`,
      name: options.isHeading ? "Heading Box" : "TextBox",
      type: "GeometricShape",
      text: options.text,
      left,
      top,
      width,
      height,
      fontName: resolveArabicFontFamily(options.text, settings.arabicFont, "Calibri") || "Calibri",
      fontSize: options.fontSize || (options.isHeading ? 24 : 16),
      fontColor: options.fontColor || "#242424",
      isTitle: Boolean(options.isHeading),
    });
    reindexMockSlides();
    return `Inserted text box on Slide ${options.slideNumber}.`;
  }

  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id");
    await context.sync();

    const slide = slides.items[options.slideNumber - 1];
    if (!slide) throw new Error(`Slide ${options.slideNumber} not found.`);

    const box = slide.shapes.addTextBox(options.text, {
      left,
      top,
      width,
      height,
    });
    box.textFrame.textRange.font.size = options.fontSize || (options.isHeading ? 24 : 16);
    if (options.fontColor) {
      box.textFrame.textRange.font.color = options.fontColor;
    }
    if (options.isHeading) {
      box.textFrame.textRange.font.bold = true;
    }
    applyDirectionAndFontToTextRange(
      box.textFrame.textRange,
      options.text,
      Boolean(options.isHeading),
      settings
    );
    await context.sync();

    return `Inserted text box on Slide ${options.slideNumber}.`;
  });
}

export async function replaceSlideText(options: {
  slideNumber?: number; // omit or 0 for all slides
  findText: string;
  replaceWith: string;
  matchCase?: boolean;
  respectProtectedTerms?: boolean;
  settings?: ZeeSettings;
}): Promise<string> {
  const settings = options.settings || DEFAULT_SETTINGS;
  if (!options.findText) {
    throw new Error("findText cannot be empty.");
  }

  const performReplaceOnString = (source: string): { updated: string; count: number } => {
    let working = source;
    let placeholders: Record<string, string> = {};
    if (options.respectProtectedTerms ?? settings.protectTermsEnabled) {
      const masked = maskProtectedTerms(source, settings.customProtectedTerms);
      // Only mask terms that are NOT explicitly the target of an intentional unit/term fix
      if (!masked.protectedValues.includes(options.findText)) {
        working = masked.maskedText;
        placeholders = masked.placeholders;
      }
    }

    const flags = options.matchCase ? "g" : "gi";
    const escaped = options.findText.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const regex = new RegExp(escaped, flags);
    let count = 0;
    const replaced = working.replace(regex, () => {
      count += 1;
      return options.replaceWith;
    });

    return {
      updated: restoreProtectedTerms(replaced, placeholders),
      count,
    };
  };

  if (!isRunningInPowerPoint()) {
    let totalCount = 0;
    const targets = options.slideNumber
      ? mockPresentationSlides.filter((s) => s.index === options.slideNumber)
      : mockPresentationSlides;

    for (const sl of targets) {
      for (const sh of sl.shapes) {
        const { updated, count } = performReplaceOnString(sh.text);
        if (count > 0) {
          sh.text = updated;
          totalCount += count;
        }
      }
      const titleCandidate = sl.shapes.find((s) => s.isTitle);
      if (titleCandidate) sl.title = titleCandidate.text;
    }
    reindexMockSlides();
    return `Replaced ${totalCount} occurrence(s) of "${options.findText}" with "${options.replaceWith}".`;
  }

  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id");
    await context.sync();

    const targetSlides = options.slideNumber
      ? [slides.items[options.slideNumber - 1]].filter(Boolean)
      : slides.items;

    let totalReplacements = 0;
    for (const slide of targetSlides) {
      slide.shapes.load("items/id,items/name");
      await context.sync();
      for (const sh of slide.shapes.items) {
        try {
          sh.textFrame.textRange.load("text");
          await context.sync();
          const current = sh.textFrame.textRange.text || "";
          const { updated, count } = performReplaceOnString(current);
          if (count > 0) {
            sh.textFrame.textRange.text = updated;
            totalReplacements += count;
          }
        } catch {
          // Non-text shape
        }
      }
      await context.sync();
    }

    return `Replaced ${totalReplacements} occurrence(s) of "${options.findText}" with "${options.replaceWith}".`;
  });
}

export async function insertTable(options: {
  slideNumber: number;
  headers: string[];
  rows: string[][];
  title?: string;
  left?: number;
  top?: number;
  width?: number;
  height?: number;
  settings?: ZeeSettings;
}): Promise<string> {
  const headers = options.headers || [];
  const rows = options.rows || [];
  const allRows = headers.length > 0 ? [headers, ...rows] : rows;

  const formattedGridText = allRows
    .map((r, idx) => {
      const line = r.join("  |  ");
      if (idx === 0 && headers.length > 0) {
        return `${line}\n${"-".repeat(Math.min(72, line.length + 6))}`;
      }
      return line;
    })
    .join("\n");

  if (!isRunningInPowerPoint()) {
    const slide = mockPresentationSlides[options.slideNumber - 1];
    if (!slide) throw new Error(`Slide ${options.slideNumber} not found.`);
    slide.shapes.push({
      id: `shape-table-${Date.now()}`,
      name: options.title ? `Table: ${options.title}` : "Data Table",
      type: "Table",
      text: formattedGridText,
      left: options.left ?? 48,
      top: options.top ?? 140,
      width: options.width ?? 624,
      height: options.height ?? 240,
      fontName: "Consolas",
      fontSize: 13,
    });
    reindexMockSlides();
    return `Inserted ${allRows.length}x${headers.length || rows[0]?.length || 1} table on Slide ${options.slideNumber}.`;
  }

  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id");
    await context.sync();

    const slide = slides.items[options.slideNumber - 1];
    if (!slide) throw new Error(`Slide ${options.slideNumber} not found.`);

    const shapesAny = slide.shapes as any;
    if (typeof shapesAny.addTable === "function" && allRows.length > 0) {
      try {
        shapesAny.addTable(allRows.length, allRows[0].length, {
          left: options.left ?? 48,
          top: options.top ?? 140,
          width: options.width ?? 624,
          height: options.height ?? 240,
          values: allRows,
        });
        await context.sync();
        return `Inserted native PowerPoint table (${allRows.length} rows) on Slide ${options.slideNumber}.`;
      } catch {
        // Fall through to structured table box
      }
    }

    // Try Office.CoercionType.Table if active slide matches
    const activeIdx = await getActiveSlideIndex();
    if (activeIdx === options.slideNumber && typeof Office !== "undefined" && Office.context?.document?.setSelectedDataAsync) {
      const tableData = new Office.TableData();
      tableData.headers = [headers];
      tableData.rows = rows;
      const insertedViaCommon = await new Promise<boolean>((resolve) => {
        Office.context.document.setSelectedDataAsync(
          tableData,
          { coercionType: Office.CoercionType.Table },
          (res) => resolve(res.status === Office.AsyncResultStatus.Succeeded)
        );
      });
      if (insertedViaCommon) {
        return `Inserted table (${rows.length} data rows) on Slide ${options.slideNumber}.`;
      }
    }

    // Structured fallback box so table insertion works on every PowerPoint version
    const box = slide.shapes.addTextBox(formattedGridText, {
      left: options.left ?? 48,
      top: options.top ?? 140,
      width: options.width ?? 624,
      height: options.height ?? 240,
    });
    box.name = options.title ? `Table: ${options.title}` : "ZeeAI_Table";
    box.textFrame.textRange.font.name = "Consolas";
    box.textFrame.textRange.font.size = 13;
    await context.sync();

    return `Inserted formatted table (${allRows.length} rows) on Slide ${options.slideNumber}.`;
  });
}

export async function insertImage(options: {
  slideNumber: number;
  base64Data: string;
  left?: number;
  top?: number;
  width?: number;
  height?: number;
}): Promise<string> {
  const cleanBase64 = options.base64Data.replace(/^data:image\/[a-zA-Z0-9+.-]+;base64,/, "");

  if (!isRunningInPowerPoint()) {
    const slide = mockPresentationSlides[options.slideNumber - 1];
    if (!slide) throw new Error(`Slide ${options.slideNumber} not found.`);
    slide.shapes.push({
      id: `shape-img-${Date.now()}`,
      name: "Inserted Image",
      type: "Image",
      text: "[Image Attached]",
      left: options.left ?? 380,
      top: options.top ?? 120,
      width: options.width ?? 280,
      height: options.height ?? 200,
    });
    reindexMockSlides();
    return `Inserted image on Slide ${options.slideNumber}.`;
  }

  return new Promise<string>((resolve, reject) => {
    try {
      if (
        typeof Office !== "undefined" &&
        Office.context?.document &&
        typeof Office.context.document.setSelectedDataAsync === "function"
      ) {
        Office.context.document.setSelectedDataAsync(
          cleanBase64,
          {
            coercionType: Office.CoercionType.Image,
            imageLeft: options.left ?? 360,
            imageTop: options.top ?? 120,
            imageWidth: options.width ?? 300,
            imageHeight: options.height ?? 210,
          },
          (res) => {
            if (res.status === Office.AsyncResultStatus.Succeeded) {
              resolve(`Inserted image onto Slide ${options.slideNumber}.`);
            } else {
              reject(new Error(res.error?.message || "Failed to insert image."));
            }
          }
        );
        return;
      }
      reject(new Error("Image insertion API is not available in this host."));
    } catch (err) {
      reject(err instanceof Error ? err : new Error(String(err)));
    }
  });
}

export async function formatSlideText(options: {
  slideNumber: number;
  shapeId?: string;
  fontName?: string;
  fontSize?: number;
  fontColor?: string;
  bold?: boolean;
  italic?: boolean;
  alignment?: "Left" | "Center" | "Right";
  direction?: "rtl" | "ltr";
  settings?: ZeeSettings;
}): Promise<string> {
  const settings = options.settings || DEFAULT_SETTINGS;

  if (!isRunningInPowerPoint()) {
    const slide = mockPresentationSlides[options.slideNumber - 1];
    if (!slide) throw new Error(`Slide ${options.slideNumber} not found.`);
    const shapes = options.shapeId
      ? slide.shapes.filter((s) => s.id === options.shapeId)
      : slide.shapes;

    for (const sh of shapes) {
      if (options.fontName) sh.fontName = options.fontName;
      if (options.fontSize) sh.fontSize = options.fontSize;
      if (options.fontColor) sh.fontColor = options.fontColor;
    }
    return `Formatted ${shapes.length} shape(s) on Slide ${options.slideNumber}.`;
  }

  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id");
    await context.sync();

    const slide = slides.items[options.slideNumber - 1];
    if (!slide) throw new Error(`Slide ${options.slideNumber} not found.`);

    slide.shapes.load("items/id,items/name");
    await context.sync();

    const targetShapes = options.shapeId
      ? slide.shapes.items.filter((s) => s.id === options.shapeId)
      : slide.shapes.items;

    let formattedCount = 0;
    for (const sh of targetShapes) {
      try {
        const tr = sh.textFrame.textRange;
        tr.load("text");
        await context.sync();

        const effectiveFont =
          options.fontName ||
          (containsArabic(tr.text) && settings.arabicFont !== "Same as document"
            ? settings.arabicFont
            : undefined);

        if (effectiveFont) tr.font.name = effectiveFont;
        if (options.fontSize && options.fontSize > 0) tr.font.size = options.fontSize;
        if (options.fontColor) tr.font.color = options.fontColor;
        if (options.bold !== undefined) tr.font.bold = options.bold;
        if (options.italic !== undefined) tr.font.italic = options.italic;

        const paraFormat = (tr as any).paragraphFormat;
        if (paraFormat) {
          if (options.alignment) {
            paraFormat.horizontalAlignment = options.alignment;
          } else if (options.direction === "rtl") {
            paraFormat.horizontalAlignment = "Right";
          } else if (options.direction === "ltr") {
            paraFormat.horizontalAlignment = "Left";
          }
        }
        formattedCount += 1;
      } catch {
        // non-text shape
      }
    }

    await context.sync();
    return `Formatted ${formattedCount} text shape(s) on Slide ${options.slideNumber}.`;
  });
}

export async function applySlideLayout(options: {
  slideNumber: number;
  layoutName?: string;
  primaryColor?: string;
  headingFont?: string;
  bodyFont?: string;
  footerText?: string;
  base64Pptx?: string;
  settings?: ZeeSettings;
}): Promise<string> {
  const settings = options.settings || DEFAULT_SETTINGS;
  const tpl = settings.companyTemplate;
  const primaryColor = options.primaryColor || tpl.primaryColor || "#B73618";
  const headingFont = options.headingFont || tpl.headingFont || "Calibri Light";
  const bodyFont = options.bodyFont || tpl.bodyFont || "Calibri";

  if (!isRunningInPowerPoint()) {
    const slide = mockPresentationSlides[options.slideNumber - 1];
    if (!slide) throw new Error(`Slide ${options.slideNumber} not found.`);
    if (options.layoutName) slide.layoutName = options.layoutName;
    for (const sh of slide.shapes) {
      if (sh.isTitle) {
        sh.fontName = headingFont;
        sh.fontColor = primaryColor;
      } else {
        sh.fontName = bodyFont;
      }
    }
    return `Applied layout "${options.layoutName || tpl.layoutName}" and styling to Slide ${options.slideNumber}.`;
  }

  return PowerPoint.run(async (context) => {
    if (options.base64Pptx) {
      const presAny = context.presentation as any;
      if (typeof presAny.insertSlidesFromBase64 === "function") {
        presAny.insertSlidesFromBase64(options.base64Pptx, {
          formatting: "UseDestinationTheme",
        });
        await context.sync();
        return "Inserted slide(s) from Base64 PowerPoint template.";
      }
    }

    const slides = context.presentation.slides;
    slides.load("items/id");
    await context.sync();

    const slide = slides.items[options.slideNumber - 1];
    if (!slide) throw new Error(`Slide ${options.slideNumber} not found.`);

    slide.shapes.load("items/id,items/name,items/top");
    await context.sync();

    for (const sh of slide.shapes.items) {
      try {
        const isTitle = /title/i.test(sh.name || "") || (sh.top !== undefined && sh.top < 95);
        sh.textFrame.textRange.load("text");
        await context.sync();
        const text = sh.textFrame.textRange.text || "";
        const arabicFont = resolveArabicFontFamily(
          text,
          settings.arabicFont,
          isTitle ? headingFont : bodyFont
        );
        if (arabicFont) {
          sh.textFrame.textRange.font.name = arabicFont;
        }
        if (isTitle) {
          sh.textFrame.textRange.font.color = primaryColor;
          sh.textFrame.textRange.font.bold = true;
        }
        applyDirectionAndFontToTextRange(sh.textFrame.textRange, text, isTitle, settings);
      } catch {
        // non-text shape
      }
    }

    await context.sync();
    return `Applied template styling (${headingFont} / ${bodyFont}, ${primaryColor}) to Slide ${options.slideNumber}.`;
  });
}

export async function listAvailableLayouts(): Promise<string[]> {
  const defaults = [
    "Title Slide",
    "Title and Content",
    "Section Header",
    "Two Content",
    "Comparison",
    "Title Only",
    "Blank",
  ];

  if (!isRunningInPowerPoint()) {
    return defaults;
  }

  try {
    return await PowerPoint.run(async (context) => {
      const masters = context.presentation.slideMasters;
      masters.load("items/layouts/items/name");
      await context.sync();
      const names = new Set<string>(defaults);
      for (const m of masters.items) {
        for (const l of m.layouts.items) {
          if (l.name) names.add(l.name);
        }
      }
      return Array.from(names);
    });
  } catch {
    return defaults;
  }
}

export async function getPresentationSummary(): Promise<PresentationSummary> {
  const slides = await listSlides();
  const activeSlideIndex = await getActiveSlideIndex();
  let activeSlideDetail: SlideDetail | undefined;
  try {
    if (slides.length > 0) {
      activeSlideDetail = await readSlide(activeSlideIndex);
    }
  } catch {
    // ignore
  }
  const selectedText = await getSelectedText();
  const availableLayouts = await listAvailableLayouts();

  return {
    documentId: getDocumentIdentifier(),
    slideCount: slides.length,
    activeSlideIndex,
    slides,
    activeSlideDetail,
    selectedText: selectedText || undefined,
    availableLayouts,
  };
}

/**
 * Builds a compact summary of the presentation (slide count, titles, active slide content)
 * instead of dumping the full deck so prompts remain fast and cheap.
 */
export function formatCompactPresentationSummary(summary: PresentationSummary): string {
  const lines: string[] = [
    `Slide Count: ${summary.slideCount} | Active Slide: #${summary.activeSlideIndex}`,
    "Slide Index & Titles:",
    ...summary.slides.map(
      (s) =>
        `  - Slide ${s.index}: "${s.title}" (${s.shapeCount} shapes) — ${s.textPreview.slice(0, 100)}`
    ),
  ];

  if (summary.activeSlideDetail) {
    lines.push(
      `\nActive Slide #${summary.activeSlideDetail.index} Full Text:\n${summary.activeSlideDetail.fullText || "(empty)"}`
    );
    if (summary.activeSlideDetail.speakerNotes) {
      lines.push(`Active Slide Speaker Notes: ${summary.activeSlideDetail.speakerNotes}`);
    }
  }

  if (summary.selectedText) {
    lines.push(`\nCurrently Selected Text:\n"${summary.selectedText}"`);
  }

  return lines.join("\n");
}
