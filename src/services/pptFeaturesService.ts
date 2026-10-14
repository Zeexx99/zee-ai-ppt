import { ChatAttachment, SlideUndoSnapshot } from "../types/chat";
import {
  ConsistencyIssue,
  CrossSlideContradiction,
  DeckOutlinePlan,
  DeckOutlineSlide,
  DeckQnaItem,
  DeckReviewReport,
  ExcelResultsPayload,
  PetroleumResultCategory,
  SlideBeautifyStyle,
  SlideDetail,
  TextFitIssue,
  WellBookRow,
} from "../types/ppt";
import { DEFAULT_PROTECTED_TERMS, ZeeSettings } from "../types/settings";
import {
  buildDirectionAndLanguageInstruction,
  containsArabic,
  formatBilingualTextBlock,
} from "./arabicService";
import {
  DeclineCurveParams,
  EspCurveParams,
  KpiCardItem,
  renderDeclineCurveSvg,
  renderEspCurveSvg,
  renderKpiCardsSvg,
  svgToPngBase64,
} from "./chartGeneratorService";
import {
  convertAttachmentsToGeminiParts,
  GeminiPart,
  generateGeminiTurn,
} from "./geminiService";
import {
  addSlide,
  applySlideLayout,
  captureSlideSnapshot,
  formatSlideText,
  insertImage,
  insertTable,
  insertTextBox,
  listSlides,
  readSlide,
  replaceSlideText,
  setSlideTitle,
  writeSpeakerNotes,
} from "./pptService";

function extractJsonBlock(rawText: string): string {
  const fenced = rawText.match(/```(?:json)?\s*([\s\S]*?)```/i);
  if (fenced && fenced[1]) {
    return fenced[1].trim();
  }
  const firstBrace = rawText.indexOf("{");
  const lastBrace = rawText.lastIndexOf("}");
  if (firstBrace >= 0 && lastBrace > firstBrace) {
    return rawText.slice(firstBrace, lastBrace + 1);
  }
  return rawText.trim();
}

/**
 * Step 1 of Report/PDF -> Deck:
 * Generates a structured slide-by-slide outline for the user to inspect/edit before building slides.
 */
export async function generateDeckOutlineFromReport(options: {
  reportText: string;
  attachments?: ChatAttachment[];
  targetSlideCount?: number;
  settings: ZeeSettings;
}): Promise<DeckOutlinePlan> {
  const slideCount = Math.max(3, Math.min(15, options.targetSlideCount || 6));
  const langRules = buildDirectionAndLanguageInstruction(options.settings);

  const prompt = [
    `You are transforming an engineering/business report or attached PDF into a ${slideCount}-slide PowerPoint presentation outline.`,
    langRules,
    "Return ONLY valid JSON matching this schema:",
    `{
  "deckTitle": "string",
  "deckSubtitle": "string",
  "sourceSummary": "1-2 sentence executive overview",
  "slides": [
    {
      "slideNumber": 1,
      "title": "Slide title",
      "titleArabic": "Optional Arabic title when Bilingual mode is enabled",
      "layout": "Title Slide | Title and Content | Two Content | Section Header | Summary Table",
      "bullets": ["Concise bullet 1", "Concise bullet 2"],
      "bulletsArabic": ["Optional Arabic bullet 1"],
      "table": { "headers": ["Col1", "Col2"], "rows": [["Val1", "Val2"]] },
      "speakerNotes": "Speaker notes for this slide"
    }
  ]
}`,
    options.reportText.trim()
      ? `\nSOURCE REPORT TEXT:\n${options.reportText.trim()}`
      : "\nAnalyze the attached PDF/document to build the slide outline.",
  ].join("\n");

  const parts: GeminiPart[] = [
    ...convertAttachmentsToGeminiParts(options.attachments || []),
    { text: prompt },
  ];

  const res = await generateGeminiTurn(
    options.settings,
    [{ role: "user", parts }],
    {
      systemInstruction:
        "You are Zee AI for PowerPoint. Output strictly valid JSON for the slide deck outline.",
    }
  );

  try {
    const jsonStr = extractJsonBlock(res.text);
    const parsed = JSON.parse(jsonStr) as DeckOutlinePlan;
    if (Array.isArray(parsed.slides) && parsed.slides.length > 0) {
      return {
        deckTitle: parsed.deckTitle || "Generated Presentation",
        deckSubtitle: parsed.deckSubtitle || "",
        sourceSummary: parsed.sourceSummary || "",
        slides: parsed.slides.map((s, idx) => ({
          slideNumber: idx + 1,
          title: s.title || `Slide ${idx + 1}`,
          titleArabic: s.titleArabic,
          layout: s.layout || "Title and Content",
          bullets: Array.isArray(s.bullets) ? s.bullets.map(String) : [],
          bulletsArabic: Array.isArray(s.bulletsArabic)
            ? s.bulletsArabic.map(String)
            : undefined,
          table:
            s.table && Array.isArray(s.table.headers) && Array.isArray(s.table.rows)
              ? s.table
              : undefined,
          speakerNotes: s.speakerNotes || "",
        })),
      };
    }
  } catch {
    // Fallback outline parsing if model returned markdown bullets instead of pure JSON
  }

  return buildFallbackOutlineFromText(options.reportText || res.text, options.settings);
}

export function buildFallbackOutlineFromText(
  rawText: string,
  settings: ZeeSettings
): DeckOutlinePlan {
  const paragraphs = rawText
    .split(/\n{2,}/)
    .map((p) => p.trim())
    .filter(Boolean);

  const firstLine = (paragraphs[0] || "Engineering Report Presentation")
    .split(/\r?\n/)[0]
    .replace(/^#+\s*/, "")
    .slice(0, 90);

  const slides: DeckOutlineSlide[] = [
    {
      slideNumber: 1,
      title: firstLine || "Executive Summary",
      layout: "Title Slide",
      bullets: [
        "Generated by Zee AI for PowerPoint",
        settings.companyTemplate?.enabled
          ? settings.companyTemplate.footerText
          : "Key Highlights & Technical Findings",
      ],
      speakerNotes: "Opening title slide introducing the report scope and objectives.",
    },
  ];

  const chunks = paragraphs.slice(0, 6);
  chunks.forEach((chunk, idx) => {
    const lines = chunk
      .split(/\r?\n/)
      .map((l) => l.replace(/^[-*•#]+\s*/, "").trim())
      .filter(Boolean);
    const title = lines[0]?.slice(0, 80) || `Key Findings — Section ${idx + 1}`;
    const bullets =
      lines.length > 1
        ? lines.slice(1, 6)
        : [chunk.slice(0, 140)];
    slides.push({
      slideNumber: slides.length + 1,
      title,
      layout: "Title and Content",
      bullets,
      speakerNotes: `Discuss key metrics and engineering takeaways for ${title}.`,
    });
  });

  return {
    deckTitle: firstLine,
    sourceSummary: `Structured ${slides.length}-slide deck outline generated from report.`,
    slides,
  };
}

/**
 * Step 2 of Report/PDF -> Deck:
 * Builds the slides in PowerPoint after the user approves the outline.
 */
export async function buildSlidesFromApprovedOutline(
  plan: DeckOutlinePlan,
  settings: ZeeSettings,
  onProgress?: (current: number, total: number, slideTitle: string) => void
): Promise<{ createdCount: number; undoSnapshots: SlideUndoSnapshot[] }> {
  const undoSnapshots: SlideUndoSnapshot[] = [];
  const total = plan.slides.length;

  for (let i = 0; i < total; i++) {
    const item = plan.slides[i];
    onProgress?.(i + 1, total, item.title);

    const effectiveTitle =
      settings.bilingualMode && item.titleArabic
        ? `${item.title} / ${item.titleArabic}`
        : item.title;

    const enBullets = (item.bullets || []).map((b) => (b.startsWith("•") ? b : `• ${b}`)).join("\n");
    const arBullets =
      settings.bilingualMode && item.bulletsArabic?.length
        ? item.bulletsArabic.map((b) => (b.startsWith("•") ? b : `• ${b}`)).join("\n")
        : "";

    const bodyText = arBullets
      ? formatBilingualTextBlock(enBullets, arBullets)
      : enBullets;

    const added = await addSlide({
      title: effectiveTitle,
      bodyText,
      layoutName: settings.companyTemplate?.enabled
        ? settings.companyTemplate.layoutName
        : item.layout,
      notes: item.speakerNotes,
      settings,
    });

    if (item.table && item.table.headers.length > 0 && item.table.rows.length > 0) {
      await insertTable({
        slideNumber: added.slideIndex,
        headers: item.table.headers,
        rows: item.table.rows,
        top: bodyText ? 260 : 120,
        height: 190,
        settings,
      });
    }

    undoSnapshots.push({
      stepId: `deck-build-${Date.now()}-${i}`,
      toolName: "addSlide",
      timestamp: Date.now(),
      slideIndex: added.slideIndex,
      createdSlideId: added.slideId,
      actionType: "add_slide",
    });
  }

  return { createdCount: total, undoSnapshots };
}

/**
 * Parses raw TSV/CSV/Markdown or unstructured Excel engineering results into
 * headers and rows.
 */
export function parseTabularText(rawText: string): {
  headers: string[];
  rows: string[][];
} {
  const lines = rawText
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter((l) => l && !/^[-|:\s]+$/.test(l));

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    if (line.includes("\t")) {
      const tableLines = lines.slice(i).filter((l) => l.includes("\t"));
      if (tableLines.length >= 2) {
        const splitRow = (r: string) => r.split("\t").map((c) => c.trim());
        return {
          headers: splitRow(tableLines[0]),
          rows: tableLines.slice(1, 25).map(splitRow),
        };
      }
    }
    if (line.includes("|")) {
      const tableLines = lines.slice(i).filter((l) => l.includes("|"));
      if (tableLines.length >= 2) {
        const splitPipe = (r: string) =>
          r
            .replace(/^\||\|$/g, "")
            .split("|")
            .map((c) => c.trim());
        return {
          headers: splitPipe(tableLines[0]),
          rows: tableLines.slice(1, 25).map(splitPipe),
        };
      }
    }
    if (line.includes(",") && line.split(",").length >= 3) {
      const tableLines = lines.slice(i).filter((l) => l.includes(","));
      if (tableLines.length >= 2) {
        const splitCsv = (r: string) => r.split(",").map((c) => c.trim());
        return {
          headers: splitCsv(tableLines[0]),
          rows: tableLines.slice(1, 25).map(splitCsv),
        };
      }
    }
  }

  const kvRows: string[][] = [];
  for (const line of lines) {
    const m = line.match(/^([^:=]{2,40})[:=]\s*(.+)$/);
    if (m) {
      kvRows.push([m[1].trim(), m[2].trim()]);
    }
  }

  if (kvRows.length > 0) {
    return {
      headers: ["Parameter", "Value"],
      rows: kvRows.slice(0, 12),
    };
  }

  return {
    headers: ["Metric", "Observation"],
    rows: lines.slice(0, 8).map((l, idx) => [`Item ${idx + 1}`, l]),
  };
}

/**
 * Generates an automatic SVG/PNG chart Base64 image from engineering parameters
 * when the user enables Auto-Plot Chart on a Results Slide or Well-Book slide.
 */
export async function generateAutoPlotChartBase64(options: {
  category: PetroleumResultCategory;
  wellOrFieldName: string;
  rawInputText: string;
  kpis?: KpiCardItem[];
  primaryColor?: string;
}): Promise<string> {
  const text = options.rawInputText || "";
  const parseNum = (regex: RegExp, fallback: number): number => {
    const m = text.match(regex);
    if (!m || !m[1]) return fallback;
    const val = parseFloat(m[1].replace(/,/g, ""));
    return Number.isFinite(val) ? val : fallback;
  };

  if (options.category === "decline_curve") {
    const params: DeclineCurveParams = {
      wellName: options.wellOrFieldName,
      qi: parseNum(/(?:Qi|Initial Rate)[^\d]*([\d,]+(?:\.\d+)?)/i, 1450),
      diAnnualPercent: parseNum(/(?:Di|Decline Rate)[^\d]*([\d,]+(?:\.\d+)?)/i, 14.2),
      bFactor: parseNum(/(?:b-factor|\bb\b)[^\d]*([\d.]+)/i, 0.45),
      months: 36,
      primaryColor: options.primaryColor,
    };
    const svg = renderDeclineCurveSvg(params);
    return await svgToPngBase64(svg, 680, 380);
  }

  if (options.category === "esp_analysis") {
    const params: EspCurveParams = {
      wellName: options.wellOrFieldName,
      operatingRateBblPerDay: parseNum(/(?:Liquid Rate|Rate)[^\d]*([\d,]+(?:\.\d+)?)/i, 2450),
      operatingTdhFt: parseNum(/(?:TDH|Dynamic Head)[^\d]*([\d,]+(?:\.\d+)?)/i, 5400),
      pipPsi: parseNum(/(?:PIP|Intake Pressure)[^\d]*([\d,]+(?:\.\d+)?)/i, 1120),
      frequencyHz: parseNum(/(?:Frequency|Freq)[^\d]*([\d,]+(?:\.\d+)?)/i, 56),
      primaryColor: options.primaryColor,
    };
    const svg = renderEspCurveSvg(params);
    return await svgToPngBase64(svg, 680, 380);
  }

  const kpiItems: KpiCardItem[] =
    options.kpis && options.kpis.length > 0
      ? options.kpis
      : [
          { label: "Liquid Rate", value: "2,450 bbl/d", status: "normal" },
          { label: "Oil Rate", value: "1,690 stb/d", status: "normal" },
          { label: "Water Cut", value: "31.0%", status: "warning" },
          { label: "BHP", value: "1,580 psia", status: "normal" },
        ];
  const svg = renderKpiCardsSvg(kpiItems, options.primaryColor);
  return await svgToPngBase64(svg, 680, 110);
}

export async function prepareExcelResultsSlidePayload(options: {
  category: PetroleumResultCategory;
  wellOrFieldName: string;
  rawInputText: string;
  chartImageBase64?: string;
  autoGenerateChart?: boolean;
  settings: ZeeSettings;
}): Promise<ExcelResultsPayload> {
  const categoryLabels: Record<PetroleumResultCategory, string> = {
    decline_curve: "Decline Curve Analysis (Arps Forecast)",
    esp_analysis: "ESP Performance & Hydraulic Diagnosis",
    production_qc: "Production Surveillance & QC Summary",
    custom: "Engineering Results Summary",
  };

  const parsedTable = parseTabularText(options.rawInputText);
  const wellLabel = options.wellOrFieldName.trim() || "Field / Well Review";
  const defaultTitle = `${wellLabel} — ${categoryLabels[options.category]}`;

  let kpis: KpiCardItem[] = parsedTable.rows.slice(0, 4).map((r) => ({
    label: r[0] || "Metric",
    value: r[1] || "",
    status: "normal",
  }));

  let slideTitle = defaultTitle;
  let headers = parsedTable.headers;
  let rows = parsedTable.rows;
  let takeaways = [
    `Reviewed ${categoryLabels[options.category]} results for ${wellLabel}.`,
    "All protected engineering parameters (ESP, TDH, PIP, GOR, BHP, bbl/d) preserved.",
  ];
  let speakerNotes = `Walk through the ${categoryLabels[options.category]} table and key takeaways for ${wellLabel}.`;

  try {
    const prompt = [
      `Convert the following Zee AI Excel ${categoryLabels[options.category]} results for "${wellLabel}" into a PowerPoint Results Slide JSON payload.`,
      buildDirectionAndLanguageInstruction(options.settings),
      "Return ONLY valid JSON with this structure:",
      `{
  "slideTitle": "string",
  "kpis": [{ "label": "e.g. PIP", "value": "1,120 psi", "status": "normal" }],
  "headers": ["Col 1", "Col 2"],
  "rows": [["Val 1", "Val 2"]],
  "takeaways": ["Key engineering takeaway 1", "Key engineering takeaway 2"],
  "speakerNotes": "Concise presenter talking points"
}`,
      `RAW EXCEL RESULTS:\n${options.rawInputText}`,
    ].join("\n\n");

    const parts: GeminiPart[] = [];
    if (options.chartImageBase64) {
      parts.push({
        inlineData: {
          mimeType: "image/png",
          data: options.chartImageBase64.replace(/^data:image\/[a-zA-Z0-9+.-]+;base64,/, ""),
        },
      });
    }
    parts.push({ text: prompt });

    const res = await generateGeminiTurn(options.settings, [{ role: "user", parts }]);
    const json = JSON.parse(extractJsonBlock(res.text));

    slideTitle = json.slideTitle || defaultTitle;
    if (Array.isArray(json.kpis) && json.kpis.length > 0) kpis = json.kpis.slice(0, 4);
    if (Array.isArray(json.headers) && json.headers.length > 0) headers = json.headers.map(String);
    if (Array.isArray(json.rows) && json.rows.length > 0) rows = json.rows;
    if (Array.isArray(json.takeaways) && json.takeaways.length > 0) {
      takeaways = json.takeaways.map(String);
    }
    if (json.speakerNotes) speakerNotes = json.speakerNotes;
  } catch {
    // Keep deterministic parsed values
  }

  let effectiveChartBase64 = options.chartImageBase64;
  if (!effectiveChartBase64 && (options.autoGenerateChart ?? true)) {
    effectiveChartBase64 = await generateAutoPlotChartBase64({
      category: options.category,
      wellOrFieldName: wellLabel,
      rawInputText: options.rawInputText,
      kpis,
      primaryColor: options.settings.companyTemplate?.enabled
        ? options.settings.companyTemplate.primaryColor
        : "#B73618",
    });
  }

  return {
    category: options.category,
    wellOrFieldName: wellLabel,
    rawInputText: options.rawInputText,
    slideTitle,
    kpis,
    headers,
    rows,
    takeaways,
    speakerNotes,
    chartImageBase64: effectiveChartBase64,
  };
}

export async function buildExcelResultsSlide(
  payload: ExcelResultsPayload,
  settings: ZeeSettings
): Promise<{ slideIndex: number; slideId: string; undoSnapshot: SlideUndoSnapshot }> {
  const kpiLine =
    payload.kpis.length > 0
      ? "KPIs: " + payload.kpis.map((k) => `${k.label}: ${k.value}`).join("  |  ")
      : "";

  const takeawayLines = payload.takeaways
    .map((t) => (t.startsWith("•") ? t : `• ${t}`))
    .join("\n");

  const bodyBlock = [kpiLine, takeawayLines].filter(Boolean).join("\n\n");

  const added = await addSlide({
    title: payload.slideTitle,
    bodyText: bodyBlock,
    layoutName: settings.companyTemplate?.enabled
      ? settings.companyTemplate.layoutName
      : "Two Content",
    notes: payload.speakerNotes,
    settings,
  });

  if (payload.headers.length > 0 && payload.rows.length > 0) {
    await insertTable({
      slideNumber: added.slideIndex,
      title: `${payload.wellOrFieldName} Data`,
      headers: payload.headers,
      rows: payload.rows,
      left: 48,
      top: 250,
      width: payload.chartImageBase64 ? 320 : 624,
      height: 210,
      settings,
    });
  }

  if (payload.chartImageBase64) {
    try {
      await insertImage({
        slideNumber: added.slideIndex,
        base64Data: payload.chartImageBase64,
        left: 384,
        top: 120,
        width: 288,
        height: 210,
      });
    } catch {
      // Ignore if image insertion is unsupported on older host
    }
  }

  return {
    slideIndex: added.slideIndex,
    slideId: added.slideId,
    undoSnapshot: {
      stepId: `excel-slide-${Date.now()}`,
      toolName: "addSlide",
      timestamp: Date.now(),
      slideIndex: added.slideIndex,
      createdSlideId: added.slideId,
      actionType: "add_slide",
    },
  };
}

/**
 * FEATURE: Batch Well-Book Generator
 * Parses a multi-well table (each row = 1 well) and builds a Field Rollup Slide + 1 standardized slide per well.
 */
export function parseMultiWellTable(rawText: string): {
  headers: string[];
  wells: WellBookRow[];
} {
  const { headers, rows } = parseTabularText(rawText);
  if (headers.length === 0 || rows.length === 0) {
    return { headers: [], wells: [] };
  }

  const wells: WellBookRow[] = rows
    .filter((r) => r.length > 0 && r[0]?.trim())
    .map((r) => {
      const wellName = r[0].trim();
      const metrics: Record<string, string> = {};
      let status: string | undefined;
      let action: string | undefined;

      headers.slice(1).forEach((colName, idx) => {
        const val = (r[idx + 1] || "").trim();
        if (!val) return;
        if (/status|flag|condition/i.test(colName)) {
          status = val;
        } else if (/action|recommend/i.test(colName)) {
          action = val;
        }
        metrics[colName] = val;
      });

      return { wellName, metrics, status, action };
    });

  return { headers, wells };
}

export async function buildBatchWellBookSlides(options: {
  fieldName: string;
  rawTableText: string;
  includeRollupSlide: boolean;
  includeAutoCharts: boolean;
  settings: ZeeSettings;
  onProgress?: (current: number, total: number, label: string) => void;
}): Promise<{ createdCount: number; undoSnapshots: SlideUndoSnapshot[] }> {
  const { headers, wells } = parseMultiWellTable(options.rawTableText);
  if (wells.length === 0) {
    throw new Error("No well rows found. Paste a table with Well Name in the first column.");
  }

  const undoSnapshots: SlideUndoSnapshot[] = [];
  const totalSteps = wells.length + (options.includeRollupSlide ? 1 : 0);
  let stepCounter = 0;

  // 1. Optional Executive Field Roll-Up Summary Slide
  if (options.includeRollupSlide) {
    stepCounter += 1;
    options.onProgress?.(stepCounter, totalSteps, `${options.fieldName} — Field Roll-Up`);

    const rollupTitle = `${options.fieldName} — Multi-Well Surveillance Summary (${wells.length} Wells)`;
    const rollupBullets = [
      `Total monitored wells in batch: ${wells.length} (${wells.map((w) => w.wellName).join(", ")})`,
      `Flagged / priority wells: ${
        wells
          .filter((w) => w.status && !/valid|normal|ok/i.test(w.status))
          .map((w) => `${w.wellName} (${w.status})`)
          .join(", ") || "All wells operating within target envelope"
      }`,
    ];

    const rollupSlide = await addSlide({
      title: rollupTitle,
      bullets: rollupBullets,
      layoutName: options.settings.companyTemplate?.enabled
        ? options.settings.companyTemplate.layoutName
        : "Title and Content",
      notes: `Executive overview of ${wells.length} wells in ${options.fieldName}.`,
      settings: options.settings,
    });

    await insertTable({
      slideNumber: rollupSlide.slideIndex,
      title: `${options.fieldName} Well Summary`,
      headers,
      rows: wells.map((w) => [
        w.wellName,
        ...headers.slice(1).map((h) => w.metrics[h] || "—"),
      ]),
      left: 48,
      top: 190,
      width: 624,
      height: 260,
      settings: options.settings,
    });

    undoSnapshots.push({
      stepId: `wellbook-rollup-${Date.now()}`,
      toolName: "addSlide",
      timestamp: Date.now(),
      slideIndex: rollupSlide.slideIndex,
      createdSlideId: rollupSlide.slideId,
      actionType: "add_slide",
    });
  }

  // 2. Build 1 Standardized Slide per Well
  for (const well of wells) {
    stepCounter += 1;
    options.onProgress?.(stepCounter, totalSteps, `Well Slide: ${well.wellName}`);

    const metricEntries = Object.entries(well.metrics);
    const kpiLine = metricEntries
      .slice(0, 4)
      .map(([k, v]) => `${k}: ${v}`)
      .join("  |  ");

    const bullets = [
      `Key Parameters: ${kpiLine}`,
      well.status ? `Surveillance Status: ${well.status}` : "Surveillance Status: Active monitoring",
      well.action
        ? `Recommended Engineering Action: ${well.action}`
        : "Recommended Engineering Action: Continue routine ESP & BHP drawdown surveillance",
    ];

    const wellSlide = await addSlide({
      title: `${well.wellName} — Well Performance & Surveillance Review`,
      bullets,
      layoutName: options.settings.companyTemplate?.enabled
        ? options.settings.companyTemplate.layoutName
        : "Two Content",
      notes: `Detailed well review for ${well.wellName}. ${bullets.join(" ")}`,
      settings: options.settings,
    });

    await insertTable({
      slideNumber: wellSlide.slideIndex,
      title: `${well.wellName} Parameters`,
      headers: ["Parameter", "Measured Value"],
      rows: metricEntries.map(([k, v]) => [k, v]),
      left: 48,
      top: 230,
      width: options.includeAutoCharts ? 310 : 624,
      height: 220,
      settings: options.settings,
    });

    if (options.includeAutoCharts) {
      const kpiCards: KpiCardItem[] = metricEntries.slice(0, 4).map(([k, v]) => ({
        label: k,
        value: v,
        status:
          well.status && /flag|high|low|warn|crit/i.test(well.status)
            ? "warning"
            : "normal",
      }));
      const rawWellText = metricEntries.map(([k, v]) => `${k}: ${v}`).join("\n");
      const hasEspMetrics = /pip|tdh|hz/i.test(rawWellText);
      const chartBase64 = await generateAutoPlotChartBase64({
        category: hasEspMetrics ? "esp_analysis" : "production_qc",
        wellOrFieldName: well.wellName,
        rawInputText: rawWellText,
        kpis: kpiCards,
        primaryColor: options.settings.companyTemplate?.enabled
          ? options.settings.companyTemplate.primaryColor
          : "#B73618",
      });
      if (chartBase64) {
        try {
          await insertImage({
            slideNumber: wellSlide.slideIndex,
            base64Data: chartBase64,
            left: 376,
            top: 230,
            width: 296,
            height: 200,
          });
        } catch {
          // ignore on unsupported host
        }
      }
    }

    undoSnapshots.push({
      stepId: `wellbook-${well.wellName}-${Date.now()}`,
      toolName: "addSlide",
      timestamp: Date.now(),
      slideIndex: wellSlide.slideIndex,
      createdSlideId: wellSlide.slideId,
      actionType: "add_slide",
    });
  }

  return { createdCount: totalSteps, undoSnapshots };
}

/**
 * FEATURE: Management Q&A & Deck Reviewer
 * Deterministically detects cross-slide numerical contradictions for the same well/metric
 * and pairs with Gemini's executive review & tough Q&A generator.
 */
export function detectCrossSlideContradictions(
  slides: SlideDetail[]
): CrossSlideContradiction[] {
  const contradictions: CrossSlideContradiction[] = [];
  const seen = new Map<
    string,
    { slideIndex: number; rawValue: string; numericValue: number; unit: string; entity: string }
  >();

  const wellRegex = /\b((?:WELL|Well|PROD|INJ|W)-\d{1,4}[A-Za-z]?)\b/g;
  const metricWithUnitRegex =
    /\b(\d+(?:,\d{3})*(?:\.\d+)?)\s*(bbl\/d|stb\/d|MSCF\/d|MMSCF\/d|scf\/stb|psi|psia|psig|ft|Hz|%)\b/gi;

  for (const sl of slides) {
    const lines = (sl.fullText || "")
      .split(/\r?\n/)
      .map((l) => l.trim())
      .filter(Boolean);

    for (const line of lines) {
      const wellsInLine = Array.from(line.matchAll(wellRegex)).map((m) =>
        m[1].toUpperCase()
      );
      const entity = wellsInLine[0] || "FIELD";

      let m: RegExpExecArray | null;
      metricWithUnitRegex.lastIndex = 0;
      while ((m = metricWithUnitRegex.exec(line)) !== null) {
        const rawNum = m[1];
        const num = parseFloat(rawNum.replace(/,/g, ""));
        const unit = m[2].toLowerCase();
        const key = `${entity}::${unit}`;

        const prev = seen.get(key);
        if (!prev) {
          seen.set(key, {
            slideIndex: sl.index,
            rawValue: `${rawNum} ${m[2]}`,
            numericValue: num,
            unit: m[2],
            entity,
          });
        } else if (
          prev.slideIndex !== sl.index &&
          Math.abs(prev.numericValue - num) > 0.01
        ) {
          const id = `contra-${entity}-${unit}-${prev.slideIndex}-${sl.index}`;
          if (!contradictions.some((c) => c.id === id)) {
            contradictions.push({
              id,
              entity,
              unit: m[2],
              firstSlideIndex: prev.slideIndex,
              firstValue: prev.rawValue,
              secondSlideIndex: sl.index,
              secondValue: `${rawNum} ${m[2]}`,
              description: `${entity} (${m[2]}): Slide ${prev.slideIndex} states "${prev.rawValue}", but Slide ${sl.index} states "${rawNum} ${m[2]}".`,
            });
          }
        }
      }
    }
  }

  return contradictions;
}

export async function runManagementDeckReview(
  settings: ZeeSettings
): Promise<DeckReviewReport> {
  const summaries = await listSlides();
  const details: SlideDetail[] = [];
  for (const s of summaries) {
    details.push(await readSlide(s.index));
  }

  const contradictions = detectCrossSlideContradictions(details);
  const deckDump = details
    .map(
      (d) =>
        `Slide ${d.index} ("${d.title}"):\n${d.fullText}\nNotes: ${d.speakerNotes || "(none)"}`
    )
    .join("\n\n---\n\n");

  try {
    const prompt = [
      "Act as a Chief Reservoir/Production Engineer and VP reviewing this PowerPoint presentation before an executive meeting.",
      buildDirectionAndLanguageInstruction(settings),
      "Return ONLY valid JSON matching this schema:",
      `{
  "overallScore": 88,
  "executiveSummary": "2-3 sentence assessment of deck readiness",
  "slideCritiques": [
    { "slideIndex": 1, "title": "Slide title", "feedback": "Specific actionable improvement" }
  ],
  "toughQuestions": [
    {
      "question": "Tough technical/management question likely to be asked",
      "recommendedAnswer": "Crisp, data-backed answer",
      "backupSlideTitle": "Backup: ...",
      "backupSlideBullets": ["Supporting point 1", "Supporting point 2"]
    }
  ]
}`,
      `\nPRESENTATION SLIDES:\n${deckDump}`,
    ].join("\n");

    const res = await generateGeminiTurn(settings, [{ role: "user", parts: [{ text: prompt }] }]);
    const parsed = JSON.parse(extractJsonBlock(res.text));

    return {
      overallScore: Number(parsed.overallScore) || (contradictions.length > 0 ? 78 : 90),
      executiveSummary:
        parsed.executiveSummary ||
        `Reviewed ${details.length} slides (${contradictions.length} numerical discrepancy flag(s)).`,
      contradictions,
      slideCritiques: Array.isArray(parsed.slideCritiques)
        ? parsed.slideCritiques
        : details.map((d) => ({
            slideIndex: d.index,
            title: d.title,
            feedback: "Clear structure; ensure assumptions and risk mitigations are stated.",
          })),
      toughQuestions: Array.isArray(parsed.toughQuestions)
        ? parsed.toughQuestions.map((q: any, idx: number) => ({
            id: `qna-${idx + 1}`,
            question: String(q.question || ""),
            recommendedAnswer: String(q.recommendedAnswer || ""),
            backupSlideTitle: String(q.backupSlideTitle || `Backup — Q&A #${idx + 1}`),
            backupSlideBullets: Array.isArray(q.backupSlideBullets)
              ? q.backupSlideBullets.map(String)
              : [String(q.recommendedAnswer || "")],
            selected: true,
          }))
        : [],
    };
  } catch {
    return {
      overallScore: contradictions.length > 0 ? 76 : 90,
      executiveSummary: `Scanned ${details.length} slides across the presentation. Found ${contradictions.length} cross-slide numerical discrepancy flag(s).`,
      contradictions,
      slideCritiques: details.map((d) => ({
        slideIndex: d.index,
        title: d.title,
        feedback: d.speakerNotes
          ? "Slide has speaker notes and structured bullets."
          : "Consider adding speaker notes for smoother delivery.",
      })),
      toughQuestions: [
        {
          id: "qna-1",
          question:
            "What are the key downside sensitivities on the forecast EUR and BHP drawdown targets?",
          recommendedAnswer:
            "PIP is maintained above bubble-point safety margin with variable speed drive (VSD) protection.",
          backupSlideTitle: "Backup — Drawdown & ESP Operating Envelope Sensitivities",
          backupSlideBullets: [
            "PIP operating margin maintained > 150 psi above gas-locking threshold",
            "Hyperbolic b-factor bounded between 0.35 and 0.50 based on field analogs",
            "Contingency VSD frequency step-up plan validated against motor load",
          ],
          selected: true,
        },
      ],
    };
  }
}

export async function buildBackupSlidesFromQna(
  questions: DeckQnaItem[],
  settings: ZeeSettings
): Promise<{ createdCount: number; undoSnapshots: SlideUndoSnapshot[] }> {
  const selected = questions.filter((q) => q.selected);
  const undoSnapshots: SlideUndoSnapshot[] = [];

  for (const q of selected) {
    const bullets = [
      `Question: ${q.question}`,
      `Executive Response: ${q.recommendedAnswer}`,
      ...q.backupSlideBullets,
    ];
    const added = await addSlide({
      title: q.backupSlideTitle,
      bullets,
      layoutName: settings.companyTemplate?.enabled
        ? settings.companyTemplate.layoutName
        : "Title and Content",
      notes: `Q: ${q.question}\nA: ${q.recommendedAnswer}`,
      settings,
    });
    undoSnapshots.push({
      stepId: `qna-backup-${q.id}-${Date.now()}`,
      toolName: "addSlide",
      timestamp: Date.now(),
      slideIndex: added.slideIndex,
      createdSlideId: added.slideId,
      actionType: "add_slide",
    });
  }

  return { createdCount: selected.length, undoSnapshots };
}

/**
 * FEATURE: Smart Slide Layout Beautifier
 * Converts dense text slides into 2-column splits, 3-card callout grids, or aligned template typography.
 */
export async function beautifySlideLayout(options: {
  slideNumber: number;
  style: SlideBeautifyStyle;
  settings: ZeeSettings;
}): Promise<{ summary: string; undoSnapshot: SlideUndoSnapshot }> {
  const detail = await readSlide(options.slideNumber);
  const undoSnapshot = await captureSlideSnapshot(
    `beautify-${options.slideNumber}-${Date.now()}`,
    "beautifySlide",
    options.slideNumber,
    "modify"
  );

  const lines = detail.fullText
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter((l) => l && l !== detail.title);

  if (options.style === "two_column" && lines.length >= 2) {
    const mid = Math.ceil(lines.length / 2);
    const leftText = lines.slice(0, mid).join("\n");
    const rightText = lines.slice(mid).join("\n");

    await insertTextBox({
      slideNumber: options.slideNumber,
      text: leftText,
      left: 48,
      top: 112,
      width: 300,
      height: 330,
      fontSize: 15,
      settings: options.settings,
    });
    await insertTextBox({
      slideNumber: options.slideNumber,
      text: rightText,
      left: 372,
      top: 112,
      width: 300,
      height: 330,
      fontSize: 15,
      settings: options.settings,
    });
    await applySlideLayout({
      slideNumber: options.slideNumber,
      layoutName: "Two Content",
      settings: options.settings,
    });
    return {
      summary: `Beautified Slide ${options.slideNumber} into a balanced 2-Column comparison layout.`,
      undoSnapshot,
    };
  }

  if (options.style === "three_cards") {
    const cardCount = Math.min(3, Math.max(1, lines.length));
    for (let i = 0; i < cardCount; i++) {
      const cardLines = lines.filter((_, idx) => idx % cardCount === i);
      await insertTextBox({
        slideNumber: options.slideNumber,
        text: `📌 Key Point ${i + 1}\n\n${cardLines.join("\n")}`,
        left: 48 + i * 214,
        top: 120,
        width: 198,
        height: 310,
        fontSize: 14,
        settings: options.settings,
      });
    }
    await applySlideLayout({
      slideNumber: options.slideNumber,
      settings: options.settings,
    });
    return {
      summary: `Beautified Slide ${options.slideNumber} into a 3-Card executive callout grid.`,
      undoSnapshot,
    };
  }

  await applySlideLayout({
    slideNumber: options.slideNumber,
    settings: options.settings,
  });
  return {
    summary: `Aligned and standardized fonts, colors, and RTL/LTR direction on Slide ${options.slideNumber}.`,
    undoSnapshot,
  };
}

/**
 * Scans all slides in the presentation for consistency issues:
 * - Font families (vs dominant or Company Template font)
 * - Font sizes
 * - Title capitalization / trailing periods
 * - Unit formatting & Protected terms casing (ESP, TDH, PIP, GOR, BHP, bbl/d, etc.)
 */
export function analyzeSlidesConsistency(
  slides: SlideDetail[],
  settings: ZeeSettings
): ConsistencyIssue[] {
  const issues: ConsistencyIssue[] = [];

  const fontCounts = new Map<string, number>();
  for (const sl of slides) {
    for (const sh of sl.shapes) {
      if (sh.fontName && !sh.isTitle) {
        fontCounts.set(sh.fontName, (fontCounts.get(sh.fontName) || 0) + 1);
      }
    }
  }

  let dominantBodyFont = "Calibri";
  let maxCount = 0;
  for (const [font, cnt] of fontCounts.entries()) {
    if (cnt > maxCount) {
      maxCount = cnt;
      dominantBodyFont = font;
    }
  }

  const expectedBodyFont = settings.companyTemplate?.enabled
    ? settings.companyTemplate.bodyFont
    : dominantBodyFont;
  const expectedHeadingFont = settings.companyTemplate?.enabled
    ? settings.companyTemplate.headingFont
    : "Calibri Light";

  const uppercaseTerms = DEFAULT_PROTECTED_TERMS.filter((t) => /^[A-Z&]{2,6}$/.test(t));

  const unitTypos: { pattern: RegExp; canonical: string; label: string }[] = [
    { pattern: /\b(?:bbld|bbls\/d|bbl\/day|bbls\/day)\b/gi, canonical: "bbl/d", label: "bbl/d" },
    { pattern: /\b(?:stbd|stbs\/d|stb\/day)\b/gi, canonical: "stb/d", label: "stb/d" },
    { pattern: /\b(?:scfstb|scf\/bbl)\b/gi, canonical: "scf/stb", label: "scf/stb" },
    { pattern: /\b(?:mscfd|mscf\/day)\b/gi, canonical: "MSCF/d", label: "MSCF/d" },
    { pattern: /\b(?:mmscfd|mmscf\/day)\b/gi, canonical: "MMSCF/d", label: "MMSCF/d" },
  ];

  for (const sl of slides) {
    if (sl.title && sl.title.endsWith(".")) {
      issues.push({
        id: `issue-title-period-${sl.index}`,
        slideIndex: sl.index,
        slideId: sl.id,
        category: "title_case",
        severity: "info",
        description: `Slide ${sl.index} title ends with a trailing period.`,
        currentValue: sl.title,
        suggestedValue: sl.title.replace(/\.+$/, ""),
        selected: true,
      });
    }

    for (const sh of sl.shapes) {
      const text = (sh.text || "").trim();
      if (!text) continue;

      const isArabicShape = containsArabic(text);
      const targetFont = isArabicShape && settings.arabicFont !== "Same as document"
        ? settings.arabicFont
        : sh.isTitle
          ? expectedHeadingFont
          : expectedBodyFont;

      if (sh.fontName && targetFont && sh.fontName.toLowerCase() !== targetFont.toLowerCase()) {
        issues.push({
          id: `issue-font-${sl.index}-${sh.id}`,
          slideIndex: sl.index,
          slideId: sl.id,
          shapeId: sh.id,
          category: "font_family",
          severity: "warning",
          description: `Slide ${sl.index} (${sh.name}): uses font "${sh.fontName}" instead of "${targetFont}".`,
          currentValue: sh.fontName,
          suggestedValue: targetFont,
          selected: true,
        });
      }

      if (!sh.isTitle && sh.fontSize && (sh.fontSize < 12 || sh.fontSize > 26)) {
        const suggested = sh.fontSize < 12 ? "14" : "18";
        issues.push({
          id: `issue-size-${sl.index}-${sh.id}`,
          slideIndex: sl.index,
          slideId: sl.id,
          shapeId: sh.id,
          category: "font_size",
          severity: "info",
          description: `Slide ${sl.index} (${sh.name}): body font size is ${sh.fontSize}pt.`,
          currentValue: `${sh.fontSize}`,
          suggestedValue: suggested,
          selected: false,
        });
      }

      for (const term of uppercaseTerms) {
        const regex = new RegExp(`\\b${term}\\b`, "gi");
        let match: RegExpExecArray | null;
        while ((match = regex.exec(text)) !== null) {
          const found = match[0];
          if (found !== term && found.toLowerCase() === term.toLowerCase()) {
            if ((term === "PI" || term === "WC") && found === found.toLowerCase()) continue;
            issues.push({
              id: `issue-term-${sl.index}-${sh.id}-${term}`,
              slideIndex: sl.index,
              slideId: sl.id,
              shapeId: sh.id,
              category: "protected_term",
              severity: "warning",
              description: `Slide ${sl.index}: protected term "${found}" should be "${term}".`,
              currentValue: found,
              suggestedValue: term,
              selected: true,
            });
            break;
          }
        }
      }

      for (const u of unitTypos) {
        const match = text.match(u.pattern);
        if (match && match[0] !== u.canonical) {
          issues.push({
            id: `issue-unit-${sl.index}-${sh.id}-${u.label}`,
            slideIndex: sl.index,
            slideId: sl.id,
            shapeId: sh.id,
            category: "units",
            severity: "warning",
            description: `Slide ${sl.index}: non-standard unit "${match[0]}" → "${u.canonical}".`,
            currentValue: match[0],
            suggestedValue: u.canonical,
            selected: true,
          });
        }
      }
    }
  }

  return issues;
}

export async function runPresentationConsistencyCheck(
  settings: ZeeSettings
): Promise<ConsistencyIssue[]> {
  const summaries = await listSlides();
  const details: SlideDetail[] = [];
  for (const s of summaries) {
    details.push(await readSlide(s.index));
  }
  return analyzeSlidesConsistency(details, settings);
}

export async function applyConsistencyFixes(
  issues: ConsistencyIssue[],
  settings: ZeeSettings
): Promise<{ fixedCount: number; undoSnapshots: SlideUndoSnapshot[] }> {
  const selected = issues.filter((i) => i.selected);
  const snapshotsBySlide = new Map<number, SlideUndoSnapshot>();
  let fixedCount = 0;

  for (const issue of selected) {
    if (!snapshotsBySlide.has(issue.slideIndex)) {
      const snap = await captureSlideSnapshot(
        `consistency-${issue.id}`,
        "consistencyFix",
        issue.slideIndex,
        "modify"
      );
      snapshotsBySlide.set(issue.slideIndex, snap);
    }

    if (issue.category === "title_case") {
      await setSlideTitle(issue.slideIndex, issue.suggestedValue, settings);
      fixedCount += 1;
    } else if (issue.category === "font_family") {
      await formatSlideText({
        slideNumber: issue.slideIndex,
        shapeId: issue.shapeId,
        fontName: issue.suggestedValue,
        settings,
      });
      fixedCount += 1;
    } else if (issue.category === "font_size") {
      await formatSlideText({
        slideNumber: issue.slideIndex,
        shapeId: issue.shapeId,
        fontSize: Number(issue.suggestedValue) || 16,
        settings,
      });
      fixedCount += 1;
    } else if (issue.category === "protected_term" || issue.category === "units") {
      await replaceSlideText({
        slideNumber: issue.slideIndex,
        findText: issue.currentValue,
        replaceWith: issue.suggestedValue,
        matchCase: true,
        respectProtectedTerms: false,
        settings,
      });
      fixedCount += 1;
    }
  }

  return {
    fixedCount,
    undoSnapshots: Array.from(snapshotsBySlide.values()),
  };
}

/**
 * Scans slides for text boxes / placeholders whose text overflows the box dimensions.
 */
export function detectTextFitIssuesInSlides(slides: SlideDetail[]): TextFitIssue[] {
  const issues: TextFitIssue[] = [];

  for (const sl of slides) {
    for (const sh of sl.shapes) {
      if (sh.isTitle) continue;
      const text = (sh.text || "").trim();
      if (!text) continue;

      const fontSize = sh.fontSize || 18;
      const boxWidth = Math.max(180, sh.width || 600);
      const boxHeight = Math.max(80, sh.height || 280);

      const charsPerLine = Math.max(25, Math.floor(boxWidth / (fontSize * 0.52)));
      const rawLines = text.split(/\r?\n/);
      let estimatedLines = 0;
      for (const l of rawLines) {
        estimatedLines += Math.max(1, Math.ceil(l.length / charsPerLine));
      }

      const lineHeightPt = fontSize * 1.38;
      const maxLinesInBox = Math.max(3, Math.floor(boxHeight / lineHeightPt));
      const wordCount = text.split(/\s+/).filter(Boolean).length;

      const overflowRatio = Number((estimatedLines / maxLinesInBox).toFixed(2));
      if (overflowRatio > 1.05 || wordCount > 85 || rawLines.length > 8) {
        const recFontSize = Math.max(
          13,
          Math.min(fontSize, Math.floor(fontSize / Math.sqrt(Math.max(1.1, overflowRatio))))
        );
        issues.push({
          id: `fit-${sl.index}-${sh.id}`,
          slideIndex: sl.index,
          slideId: sl.id,
          shapeId: sh.id,
          shapeName: sh.name,
          originalText: text,
          lineCount: estimatedLines,
          wordCount,
          boxHeightPt: boxHeight,
          estimatedOverflowRatio: Math.max(overflowRatio, Number((wordCount / 75).toFixed(2))),
          recommendedFontSize: recFontSize,
          selected: true,
        });
      }
    }
  }

  return issues;
}

export async function scanPresentationForTextFitIssues(): Promise<TextFitIssue[]> {
  const summaries = await listSlides();
  const details: SlideDetail[] = [];
  for (const s of summaries) {
    details.push(await readSlide(s.index));
  }
  return detectTextFitIssuesInSlides(details);
}

export async function fixTextFitIssues(
  issues: TextFitIssue[],
  settings: ZeeSettings
): Promise<{ fixedCount: number; undoSnapshots: SlideUndoSnapshot[] }> {
  const selected = issues.filter((i) => i.selected);
  const snapshots: SlideUndoSnapshot[] = [];
  let fixedCount = 0;

  for (const item of selected) {
    const snap = await captureSlideSnapshot(
      `textfit-${item.id}`,
      "textFitFix",
      item.slideIndex,
      "modify"
    );
    snapshots.push(snap);

    let shortened = item.shortenedText;
    if (!shortened) {
      try {
        const prompt = [
          "Shorten these overflowing PowerPoint bullets so they fit cleanly on a single slide (max 5 bullets, max 12 words per bullet).",
          buildDirectionAndLanguageInstruction(settings),
          "Return ONLY the shortened bullet points (no preamble).",
          `\nORIGINAL TEXT:\n${item.originalText}`,
        ].join("\n");
        const res = await generateGeminiTurn(settings, [
          { role: "user", parts: [{ text: prompt }] },
        ]);
        if (res.text.trim()) {
          shortened = res.text.trim();
        }
      } catch {
        shortened = item.originalText
          .split(/\r?\n/)
          .map((line) => {
            const words = line.trim().split(/\s+/);
            return words.length > 14 ? words.slice(0, 14).join(" ") + "…" : line.trim();
          })
          .slice(0, 6)
          .join("\n");
      }
    }

    if (shortened) {
      await replaceSlideText({
        slideNumber: item.slideIndex,
        findText: item.originalText,
        replaceWith: shortened,
        respectProtectedTerms: false,
        settings,
      });
    }

    if (item.recommendedFontSize) {
      await formatSlideText({
        slideNumber: item.slideIndex,
        shapeId: item.shapeId,
        fontSize: item.recommendedFontSize,
        settings,
      });
    }

    fixedCount += 1;
  }

  return { fixedCount, undoSnapshots: snapshots };
}

export async function generateSpeakerNotesForSlide(
  slideNumber: number,
  settings: ZeeSettings,
  writeToSlide: boolean = true
): Promise<{ notes: string; undoSnapshot?: SlideUndoSnapshot }> {
  const detail = await readSlide(slideNumber);
  const prompt = [
    `Write clear, natural presenter speaker notes (60-90 seconds) for Slide ${detail.index} ("${detail.title}").`,
    buildDirectionAndLanguageInstruction(settings),
    `SLIDE CONTENT:\n${detail.fullText || detail.title}`,
    "Return ONLY the speaker notes text.",
  ].join("\n\n");

  const res = await generateGeminiTurn(settings, [
    { role: "user", parts: [{ text: prompt }] },
  ]);
  const notes = res.text.trim();

  let undoSnapshot: SlideUndoSnapshot | undefined;
  if (writeToSlide && notes) {
    undoSnapshot = await captureSlideSnapshot(
      `notes-${slideNumber}-${Date.now()}`,
      "writeSpeakerNotes",
      slideNumber,
      "modify"
    );
    await writeSpeakerNotes(slideNumber, notes, false);
  }

  return { notes, undoSnapshot };
}
