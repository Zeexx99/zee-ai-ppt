import { describe, expect, it } from "vitest";
import {
  computeArpsDeclineCurve,
  renderDeclineCurveSvg,
  renderEspCurveSvg,
  renderKpiCardsSvg,
  renderWellboreSchematicSvg,
  renderProcessFlowchartSvg,
  renderWorkoverTimelineSvg,
} from "../src/services/chartGeneratorService";
import {
  beautifySlideLayout,
  detectCrossSlideContradictions,
  extractDeterministicAgendaAndActions,
  parseMultiWellTable,
} from "../src/services/pptFeaturesService";
import { DEFAULT_SETTINGS } from "../src/types/settings";

describe("chartGeneratorService — Arps Decline, ESP Curve & KPI Cards", () => {
  it("computes exponential and hyperbolic Arps decline rates and EUR", () => {
    const expRes = computeArpsDeclineCurve({
      wellName: "W-101",
      qi: 1200,
      diAnnualPercent: 30,
      bFactor: 0,
      months: 12,
    });
    expect(expRes.points).toHaveLength(13);
    expect(expRes.points[0].rateStbPerDay).toBe(1200);
    expect(expRes.points[12].rateStbPerDay).toBeLessThan(1200);
    expect(expRes.eurMstb).toBeGreaterThan(0);

    const hypRes = computeArpsDeclineCurve({
      wellName: "W-101",
      qi: 1200,
      diAnnualPercent: 30,
      bFactor: 0.5,
      months: 12,
    });
    expect(hypRes.points[12].rateStbPerDay).toBeGreaterThan(
      expRes.points[12].rateStbPerDay
    );
  });

  it("renders valid SVG markup for Decline Curve, ESP Curve, and KPI Cards", () => {
    const dcaSvg = renderDeclineCurveSvg({
      wellName: "W-101",
      qi: 1150,
      diAnnualPercent: 24,
      bFactor: 0.4,
      months: 24,
    });
    expect(dcaSvg).toContain("<svg");
    expect(dcaSvg).toContain("W-101");
    expect(dcaSvg).toContain("MSTB");

    const espSvg = renderEspCurveSvg({
      wellName: "W-104",
      operatingRateBblPerDay: 2100,
      operatingTdhFt: 5400,
      pipPsi: 420,
      frequencyHz: 56,
    });
    expect(espSvg).toContain("<svg");
    expect(espSvg).toContain("ESP");
    expect(espSvg).toContain("56 Hz");

    const kpiSvg = renderKpiCardsSvg([
      { label: "Oil Rate", value: "1,420 bbl/d", status: "normal" },
      { label: "PIP", value: "385 psi", status: "warning" },
    ]);
    expect(kpiSvg).toContain("<svg");
    expect(kpiSvg).toContain("1,420 bbl/d");
  });

  it("renders valid SVG markup for Wellbore Schematic, Process Flowchart, and Workover Timeline", () => {
    const wellboreSvg = renderWellboreSchematicSvg({
      wellName: "WELL-04",
      surfaceCasingDepthFt: 1850,
      productionCasingDepthFt: 8600,
      espPumpDepthFt: 6850,
      perforationTopFt: 7950,
      perforationBottomFt: 8250,
    });
    expect(wellboreSvg).toContain("<svg");
    expect(wellboreSvg).toContain("WELL-04");
    expect(wellboreSvg).toContain("6,850 ft");
    expect(wellboreSvg).toContain("7,950");

    const flowSvg = renderProcessFlowchartSvg({
      title: "CPF Production Train",
      nodes: [
        { id: "n1", label: "Wellhead", subLabel: "Manifold", metric: "420 psig", status: "normal" },
        { id: "n2", label: "Separator", subLabel: "3-Phase", metric: "14,500 BFPD", status: "warning" },
        { id: "n3", label: "Export", subLabel: "Pipeline", metric: "7,450 BOPD", status: "normal" },
      ],
    });
    expect(flowSvg).toContain("<svg");
    expect(flowSvg).toContain("CPF Production Train");
    expect(flowSvg).toContain("14,500 BFPD");

    const timelineSvg = renderWorkoverTimelineSvg({
      title: "Workover Rig Schedule",
      milestones: [
        { phase: "Phase 1", task: "Move & Rig Up", durationOrDate: "Days 1–2", status: "completed" },
        { phase: "Phase 2", task: "RIH 562-Series ESP", durationOrDate: "Days 3–4", status: "active" },
      ],
    });
    expect(timelineSvg).toContain("<svg");
    expect(timelineSvg).toContain("Workover Rig Schedule");
    expect(timelineSvg).toContain("562-Series");
  });
});

describe("pptFeaturesService — Multi-Well Parser, Contradictions, Beautifier & Agenda Builder", () => {
  it("parses multi-well Excel/TSV tables into structured WellBookRow items", () => {
    const raw = [
      "Well\tOil Rate (bbl/d)\tWC (%)\tGOR (scf/stb)\tPIP (psi)\tESP Hz\tStatus",
      "W-101\t1,240\t18%\t620\t450\t55\tRunning Stable",
      "W-102\t890\t34%\t710\t380\t58\tMonitor Gas Lock",
    ].join("\n");

    const parsed = parseMultiWellTable(raw);
    expect(parsed.wells).toHaveLength(2);
    expect(parsed.wells[0].wellName).toBe("W-101");
    expect(parsed.wells[0].metrics["Oil Rate (bbl/d)"]).toBe("1,240");
    expect(parsed.wells[1].wellName).toBe("W-102");
    expect(parsed.wells[1].metrics["ESP Hz"]).toBe("58");
  });

  it("detects cross-slide numerical contradictions for the same well and unit", () => {
    const contradictions = detectCrossSlideContradictions([
      {
        index: 1,
        id: "sl-1",
        title: "Executive Summary",
        fullText: "Well W-101 is producing 1,250 bbl/d with PIP 420 psi.",
        shapes: [],
        speakerNotes: "",
      },
      {
        index: 3,
        id: "sl-3",
        title: "Well W-101 Surveillance",
        fullText: "Current test on W-101 shows 1,080 bbl/d and PIP 420 psi.",
        shapes: [],
        speakerNotes: "",
      },
    ]);

    expect(contradictions.length).toBeGreaterThanOrEqual(1);
    expect(contradictions[0].entity).toContain("W-101");
    expect(contradictions[0].description).toContain("Slide 1");
    expect(contradictions[0].description).toContain("Slide 3");
  });

  it("applies Smart Slide Layout Beautifier in preview mode and returns undo snapshot", async () => {
    const res = await beautifySlideLayout({
      slideNumber: 1,
      style: "two_column",
      settings: DEFAULT_SETTINGS,
    });
    expect(res.summary).toContain("Slide 1");
    expect(res.undoSnapshot).toBeDefined();
  });

  it("extracts deterministic Agenda sections and Action-Tracker Matrix items from slides", () => {
    const plan = extractDeterministicAgendaAndActions([
      {
        index: 1,
        id: "s-1",
        title: "Q3 Production Operations Review",
        fullText: "Q3 Production Operations Review\nField Engineering Team",
        shapes: [],
        speakerNotes: "",
      },
      {
        index: 2,
        id: "s-2",
        title: "ESP Surveillance & Reliability",
        fullText: "ESP Surveillance & Reliability\nAction: Upsize W-104 ESP to 562-Series for +380 bbl/d.",
        shapes: [],
        speakerNotes: "",
      },
      {
        index: 3,
        id: "s-3",
        title: "Water Injection Optimization",
        fullText: "Water Injection Optimization\nRecommend acidizing injector WI-02 to restore 2,500 bwpd.",
        shapes: [],
        speakerNotes: "",
      },
    ]);

    expect(plan.sections.length).toBeGreaterThanOrEqual(2);
    expect(plan.sections[0].title).toBe("Q3 Production Operations Review");
    expect(plan.actions.length).toBeGreaterThanOrEqual(2);
    expect(plan.actions[0].action).toContain("Upsize W-104 ESP");
  });
});

