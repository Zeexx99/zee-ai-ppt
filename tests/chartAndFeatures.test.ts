import { describe, expect, it } from "vitest";
import {
  computeArpsDeclineCurve,
  renderDeclineCurveSvg,
  renderEspCurveSvg,
  renderKpiCardsSvg,
} from "../src/services/chartGeneratorService";
import {
  beautifySlideLayout,
  detectCrossSlideContradictions,
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
});

describe("pptFeaturesService — Multi-Well Parser, Contradictions & Beautifier", () => {
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
});
