import { describe, expect, it } from "vitest";
import {
  appendConvertedUnitsToText,
  containsArabic,
  convertPetroleumUnit,
  detectTextDirection,
  findMissingProtectedTerms,
  formatBilingualTextBlock,
  isPredominantlyArabic,
  maskProtectedTerms,
  resolveArabicFontFamily,
  resolveSlideTextAlignment,
  restoreProtectedTerms,
} from "../src/services/arabicService";

describe("arabicService", () => {
  it("detects Arabic characters and predominant direction accurately", () => {
    expect(containsArabic("North Field ESP Performance")).toBe(false);
    expect(containsArabic("أداء المضخة الغاطسة ESP في البئر WELL-04")).toBe(true);

    expect(isPredominantlyArabic("أداء المضخة الغاطسة ESP في البئر WELL-04")).toBe(true);
    expect(detectTextDirection("أداء المضخة الغاطسة ESP", "auto")).toBe("rtl");
    expect(detectTextDirection("ESP Performance Summary", "auto")).toBe("ltr");
    expect(detectTextDirection("ESP Performance Summary", "rtl")).toBe("rtl");
  });

  it("resolves paragraph alignment with centered headings and RTL/LTR body", () => {
    const arBody = resolveSlideTextAlignment("تحليل إنتاج البئر WELL-01", {
      directionMode: "auto",
      isHeading: false,
      centerHeadings: true,
    });
    expect(arBody.direction).toBe("rtl");
    expect(arBody.alignment).toBe("Right");

    const enBody = resolveSlideTextAlignment("Production rate increased by 320 bbl/d", {
      directionMode: "auto",
      isHeading: false,
      centerHeadings: true,
    });
    expect(enBody.direction).toBe("ltr");
    expect(enBody.alignment).toBe("Left");

    const heading = resolveSlideTextAlignment("ملخص الأداء التشغيلي", {
      directionMode: "auto",
      isHeading: true,
      centerHeadings: true,
    });
    expect(heading.alignment).toBe("Center");
  });

  it("resolves Arabic font preset only when text contains Arabic", () => {
    expect(
      resolveArabicFontFamily("English only text", "Sakkal Majalla", "Calibri")
    ).toBe("Calibri");
    expect(
      resolveArabicFontFamily("نص عربي", "Same as document", "Calibri")
    ).toBe("Calibri");
    expect(
      resolveArabicFontFamily("نص عربي", "Sakkal Majalla", "Calibri")
    ).toBe("Sakkal Majalla");
  });

  it("masks and restores protected petroleum terms, well names, and units", () => {
    const input =
      "Well WELL-04 operates with ESP at PIP 1,120 psi, TDH 5,400 ft, GOR 420 scf/stb, BHP 1,640 psia, rate 2,450 bbl/d.";
    const masked = maskProtectedTerms(input, ["SCADA"]);

    expect(masked.maskedText).not.toContain("WELL-04");
    expect(masked.maskedText).not.toContain("2,450 bbl/d");
    expect(masked.protectedValues).toContain("ESP");
    expect(masked.protectedValues).toContain("TDH");
    expect(masked.protectedValues).toContain("GOR");
    expect(masked.protectedValues).toContain("WELL-04");

    const restored = restoreProtectedTerms(masked.maskedText, masked.placeholders);
    expect(restored).toBe(input);

    const missing = findMissingProtectedTerms(
      input,
      "Well WELL-04 operates with ESP at PIP 1,120 psi and TDH 5,400 ft."
    );
    expect(missing).toContain("GOR");
  });

  it("formats bilingual text blocks cleanly", () => {
    const block = formatBilingualTextBlock(
      "• Target rate: 2,500 bbl/d",
      "• معدل الإنتاج المستهدف: 2,500 bbl/d"
    );
    expect(block).toContain("• Target rate: 2,500 bbl/d");
    expect(block).toContain("• معدل الإنتاج المستهدف: 2,500 bbl/d");
  });

  it("converts petroleum engineering units accurately", () => {
    const bblRes = convertPetroleumUnit(1000, "bbl/d");
    expect(bblRes).not.toBeNull();
    expect(bblRes?.toUnit).toBe("m³/d");
    expect(bblRes?.outputValue).toBeCloseTo(158.99, 1);

    const psiRes = convertPetroleumUnit(1000, "psi");
    expect(psiRes?.toUnit).toBe("bar");
    expect(psiRes?.outputValue).toBeCloseTo(68.95, 1);

    const inline = appendConvertedUnitsToText(
      "WELL-04 produces 1,000 bbl/d at PIP 1,000 psi."
    );
    expect(inline.conversions.length).toBe(2);
    expect(inline.convertedText).toContain("m³/d");
    expect(inline.convertedText).toContain("bar");
  });
});
