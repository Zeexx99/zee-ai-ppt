import {
  ArabicFontPreset,
  DEFAULT_PROTECTED_TERMS,
  DirectionMode,
  ZeeSettings,
} from "../types/settings";

const ARABIC_CHAR_REGEX = /[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]/g;
const LATIN_CHAR_REGEX = /[A-Za-z]/g;

// Matches well names like WELL-01, Well-14A, W-12, PROD-3, INJ-02, ESP-04, AA-101
const WELL_NAME_REGEX = /\b(?:WELL|Well|PROD|Prod|INJ|Inj|field|PAD|[A-Z]{1,4})-\d{1,4}[A-Za-z]?\b/g;

// Matches numbers with petroleum/engineering units (e.g. "2,450 bbl/d", "1850 psi", "450 scf/stb", "62 Hz")
const NUMBER_WITH_UNIT_REGEX =
  /\b\d+(?:,\d{3})*(?:\.\d+)?\s*(?:bbl\/d|stb\/d|MSCF\/d|MMSCF\/d|scf\/stb|m3\/d|sm3\/d|psi|psia|psig|bar|kPa|MPa|ft|m|md|mD|cp|cP|Hz|HP|kW|degF|degC|°F|°C|%)\b/gi;

export function containsArabic(text: string): boolean {
  if (!text) return false;
  ARABIC_CHAR_REGEX.lastIndex = 0;
  return ARABIC_CHAR_REGEX.test(text);
}

export function countArabicAndLatin(text: string): { arabic: number; latin: number } {
  const arabicMatches = text.match(ARABIC_CHAR_REGEX);
  const latinMatches = text.match(LATIN_CHAR_REGEX);
  return {
    arabic: arabicMatches ? arabicMatches.length : 0,
    latin: latinMatches ? latinMatches.length : 0,
  };
}

export function isPredominantlyArabic(text: string): boolean {
  if (!text) return false;
  const { arabic, latin } = countArabicAndLatin(text);
  if (arabic === 0) return false;
  // Even mixed Arabic slides with English engineering acronyms (ESP, BHP, well names)
  // should use RTL when Arabic characters account for at least 30% of letters
  return arabic / (arabic + latin) >= 0.3;
}

export function detectTextDirection(
  text: string,
  mode: DirectionMode = "auto"
): "rtl" | "ltr" {
  if (mode === "rtl") return "rtl";
  if (mode === "ltr") return "ltr";
  return isPredominantlyArabic(text) ? "rtl" : "ltr";
}

export function resolveSlideTextAlignment(
  text: string,
  options: {
    directionMode?: DirectionMode;
    isHeading?: boolean;
    centerHeadings?: boolean;
  } = {}
): {
  direction: "rtl" | "ltr";
  alignment: "Left" | "Right" | "Center";
} {
  const direction = detectTextDirection(text, options.directionMode ?? "auto");
  if (options.isHeading && (options.centerHeadings ?? true)) {
    return { direction, alignment: "Center" };
  }
  return {
    direction,
    alignment: direction === "rtl" ? "Right" : "Left",
  };
}

export function resolveArabicFontFamily(
  text: string,
  arabicFont: ArabicFontPreset,
  fallbackFont?: string
): string | undefined {
  if (!containsArabic(text)) {
    return fallbackFont;
  }
  if (!arabicFont || arabicFont === "Same as document") {
    return fallbackFont;
  }
  return arabicFont;
}

export function getAllProtectedTerms(customTerms: string[] = []): string[] {
  const merged = [...DEFAULT_PROTECTED_TERMS, ...customTerms]
    .map((t) => t.trim())
    .filter(Boolean);
  return Array.from(new Set(merged));
}

function escapeRegExp(str: string): string {
  return str.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

export interface MaskedTextResult {
  maskedText: string;
  placeholders: Record<string, string>;
  protectedValues: string[];
}

/**
 * Masks protected petroleum/engineering terms, well names, and unit quantities
 * with deterministic tokens (__ZEE_PROT_0__, etc.) and can restore them if altered.
 */
export function maskProtectedTerms(
  text: string,
  customTerms: string[] = []
): MaskedTextResult {
  if (!text) {
    return { maskedText: "", placeholders: {}, protectedValues: [] };
  }

  const placeholders: Record<string, string> = {};
  const protectedValues: string[] = [];
  let counter = 0;

  const replaceWithToken = (match: string): string => {
    const token = `__ZEE_PROT_${counter++}__`;
    placeholders[token] = match;
    protectedValues.push(match);
    return token;
  };

  // 1. Protect numeric values with units first
  let working = text.replace(NUMBER_WITH_UNIT_REGEX, replaceWithToken);

  // 2. Protect well names (e.g., WELL-01, W-12, PROD-04)
  working = working.replace(WELL_NAME_REGEX, replaceWithToken);

  // 3. Protect explicit terms (sorted longest first so MMSCF/d matches before MSCF/d)
  const terms = getAllProtectedTerms(customTerms).sort((a, b) => b.length - a.length);
  for (const term of terms) {
    const pattern = new RegExp(`(?<![A-Za-z0-9_])${escapeRegExp(term)}(?![A-Za-z0-9_])`, "g");
    working = working.replace(pattern, replaceWithToken);
  }

  return {
    maskedText: working,
    placeholders,
    protectedValues: Array.from(new Set(protectedValues)),
  };
}

export function restoreProtectedTerms(
  text: string,
  placeholders: Record<string, string>
): string {
  let output = text;
  for (const [token, originalValue] of Object.entries(placeholders)) {
    output = output.split(token).join(originalValue);
  }
  return output;
}

/**
 * Checks whether all protected terms present in the source text are preserved
 * verbatim in the target text.
 */
export function findMissingProtectedTerms(
  sourceText: string,
  targetText: string,
  customTerms: string[] = []
): string[] {
  const { protectedValues } = maskProtectedTerms(sourceText, customTerms);
  return protectedValues.filter((val) => !targetText.includes(val));
}

export function buildProtectedTermsInstruction(settings: ZeeSettings): string {
  if (!settings.protectTermsEnabled) return "";
  const terms = getAllProtectedTerms(settings.customProtectedTerms);
  return [
    "PROTECTED TERMS RULE (CRITICAL):",
    `Keep these oilfield/engineering acronyms, units, and well names EXACTLY unchanged in Latin characters (never translate or alter them): ${terms.join(", ")}, plus all well identifiers (e.g. WELL-01, W-12, Producer-1) and numeric values with units.`,
  ].join("\n");
}

export function buildDirectionAndLanguageInstruction(settings: ZeeSettings): string {
  const parts: string[] = [];

  if (settings.directionMode === "rtl") {
    parts.push(
      "Direction preference: RTL (Right-to-Left, right-aligned paragraphs for body text, centered headings)."
    );
  } else if (settings.directionMode === "ltr") {
    parts.push("Direction preference: LTR (Left-to-Right, left-aligned body text).");
  } else {
    parts.push(
      "Direction preference: Auto-detect per paragraph/slide (Arabic -> RTL & Right-aligned; English -> LTR & Left-aligned; Headings -> Centered)."
    );
  }

  if (settings.arabicFont !== "Same as document") {
    parts.push(`Preferred Arabic font preset: "${settings.arabicFont}".`);
  }

  if (settings.bilingualMode) {
    parts.push(
      "BILINGUAL MODE IS ENABLED: Provide output in both English and Arabic (English bullet/section first, followed by its accurate Arabic counterpart, keeping all protected technical terms and units in Latin script)."
    );
  }

  const protectedRule = buildProtectedTermsInstruction(settings);
  if (protectedRule) {
    parts.push(protectedRule);
  }

  return parts.join("\n");
}

export function formatBilingualTextBlock(englishText: string, arabicText: string): string {
  const en = englishText.trim();
  const ar = arabicText.trim();
  if (!en) return ar;
  if (!ar) return en;
  return `${en}\n—\n${ar}`;
}

export interface UnitConversionResult {
  inputValue: number;
  fromUnit: string;
  toUnit: string;
  outputValue: number;
  formatted: string;
}

const UNIT_CONVERSIONS: Record<
  string,
  { toUnit: string; convert: (v: number) => number; decimals: number }
> = {
  "bbl/d": { toUnit: "m³/d", convert: (v) => v * 0.1589873, decimals: 2 },
  "stb/d": { toUnit: "sm³/d", convert: (v) => v * 0.1589873, decimals: 2 },
  "m3/d": { toUnit: "bbl/d", convert: (v) => v / 0.1589873, decimals: 1 },
  bbl: { toUnit: "m³", convert: (v) => v * 0.1589873, decimals: 2 },
  psi: { toUnit: "bar", convert: (v) => v * 0.06894757, decimals: 2 },
  psia: { toUnit: "bara", convert: (v) => v * 0.06894757, decimals: 2 },
  psig: { toUnit: "barg", convert: (v) => v * 0.06894757, decimals: 2 },
  bar: { toUnit: "psi", convert: (v) => v / 0.06894757, decimals: 1 },
  ft: { toUnit: "m", convert: (v) => v * 0.3048, decimals: 2 },
  m: { toUnit: "ft", convert: (v) => v / 0.3048, decimals: 1 },
  "scf/stb": { toUnit: "sm³/sm³", convert: (v) => v * 0.1781076, decimals: 2 },
  "mscf/d": { toUnit: "10³ sm³/d", convert: (v) => v * 0.02831685, decimals: 2 },
  "mmscf/d": { toUnit: "10⁶ sm³/d", convert: (v) => v * 0.02831685, decimals: 3 },
  degf: { toUnit: "°C", convert: (v) => ((v - 32) * 5) / 9, decimals: 1 },
  "°f": { toUnit: "°C", convert: (v) => ((v - 32) * 5) / 9, decimals: 1 },
  degc: { toUnit: "°F", convert: (v) => (v * 9) / 5 + 32, decimals: 1 },
  "°c": { toUnit: "°F", convert: (v) => (v * 9) / 5 + 32, decimals: 1 },
};

export function convertPetroleumUnit(
  value: number,
  fromUnit: string
): UnitConversionResult | null {
  const key = fromUnit.trim().toLowerCase();
  const spec = UNIT_CONVERSIONS[key];
  if (!spec || Number.isNaN(value)) return null;
  const outputValue = Number(spec.convert(value).toFixed(spec.decimals));
  return {
    inputValue: value,
    fromUnit: fromUnit.trim(),
    toUnit: spec.toUnit,
    outputValue,
    formatted: `${value.toLocaleString("en-US")} ${fromUnit.trim()} (${outputValue.toLocaleString("en-US")} ${spec.toUnit})`,
  };
}

export function appendConvertedUnitsToText(text: string): {
  convertedText: string;
  conversions: UnitConversionResult[];
} {
  const conversions: UnitConversionResult[] = [];
  const regex =
    /\b(\d+(?:,\d{3})*(?:\.\d+)?)\s*(bbl\/d|stb\/d|m3\/d|bbl|psi|psia|psig|bar|ft|m|scf\/stb|MSCF\/d|MMSCF\/d|degF|degC|°F|°C)(?![^\s,.;)]|\s*\()/gi;

  const convertedText = text.replace(regex, (fullMatch, numStr: string, unitStr: string) => {
    const num = parseFloat(numStr.replace(/,/g, ""));
    const res = convertPetroleumUnit(num, unitStr);
    if (!res) return fullMatch;
    conversions.push(res);
    return `${fullMatch} [${res.outputValue.toLocaleString("en-US")} ${res.toUnit}]`;
  });

  return { convertedText, conversions };
}
