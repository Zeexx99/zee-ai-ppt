export type AppearanceMode = "auto" | "light" | "dark";
export type DirectionMode = "auto" | "rtl" | "ltr";
export type AuthMode = "apiKey" | "oauth" | "proxy";

export type ArabicFontPreset =
  | "Same as document"
  | "Simplified Arabic"
  | "Traditional Arabic"
  | "Sakkal Majalla"
  | "Cairo"
  | "Tajawal";

export interface CompanyTemplateConfig {
  enabled: boolean;
  layoutName: string;
  masterName?: string;
  primaryColor: string;
  accentColor: string;
  headingFont: string;
  bodyFont: string;
  footerText: string;
}

export interface ZeeSettings {
  authMode: AuthMode;
  geminiApiKey: string;
  googleOAuthClientId: string;
  oauthAccessToken: string;
  oauthEmail: string;
  oauthExpiry: number;
  proxyUrl: string;
  proxyToken: string;
  model: string;
  temperature: number;
  appearance: AppearanceMode;
  directionMode: DirectionMode;
  centerHeadings: boolean;
  arabicFont: ArabicFontPreset;
  bilingualMode: boolean;
  protectTermsEnabled: boolean;
  customProtectedTerms: string[];
  webSearchEnabled: boolean;
  codeExecutionEnabled: boolean;
  agentMode: boolean;
  agentMaxSteps: number;
  costGuardEnabled: boolean;
  costGuardTokenLimit: number;
  companyTemplate: CompanyTemplateConfig;
}

export const STORAGE_PREFIX = "zee-ppt:";

export const DEFAULT_PROTECTED_TERMS: string[] = [
  "ESP",
  "TDH",
  "PIP",
  "GOR",
  "BHP",
  "WHP",
  "WHT",
  "BHT",
  "PI",
  "WC",
  "BS&W",
  "NPV",
  "IRR",
  "API",
  "bbl/d",
  "stb/d",
  "MSCF/d",
  "MMSCF/d",
  "scf/stb",
  "psi",
  "psia",
  "psig",
  "md",
  "mD",
  "cp",
  "cP",
];

export const ARABIC_FONT_PRESETS: ArabicFontPreset[] = [
  "Same as document",
  "Simplified Arabic",
  "Traditional Arabic",
  "Sakkal Majalla",
  "Cairo",
  "Tajawal",
];

export const AVAILABLE_MODELS = [
  { id: "gemini-2.5-flash", label: "Gemini 2.5 Flash (Fast & Recommended)" },
  { id: "gemini-2.5-pro", label: "Gemini 2.5 Pro (Deep Reasoning)" },
  { id: "gemini-2.0-flash", label: "Gemini 2.0 Flash" },
];

export const DEFAULT_SETTINGS: ZeeSettings = {
  authMode: "apiKey",
  geminiApiKey: "",
  googleOAuthClientId: "",
  oauthAccessToken: "",
  oauthEmail: "",
  oauthExpiry: 0,
  proxyUrl: "",
  proxyToken: "",
  model: "gemini-2.5-flash",
  temperature: 0.3,
  appearance: "auto",
  directionMode: "auto",
  centerHeadings: true,
  arabicFont: "Same as document",
  bilingualMode: false,
  protectTermsEnabled: true,
  customProtectedTerms: [],
  webSearchEnabled: false,
  codeExecutionEnabled: false,
  agentMode: false,
  agentMaxSteps: 8,
  costGuardEnabled: true,
  costGuardTokenLimit: 12000,
  companyTemplate: {
    enabled: false,
    layoutName: "Title and Content",
    masterName: "Office Theme",
    primaryColor: "#B73618",
    accentColor: "#0F6CBD",
    headingFont: "Calibri Light",
    bodyFont: "Calibri",
    footerText: "Confidential — Zee Engineering",
  },
};

declare const __APP_VERSION__: string | undefined;
declare const __BUILD_DATE__: string | undefined;

export const APP_VERSION =
  typeof __APP_VERSION__ !== "undefined" ? __APP_VERSION__ : "1.2.0";
export const BUILD_DATE =
  typeof __BUILD_DATE__ !== "undefined" ? __BUILD_DATE__ : "2026-10-13";
