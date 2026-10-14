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

export interface GeminiModelOption {
  id: string;
  label: string;
  group?: "Gemini 3 (Stable)" | "Gemini 3 (Preview)" | "Live / Custom";
}

export const AVAILABLE_MODELS: GeminiModelOption[] = [
  {
    id: "gemini-3.8-flash",
    label: "Gemini 3.8 Flash (Latest Stable — Agents & Complex Workflows)",
    group: "Gemini 3 (Stable)",
  },
  {
    id: "gemini-3.1-pro-preview",
    label: "Gemini 3.1 Pro (Preview — Deep Reasoning & Coding)",
    group: "Gemini 3 (Preview)",
  },
  {
    id: "gemini-3.7-flash",
    label: "Gemini 3.7 Flash (Stable — Multi-Step Agentic)",
    group: "Gemini 3 (Stable)",
  },
  {
    id: "gemini-3.6-flash",
    label: "Gemini 3.6 Flash (Stable — Balanced Multimodal)",
    group: "Gemini 3 (Stable)",
  },
  {
    id: "gemini-3.5-flash",
    label: "Gemini 3.5 Flash (Stable — High-Throughput)",
    group: "Gemini 3 (Stable)",
  },
  {
    id: "gemini-3.5-flash-lite",
    label: "Gemini 3.5 Flash-Lite (Fastest & Ultra-Low Cost)",
    group: "Gemini 3 (Stable)",
  },
  {
    id: "gemini-3.1-flash-lite",
    label: "Gemini 3.1 Flash-Lite (Frontier Efficiency)",
    group: "Gemini 3 (Stable)",
  },
  {
    id: "gemini-3-flash-preview",
    label: "Gemini 3 Flash (Preview)",
    group: "Gemini 3 (Preview)",
  },
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
  model: "gemini-3.8-flash",
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
  typeof __APP_VERSION__ !== "undefined" ? __APP_VERSION__ : "1.5.0";
export const BUILD_DATE =
  typeof __BUILD_DATE__ !== "undefined" ? __BUILD_DATE__ : "2026-10-14";
