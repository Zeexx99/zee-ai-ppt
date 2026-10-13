export type DarkModeSetting = 'auto' | 'light' | 'dark';

export type ArabicFontChoice =
  | 'Same as document'
  | 'Simplified Arabic'
  | 'Traditional Arabic'
  | 'Sakkal Majalla'
  | 'Segoe UI'
  | 'Cairo';

export interface CompanyTemplateSettings {
  enabled: boolean;
  brandFont: string;
  primaryColor: string;
  accentColor: string;
  textColor: string;
  slideThemeName: string;
  defaultLayout: 'Title and Content' | 'Two Content' | 'Blank';
}

export interface AppSettings {
  geminiApiKey: string;
  googleClientId: string;
  googleAuthToken: string;
  teamProxyUrl: string;
  teamProxyToken: string;
  model: string;
  darkMode: DarkModeSetting;
  autoRtl: boolean;
  arabicFont: ArabicFontChoice;
  protectedTerms: string[];
  bilingualMode: boolean;
  webSearch: boolean;
  codeExecution: boolean;
  costGuardEnabled: boolean;
  costGuardThreshold: number; // e.g. 12000 tokens warning
  maxAgentSteps: number; // e.g. 5 steps per batch
  companyTemplate: CompanyTemplateSettings;
}

export const DEFAULT_PROTECTED_TERMS = [
  'ESP',
  'TDH',
  'PIP',
  'GOR',
  'BHP',
  'bbl/d',
  'MSCF/d',
  'STB/d',
  'API',
  'WC',
  'BS&W',
  'PI',
  'Pwf',
  'Pws',
  'FBHP',
  'SBHP',
  'IPR',
  'VLP',
  'Hz',
  'psi',
  'psia',
  'psig',
  'cp',
  'GOR',
  'GLR',
  'TVD',
  'MD'
];

export const DEFAULT_SETTINGS: AppSettings = {
  geminiApiKey: '',
  googleClientId: '',
  googleAuthToken: '',
  teamProxyUrl: '',
  teamProxyToken: '',
  model: 'gemini-2.5-flash',
  darkMode: 'auto',
  autoRtl: true,
  arabicFont: 'Same as document',
  protectedTerms: DEFAULT_PROTECTED_TERMS,
  bilingualMode: false,
  webSearch: false,
  codeExecution: false,
  costGuardEnabled: true,
  costGuardThreshold: 12000,
  maxAgentSteps: 5,
  companyTemplate: {
    enabled: false,
    brandFont: 'Segoe UI',
    primaryColor: '#D24726',
    accentColor: '#107C41',
    textColor: '#242424',
    slideThemeName: 'Office Theme',
    defaultLayout: 'Title and Content'
  }
};
