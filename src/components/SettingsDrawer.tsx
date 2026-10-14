import React, { useState } from "react";
import {
  Badge,
  Button,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
  Input,
  Label,
  Select,
  Switch,
} from "@fluentui/react-components";
import { Dismiss20Regular } from "@fluentui/react-icons";
import {
  APP_VERSION,
  ARABIC_FONT_PRESETS,
  ArabicFontPreset,
  AppearanceMode,
  AuthMode,
  AVAILABLE_MODELS,
  BUILD_DATE,
  DEFAULT_PROTECTED_TERMS,
  DirectionMode,
  GeminiModelOption,
  ZeeSettings,
} from "../types/settings";
import { PowerPointApiRequirements } from "../types/ppt";
import {
  fetchLiveGeminiModels,
  startGoogleOAuthSignIn,
} from "../services/authService";

interface SettingsDrawerProps {
  open: boolean;
  settings: ZeeSettings;
  requirements: PowerPointApiRequirements;
  availableLayouts: string[];
  onClose: () => void;
  onUpdateSettings: (partial: Partial<ZeeSettings>) => void;
}

export const SettingsDrawer: React.FC<SettingsDrawerProps> = ({
  open,
  settings,
  requirements,
  availableLayouts,
  onClose,
  onUpdateSettings,
}) => {
  const [oauthStatus, setOauthStatus] = useState<string>("");
  const [modelOptions, setModelOptions] =
    useState<GeminiModelOption[]>(AVAILABLE_MODELS);
  const [modelFetchStatus, setModelFetchStatus] = useState<string>("");
  const [customTermInput, setCustomTermInput] = useState<string>(
    settings.customProtectedTerms.join(", ")
  );

  const handleRefreshModels = async () => {
    setModelFetchStatus("Fetching live models from Google API...");
    try {
      const merged = await fetchLiveGeminiModels(settings);
      setModelOptions(merged);
      setModelFetchStatus(`✓ Loaded ${merged.length} available Gemini models.`);
    } catch (err) {
      setModelFetchStatus(
        `⚠️ ${err instanceof Error ? err.message : String(err)}`
      );
    }
  };

  const handleOAuthClick = async () => {
    setOauthStatus("Opening Google Sign-In...");
    try {
      const res = await startGoogleOAuthSignIn(settings.googleOAuthClientId);
      onUpdateSettings({
        oauthAccessToken: res.accessToken,
        oauthExpiry: Date.now() + res.expiresIn * 1000,
      });
      setOauthStatus("✓ Signed in with Google OAuth!");
    } catch (err) {
      setOauthStatus(
        `Sign-in failed: ${err instanceof Error ? err.message : String(err)}`
      );
    }
  };

  const handleCustomTermsBlur = () => {
    const parsed = customTermInput
      .split(",")
      .map((s) => s.trim())
      .filter(Boolean);
    onUpdateSettings({ customProtectedTerms: parsed });
  };

  return (
    <Dialog open={open} onOpenChange={(_, data) => !data.open && onClose()}>
      <DialogSurface style={{ maxWidth: 440, width: "94vw", maxHeight: "90vh" }}>
        <DialogBody>
          <DialogTitle
            action={
              <Button
                appearance="subtle"
                aria-label="Close settings"
                icon={<Dismiss20Regular />}
                onClick={onClose}
              />
            }
          >
            Zee AI Settings
          </DialogTitle>

          <DialogContent
            style={{
              display: "flex",
              flexDirection: "column",
              gap: 14,
              overflowY: "auto",
              maxHeight: "72vh",
              paddingRight: 4,
              fontSize: 12,
            }}
          >
            {/* 1. AI Connection */}
            <section
              style={{
                border: "1px solid var(--colorNeutralStroke2, #e1dfdd)",
                borderRadius: 8,
                padding: 10,
                display: "flex",
                flexDirection: "column",
                gap: 8,
              }}
            >
              <div style={{ fontWeight: 600, fontSize: 13, color: "#b73618" }}>
                1. AI Connection (Gemini)
              </div>

              <div>
                <Label size="small">Authentication Method</Label>
                <Select
                  size="small"
                  value={settings.authMode}
                  onChange={(_, d) => onUpdateSettings({ authMode: d.value as AuthMode })}
                >
                  <option value="apiKey">Gemini API Key (Local per PC)</option>
                  <option value="proxy">Cloudflare Team Proxy (Shared Server Key)</option>
                  <option value="oauth">Google OAuth Sign-In</option>
                </Select>
              </div>

              {settings.authMode === "apiKey" && (
                <div>
                  <Label size="small">Gemini API Key</Label>
                  <Input
                    size="small"
                    type="password"
                    style={{ width: "100%" }}
                    placeholder="AIzaSy..."
                    value={settings.geminiApiKey}
                    onChange={(_, d) => onUpdateSettings({ geminiApiKey: d.value })}
                  />
                  <div style={{ fontSize: 11, opacity: 0.75, marginTop: 2 }}>
                    Stored locally on this PC under <code>zee-ppt:settings</code>.
                  </div>
                </div>
              )}

              {settings.authMode === "proxy" && (
                <>
                  <div>
                    <Label size="small">Cloudflare Team Proxy URL</Label>
                    <Input
                      size="small"
                      style={{ width: "100%" }}
                      placeholder="https://zee-ai-proxy.your-team.workers.dev"
                      value={settings.proxyUrl}
                      onChange={(_, d) => onUpdateSettings({ proxyUrl: d.value })}
                    />
                  </div>
                  <div>
                    <Label size="small">Team Proxy Token</Label>
                    <Input
                      size="small"
                      type="password"
                      style={{ width: "100%" }}
                      placeholder="Your personal team token"
                      value={settings.proxyToken}
                      onChange={(_, d) => onUpdateSettings({ proxyToken: d.value })}
                    />
                  </div>
                </>
              )}

              {settings.authMode === "oauth" && (
                <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                  <Label size="small">Google OAuth Client ID</Label>
                  <Input
                    size="small"
                    style={{ width: "100%" }}
                    placeholder="123456789-abc.apps.googleusercontent.com"
                    value={settings.googleOAuthClientId}
                    onChange={(_, d) =>
                      onUpdateSettings({ googleOAuthClientId: d.value })
                    }
                  />
                  <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                    <Button
                      size="small"
                      appearance="primary"
                      onClick={handleOAuthClick}
                      style={{ backgroundColor: "#b73618" }}
                    >
                      Sign in with Google
                    </Button>
                    {settings.oauthAccessToken && (
                      <Badge color="success" appearance="tint">
                        Token Active
                      </Badge>
                    )}
                  </div>
                  {oauthStatus && <div style={{ fontSize: 11 }}>{oauthStatus}</div>}
                </div>
              )}

              <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                <div
                  style={{
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "space-between",
                    gap: 8,
                  }}
                >
                  <Label size="small">Gemini Model (Latest Gemini 3 Series)</Label>
                  <Button
                    size="small"
                    appearance="subtle"
                    onClick={handleRefreshModels}
                    style={{ fontSize: 11, padding: "2px 6px" }}
                  >
                    🔄 Fetch Live API Models
                  </Button>
                </div>

                <Select
                  size="small"
                  value={
                    modelOptions.some((m) => m.id === settings.model)
                      ? settings.model
                      : modelOptions[0]?.id
                  }
                  onChange={(_, d) => onUpdateSettings({ model: d.value })}
                >
                  {(
                    [
                      "Gemini 3 (Stable)",
                      "Gemini 3 (Preview)",
                      "Live / Custom",
                    ] as const
                  ).map((grp) => {
                    const items = modelOptions.filter((m) => m.group === grp);
                    if (items.length === 0) return null;
                    return (
                      <optgroup key={grp} label={grp}>
                        {items.map((m) => (
                          <option key={m.id} value={m.id}>
                            {m.label}
                          </option>
                        ))}
                      </optgroup>
                    );
                  })}
                </Select>

                <div>
                  <Label size="small" style={{ fontSize: 11, opacity: 0.85 }}>
                    Active / Custom Model ID (type any new Gemini 3 model endpoint):
                  </Label>
                  <Input
                    size="small"
                    style={{ width: "100%", marginTop: 2 }}
                    placeholder="e.g. gemini-3.8-flash, gemini-3.1-pro-preview, gemini-3.7-flash"
                    value={settings.model}
                    onChange={(_, d) => onUpdateSettings({ model: d.value.trim() })}
                  />
                </div>

                {modelFetchStatus && (
                  <div style={{ fontSize: 11, opacity: 0.85 }}>{modelFetchStatus}</div>
                )}
              </div>
            </section>

            {/* 2. Arabic, RTL, Protected Terms & Bilingual Mode */}
            <section
              style={{
                border: "1px solid var(--colorNeutralStroke2, #e1dfdd)",
                borderRadius: 8,
                padding: 10,
                display: "flex",
                flexDirection: "column",
                gap: 8,
              }}
            >
              <div style={{ fontWeight: 600, fontSize: 13, color: "#b73618" }}>
                2. Arabic / English Direction &amp; Protected Terms
              </div>

              <div>
                <Label size="small">Paragraph Direction &amp; Alignment</Label>
                <Select
                  size="small"
                  value={settings.directionMode}
                  onChange={(_, d) =>
                    onUpdateSettings({ directionMode: d.value as DirectionMode })
                  }
                >
                  <option value="auto">
                    Auto (Arabic → RTL &amp; Right-Aligned, English → LTR &amp; Left-Aligned)
                  </option>
                  <option value="rtl">Force Right-to-Left (RTL)</option>
                  <option value="ltr">Force Left-to-Right (LTR)</option>
                </Select>
              </div>

              <Switch
                checked={settings.centerHeadings}
                onChange={(_, d) => onUpdateSettings({ centerHeadings: d.checked })}
                label="Center slide headings automatically"
              />

              <div>
                <Label size="small">Arabic Font Preset</Label>
                <Select
                  size="small"
                  value={settings.arabicFont}
                  onChange={(_, d) =>
                    onUpdateSettings({ arabicFont: d.value as ArabicFontPreset })
                  }
                >
                  {ARABIC_FONT_PRESETS.map((font) => (
                    <option key={font} value={font}>
                      {font}
                    </option>
                  ))}
                </Select>
              </div>

              <Switch
                checked={settings.bilingualMode}
                onChange={(_, d) => onUpdateSettings({ bilingualMode: d.checked })}
                label="Bilingual Mode (English + Arabic output)"
              />

              <Switch
                checked={settings.protectTermsEnabled}
                onChange={(_, d) =>
                  onUpdateSettings({ protectTermsEnabled: d.checked })
                }
                label="Keep Protected Terms unchanged (ESP, TDH, PIP, GOR, BHP, bbl/d, well names)"
              />

              {settings.protectTermsEnabled && (
                <div>
                  <Label size="small">
                    Extra Custom Protected Terms (comma-separated)
                  </Label>
                  <Input
                    size="small"
                    style={{ width: "100%" }}
                    placeholder="e.g. SCADA, VSD, Arps, Well-19"
                    value={customTermInput}
                    onChange={(_, d) => setCustomTermInput(d.value)}
                    onBlur={handleCustomTermsBlur}
                  />
                  <div style={{ fontSize: 10, opacity: 0.7, marginTop: 3 }}>
                    Built-in: {DEFAULT_PROTECTED_TERMS.slice(0, 12).join(", ")} + well names &amp; units
                  </div>
                </div>
              )}
            </section>

            {/* 3. Company Template Mode */}
            <section
              style={{
                border: "1px solid var(--colorNeutralStroke2, #e1dfdd)",
                borderRadius: 8,
                padding: 10,
                display: "flex",
                flexDirection: "column",
                gap: 8,
              }}
            >
              <div style={{ fontWeight: 600, fontSize: 13, color: "#b73618" }}>
                3. Company Template Mode
              </div>

              <Switch
                checked={settings.companyTemplate.enabled}
                onChange={(_, d) =>
                  onUpdateSettings({
                    companyTemplate: {
                      ...settings.companyTemplate,
                      enabled: d.checked,
                    },
                  })
                }
                label="Always build & format slides with Company Template"
              />

              {settings.companyTemplate.enabled && (
                <>
                  <div>
                    <Label size="small">Preferred Master Layout</Label>
                    <Select
                      size="small"
                      value={settings.companyTemplate.layoutName}
                      onChange={(_, d) =>
                        onUpdateSettings({
                          companyTemplate: {
                            ...settings.companyTemplate,
                            layoutName: d.value,
                          },
                        })
                      }
                    >
                      {availableLayouts.map((layout) => (
                        <option key={layout} value={layout}>
                          {layout}
                        </option>
                      ))}
                    </Select>
                  </div>

                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 6 }}>
                    <div>
                      <Label size="small">Heading Font</Label>
                      <Input
                        size="small"
                        value={settings.companyTemplate.headingFont}
                        onChange={(_, d) =>
                          onUpdateSettings({
                            companyTemplate: {
                              ...settings.companyTemplate,
                              headingFont: d.value,
                            },
                          })
                        }
                      />
                    </div>
                    <div>
                      <Label size="small">Body Font</Label>
                      <Input
                        size="small"
                        value={settings.companyTemplate.bodyFont}
                        onChange={(_, d) =>
                          onUpdateSettings({
                            companyTemplate: {
                              ...settings.companyTemplate,
                              bodyFont: d.value,
                            },
                          })
                        }
                      />
                    </div>
                  </div>

                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 6 }}>
                    <div>
                      <Label size="small">Primary Brand Color</Label>
                      <Input
                        size="small"
                        value={settings.companyTemplate.primaryColor}
                        onChange={(_, d) =>
                          onUpdateSettings({
                            companyTemplate: {
                              ...settings.companyTemplate,
                              primaryColor: d.value,
                            },
                          })
                        }
                      />
                    </div>
                    <div>
                      <Label size="small">Slide Footer Text</Label>
                      <Input
                        size="small"
                        value={settings.companyTemplate.footerText}
                        onChange={(_, d) =>
                          onUpdateSettings({
                            companyTemplate: {
                              ...settings.companyTemplate,
                              footerText: d.value,
                            },
                          })
                        }
                      />
                    </div>
                  </div>
                </>
              )}
            </section>

            {/* 4. Appearance, Cost Guard & Agent */}
            <section
              style={{
                border: "1px solid var(--colorNeutralStroke2, #e1dfdd)",
                borderRadius: 8,
                padding: 10,
                display: "flex",
                flexDirection: "column",
                gap: 8,
              }}
            >
              <div style={{ fontWeight: 600, fontSize: 13, color: "#b73618" }}>
                4. Appearance, Cost Guard &amp; Agent
              </div>

              <div>
                <Label size="small">Appearance Theme</Label>
                <Select
                  size="small"
                  value={settings.appearance}
                  onChange={(_, d) =>
                    onUpdateSettings({ appearance: d.value as AppearanceMode })
                  }
                >
                  <option value="auto">Auto (Follow System/Office)</option>
                  <option value="light">Light Mode</option>
                  <option value="dark">Dark Mode</option>
                </Select>
              </div>

              <Switch
                checked={settings.costGuardEnabled}
                onChange={(_, d) =>
                  onUpdateSettings({ costGuardEnabled: d.checked })
                }
                label="Cost Guard: warn before sending large prompts"
              />

              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 6 }}>
                <div>
                  <Label size="small">Cost Guard Limit (tokens)</Label>
                  <Input
                    size="small"
                    type="number"
                    value={String(settings.costGuardTokenLimit)}
                    onChange={(_, d) =>
                      onUpdateSettings({
                        costGuardTokenLimit: Math.max(1000, Number(d.value) || 12000),
                      })
                    }
                  />
                </div>
                <div>
                  <Label size="small">Agent Step Limit</Label>
                  <Input
                    size="small"
                    type="number"
                    value={String(settings.agentMaxSteps)}
                    onChange={(_, d) =>
                      onUpdateSettings({
                        agentMaxSteps: Math.max(2, Math.min(20, Number(d.value) || 8)),
                      })
                    }
                  />
                </div>
              </div>
            </section>

            {/* 5. Version, Build Date & Requirement Set Diagnostics */}
            <section
              style={{
                border: "1px solid var(--colorNeutralStroke2, #e1dfdd)",
                borderRadius: 8,
                padding: 10,
                fontSize: 11,
                display: "flex",
                flexDirection: "column",
                gap: 4,
              }}
            >
              <div style={{ display: "flex", justifyContent: "space-between", fontWeight: 600 }}>
                <span>Zee AI for PowerPoint v{APP_VERSION}</span>
                <span>Build: {BUILD_DATE}</span>
              </div>
              <div style={{ opacity: 0.78 }}>
                Host: {requirements.isPowerPointHost ? "Microsoft PowerPoint" : "Browser Preview Mode"} |
                Storage prefix: <code>zee-ppt:</code>
              </div>
              <div style={{ display: "flex", flexWrap: "wrap", gap: 4, marginTop: 4 }}>
                {(
                  [
                    ["1.1", requirements.pptApi1_1],
                    ["1.2", requirements.pptApi1_2],
                    ["1.3", requirements.pptApi1_3],
                    ["1.4", requirements.pptApi1_4],
                    ["1.5", requirements.pptApi1_5],
                    ["1.6", requirements.pptApi1_6],
                    ["1.7", requirements.pptApi1_7],
                    ["1.8", requirements.pptApi1_8],
                  ] as [string, boolean][]
                ).map(([ver, ok]) => (
                  <Badge
                    key={ver}
                    size="extra-small"
                    appearance="tint"
                    color={ok ? "success" : "subtle"}
                  >
                    PPT {ver}: {ok ? "✓" : "–"}
                  </Badge>
                ))}
              </div>
            </section>
          </DialogContent>

          <DialogActions>
            <Button
              appearance="primary"
              onClick={onClose}
              style={{ backgroundColor: "#b73618" }}
            >
              Done
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
};
