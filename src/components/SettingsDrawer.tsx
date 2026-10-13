import React, { useState } from 'react';
import {
  Drawer,
  DrawerHeader,
  DrawerHeaderTitle,
  DrawerBody,
  Button,
  Input,
  Label,
  Switch,
  Select,
  Field,
  Divider,
  Tag,
  TagGroup,
  Badge,
  TabList,
  Tab,
  TabValue
} from '@fluentui/react-components';
import {
  Dismiss24Regular,
  Save24Regular,
  Key24Regular,
  ShieldCheckmark24Regular,
  Globe24Regular,
  TextDirectionHorizontalLtr24Regular,
  Eye24Regular,
  EyeOff24Regular,
  Delete24Regular
} from '@fluentui/react-icons';
import { AppSettings, ArabicFontChoice, DarkModeSetting, DEFAULT_PROTECTED_TERMS } from '../types/settings';
import { RequirementSetStatus } from '../types/ppt';
import { authService } from '../services/authService';

interface SettingsDrawerProps {
  isOpen: boolean;
  onClose: () => void;
  settings: AppSettings;
  onSave: (updated: AppSettings) => void;
  requirementStatus: RequirementSetStatus;
}

export const SettingsDrawer: React.FC<SettingsDrawerProps> = ({
  isOpen,
  onClose,
  settings,
  onSave,
  requirementStatus
}) => {
  const [formData, setFormData] = useState<AppSettings>({ ...settings });
  const [showKey, setShowKey] = useState(false);
  const [newTerm, setNewTerm] = useState('');
  const [activeTab, setActiveTab] = useState<TabValue>('connection');
  const [authStatus, setAuthStatus] = useState<string>('');

  // Sync if settings change externally
  React.useEffect(() => {
    setFormData({ ...settings });
  }, [settings]);

  const handleChange = <K extends keyof AppSettings>(key: K, val: AppSettings[K]) => {
    setFormData(prev => ({ ...prev, [key]: val }));
  };

  const handleSaveAndClose = () => {
    onSave(formData);
    onClose();
  };

  const handleAddTerm = () => {
    const trimmed = newTerm.trim();
    if (trimmed && !formData.protectedTerms.includes(trimmed)) {
      setFormData(prev => ({
        ...prev,
        protectedTerms: [...prev.protectedTerms, trimmed]
      }));
      setNewTerm('');
    }
  };

  const handleRemoveTerm = (termToRemove: string) => {
    setFormData(prev => ({
      ...prev,
      protectedTerms: prev.protectedTerms.filter(t => t !== termToRemove)
    }));
  };

  const handleResetTerms = () => {
    setFormData(prev => ({
      ...prev,
      protectedTerms: [...DEFAULT_PROTECTED_TERMS]
    }));
  };

  const handleGoogleSignIn = async () => {
    setAuthStatus('Connecting to Google...');
    const res = await authService.signInWithGoogle(formData.googleClientId);
    if (res.token) {
      handleChange('googleAuthToken', res.token);
      setAuthStatus('Signed in successfully!');
    } else {
      setAuthStatus(`Sign-in failed: ${res.error || 'Unknown error'}`);
    }
  };

  const handleSignOutGoogle = () => {
    handleChange('googleAuthToken', '');
    setAuthStatus('Signed out.');
  };

  return (
    <Drawer
      type="overlay"
      position="end"
      open={isOpen}
      onOpenChange={(_, data) => !data.open && onClose()}
      style={{ width: '340px', maxWidth: '90vw' }}
    >
      <DrawerHeader style={{ borderBottom: '1px solid rgba(128,128,128,0.2)' }}>
        <DrawerHeaderTitle
          action={
            <Button
              appearance="subtle"
              aria-label="Close"
              icon={<Dismiss24Regular />}
              onClick={onClose}
            />
          }
        >
          Zee AI Settings
        </DrawerHeaderTitle>
      </DrawerHeader>

      <TabList
        selectedValue={activeTab}
        onTabSelect={(_, data) => setActiveTab(data.value)}
        style={{ padding: '0 8px', borderBottom: '1px solid rgba(128,128,128,0.1)' }}
      >
        <Tab value="connection">AI & Auth</Tab>
        <Tab value="presentation">Slides & RTL</Tab>
        <Tab value="advanced">Advanced</Tab>
      </TabList>

      <DrawerBody style={{ padding: '16px', display: 'flex', flexDirection: 'column', gap: '16px' }}>
        {activeTab === 'connection' && (
          <>
            {/* Gemini API Key */}
            <Field label="Gemini API Key" hint="Stored locally on this PC only">
              <div style={{ display: 'flex', gap: '4px' }}>
                <Input
                  type={showKey ? 'text' : 'password'}
                  value={formData.geminiApiKey}
                  onChange={(_, data) => handleChange('geminiApiKey', data.value)}
                  placeholder="AIzaSy..."
                  style={{ flex: 1 }}
                />
                <Button
                  appearance="subtle"
                  icon={showKey ? <EyeOff24Regular /> : <Eye24Regular />}
                  onClick={() => setShowKey(!showKey)}
                  aria-label="Toggle API key visibility"
                />
              </div>
            </Field>

            <Divider style={{ margin: '4px 0' }} />

            {/* Cloudflare Team Proxy */}
            <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
              <Label weight="semibold">Team Proxy (Cloudflare Worker)</Label>
              <Input
                placeholder="https://zee-proxy.workers.dev"
                value={formData.teamProxyUrl}
                onChange={(_, data) => handleChange('teamProxyUrl', data.value)}
              />
              <Input
                type="password"
                placeholder="Team Proxy Token (optional)"
                value={formData.teamProxyToken}
                onChange={(_, data) => handleChange('teamProxyToken', data.value)}
              />
              <span style={{ fontSize: '11px', color: 'var(--colorNeutralForeground3)' }}>
                Allows using a shared server key with origin locking and rate limits.
              </span>
            </div>

            <Divider style={{ margin: '4px 0' }} />

            {/* Google OAuth Sign-in */}
            <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
              <Label weight="semibold">Google Account (Optional)</Label>
              <Input
                placeholder="Google OAuth Client ID"
                value={formData.googleClientId}
                onChange={(_, data) => handleChange('googleClientId', data.value)}
              />
              {formData.googleAuthToken ? (
                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                  <Badge color="success">Signed in</Badge>
                  <Button size="small" appearance="subtle" onClick={handleSignOutGoogle}>
                    Sign Out
                  </Button>
                </div>
              ) : (
                <Button size="small" appearance="secondary" onClick={handleGoogleSignIn}>
                  Sign in with Google
                </Button>
              )}
              {authStatus && (
                <span style={{ fontSize: '11px', color: 'var(--colorNeutralForeground3)' }}>
                  {authStatus}
                </span>
              )}
            </div>

            <Divider style={{ margin: '4px 0' }} />

            {/* Model Selection */}
            <Field label="Gemini Model">
              <Select
                value={formData.model}
                onChange={e => handleChange('model', e.target.value)}
              >
                <option value="gemini-2.5-flash">Gemini 2.5 Flash (Fastest, Recommended)</option>
                <option value="gemini-2.0-flash">Gemini 2.0 Flash</option>
                <option value="gemini-1.5-pro">Gemini 1.5 Pro (Deep reasoning)</option>
                <option value="gemini-1.5-flash">Gemini 1.5 Flash</option>
              </Select>
            </Field>
          </>
        )}

        {activeTab === 'presentation' && (
          <>
            {/* Arabic & RTL */}
            <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
              <Label weight="semibold">Arabic & Direction</Label>
              <Switch
                label="Auto RTL Direction & Right Alignment"
                checked={formData.autoRtl}
                onChange={(_, data) => handleChange('autoRtl', data.checked)}
              />
              <Switch
                label="Bilingual Mode (English + Arabic output)"
                checked={formData.bilingualMode}
                onChange={(_, data) => handleChange('bilingualMode', data.checked)}
              />

              <Field label="Arabic Font Preset">
                <Select
                  value={formData.arabicFont}
                  onChange={e => handleChange('arabicFont', e.target.value as ArabicFontChoice)}
                >
                  <option value="Same as document">Same as document</option>
                  <option value="Simplified Arabic">Simplified Arabic</option>
                  <option value="Traditional Arabic">Traditional Arabic</option>
                  <option value="Sakkal Majalla">Sakkal Majalla</option>
                  <option value="Segoe UI">Segoe UI</option>
                  <option value="Cairo">Cairo</option>
                </Select>
              </Field>
            </div>

            <Divider style={{ margin: '4px 0' }} />

            {/* Protected Terms */}
            <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <Label weight="semibold">Protected Terms & Units</Label>
                <Button size="small" appearance="transparent" onClick={handleResetTerms}>
                  Reset
                </Button>
              </div>
              <span style={{ fontSize: '11px', color: 'var(--colorNeutralForeground3)' }}>
                These terms will never be translated, altered, or lowercased.
              </span>
              <div style={{ display: 'flex', gap: '4px' }}>
                <Input
                  size="small"
                  placeholder="Add acronym/well (e.g. ESP)"
                  value={newTerm}
                  onChange={(_, data) => setNewTerm(data.value)}
                  onKeyDown={e => e.key === 'Enter' && handleAddTerm()}
                  style={{ flex: 1 }}
                />
                <Button size="small" onClick={handleAddTerm}>
                  Add
                </Button>
              </div>
              <div
                style={{
                  display: 'flex',
                  flexWrap: 'wrap',
                  gap: '4px',
                  maxHeight: '120px',
                  overflowY: 'auto',
                  padding: '4px',
                  border: '1px solid rgba(128,128,128,0.2)',
                  borderRadius: '4px'
                }}
              >
                {formData.protectedTerms.map(term => (
                  <Tag
                    key={term}
                    size="small"
                    dismissible
                    dismissIcon={<Dismiss24Regular />}
                    onClick={() => handleRemoveTerm(term)}
                  >
                    {term}
                  </Tag>
                ))}
              </div>
            </div>

            <Divider style={{ margin: '4px 0' }} />

            {/* Company Template Mode */}
            <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
              <Label weight="semibold">Company Template Mode</Label>
              <Switch
                label="Apply Brand Guidelines"
                checked={formData.companyTemplate.enabled}
                onChange={(_, data) =>
                  setFormData(prev => ({
                    ...prev,
                    companyTemplate: { ...prev.companyTemplate, enabled: data.checked }
                  }))
                }
              />
              {formData.companyTemplate.enabled && (
                <div style={{ display: 'flex', flexDirection: 'column', gap: '6px', paddingLeft: '8px' }}>
                  <Field label="Brand Font">
                    <Input
                      size="small"
                      value={formData.companyTemplate.brandFont}
                      onChange={(_, d) =>
                        setFormData(p => ({
                          ...p,
                          companyTemplate: { ...p.companyTemplate, brandFont: d.value }
                        }))
                      }
                    />
                  </Field>
                  <div style={{ display: 'flex', gap: '8px' }}>
                    <Field label="Title Color" style={{ flex: 1 }}>
                      <Input
                        size="small"
                        value={formData.companyTemplate.primaryColor}
                        onChange={(_, d) =>
                          setFormData(p => ({
                            ...p,
                            companyTemplate: { ...p.companyTemplate, primaryColor: d.value }
                          }))
                        }
                      />
                    </Field>
                    <Field label="Accent Color" style={{ flex: 1 }}>
                      <Input
                        size="small"
                        value={formData.companyTemplate.accentColor}
                        onChange={(_, d) =>
                          setFormData(p => ({
                            ...p,
                            companyTemplate: { ...p.companyTemplate, accentColor: d.value }
                          }))
                        }
                      />
                    </Field>
                  </div>
                </div>
              )}
            </div>
          </>
        )}

        {activeTab === 'advanced' && (
          <>
            {/* Appearance */}
            <Field label="Appearance">
              <Select
                value={formData.darkMode}
                onChange={e => handleChange('darkMode', e.target.value as DarkModeSetting)}
              >
                <option value="auto">Auto (Match Office / System)</option>
                <option value="light">Light</option>
                <option value="dark">Dark</option>
              </Select>
            </Field>

            <Divider style={{ margin: '4px 0' }} />

            {/* Gemini Tools & Cost Guard */}
            <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
              <Label weight="semibold">Tools & Safety</Label>
              <Switch
                label="Google Web Search"
                checked={formData.webSearch}
                onChange={(_, data) => handleChange('webSearch', data.checked)}
              />
              <Switch
                label="Python Code Execution"
                checked={formData.codeExecution}
                onChange={(_, data) => handleChange('codeExecution', data.checked)}
              />
              <Switch
                label="Cost Guard Token Warning"
                checked={formData.costGuardEnabled}
                onChange={(_, data) => handleChange('costGuardEnabled', data.checked)}
              />
              {formData.costGuardEnabled && (
                <Field label="Token Warning Threshold" hint="Warns before sending very large prompts">
                  <Input
                    type="number"
                    size="small"
                    value={String(formData.costGuardThreshold)}
                    onChange={(_, data) => handleChange('costGuardThreshold', Number(data.value) || 12000)}
                  />
                </Field>
              )}
              <Field label="Agent Step Limit" hint="Requires clicking 'Continue' after this many actions">
                <Input
                  type="number"
                  size="small"
                  value={String(formData.maxAgentSteps)}
                  onChange={(_, data) => handleChange('maxAgentSteps', Number(data.value) || 5)}
                />
              </Field>
            </div>

            <Divider style={{ margin: '4px 0' }} />

            {/* PowerPoint API Diagnostic Status */}
            <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
              <Label weight="semibold">PowerPointApi Capability Status</Label>
              <div
                style={{
                  fontSize: '11px',
                  background: 'var(--colorNeutralBackground3)',
                  padding: '8px',
                  borderRadius: '4px',
                  display: 'flex',
                  flexDirection: 'column',
                  gap: '4px'
                }}
              >
                <div>Highest Api Set: <b>{requirementStatus.highestVersion}</b></div>
                <div>Shape & Text Formatting (1.3+): {requirementStatus.canFormatShapes ? '✅ Supported' : '❌ Limited'}</div>
                <div>Tables Creation (1.4+): {requirementStatus.canAddTables ? '✅ Supported' : '❌ Hidden'}</div>
                <div>Speaker Notes Access (1.5+): {requirementStatus.canManageNotes ? '✅ Supported' : '❌ Hidden'}</div>
                <div>Base64 Slides (1.2+): {requirementStatus.canInsertSlidesBase64 ? '✅ Supported' : '❌ Limited'}</div>
              </div>
            </div>
          </>
        )}

        {/* Footer with save button & version */}
        <div style={{ marginTop: 'auto', display: 'flex', flexDirection: 'column', gap: '8px', paddingTop: '12px' }}>
          <Button
            appearance="primary"
            icon={<Save24Regular />}
            onClick={handleSaveAndClose}
            style={{ width: '100%' }}
          >
            Save Settings
          </Button>
          <div style={{ textAlign: 'center', fontSize: '10px', color: 'var(--colorNeutralForeground4)' }}>
            Zee AI for PowerPoint • v1.0.0 (Oct 2026)
          </div>
        </div>
      </DrawerBody>
    </Drawer>
  );
};
