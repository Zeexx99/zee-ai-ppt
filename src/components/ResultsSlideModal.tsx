import React, { useState } from 'react';
import {
  Dialog,
  DialogSurface,
  DialogTitle,
  DialogBody,
  DialogActions,
  DialogContent,
  Button,
  Textarea,
  Field
} from '@fluentui/react-components';
import { DocumentTable24Regular, Dismiss24Regular, Sparkle24Regular } from '@fluentui/react-icons';
import { pptFeaturesService } from '../services/pptFeaturesService';
import { AppSettings } from '../types/settings';

interface ResultsSlideModalProps {
  isOpen: boolean;
  onClose: () => void;
  settings: AppSettings;
  onSlideCreated: (slideIndex: number) => void;
}

const ESP_SAMPLE = `Well Name: W-104
ESP Pump: Reda DN1750 (55 stages)
Operating Frequency: 58 Hz
PIP (Pump Intake Pressure): 1,240 psi
PDP (Pump Discharge Pressure): 3,680 psi
TDH (Total Dynamic Head): 5,820 ft
Liquid Rate: 3,450 bbl/d
Water Cut: 42%
GOR: 580 MSCF/bbl
Motor Temp: 198 °F
Motor Current: 64 A`;

const DECLINE_SAMPLE = `Asset: North Field - Block 3
Initial Rate (qi): 12,500 bbl/d
Current Rate: 8,200 bbl/d
Decline Rate (b-factor): 0.45 (Harmonic/Hyperbolic)
Nominal Decline: 14.2% / year
Cumulative Production: 4.8 MMbbl
Estimated Ultimate Recovery (EUR): 18.2 MMbbl
Remaining Reserves: 13.4 MMbbl`;

export const ResultsSlideModal: React.FC<ResultsSlideModalProps> = ({
  isOpen,
  onClose,
  settings,
  onSlideCreated
}) => {
  const [dataText, setDataText] = useState('');
  const [isLoading, setIsLoading] = useState(false);

  const handleBuild = async () => {
    if (!dataText.trim()) return;
    setIsLoading(true);
    try {
      const res = await pptFeaturesService.buildResultsSlideFromData(dataText, settings);
      onSlideCreated(res.slideIndex);
      onClose();
      setDataText('');
    } catch (err) {
      alert(err instanceof Error ? err.message : 'Failed to generate results slide.');
    } finally {
      setIsLoading(false);
    }
  };

  return (
    <Dialog open={isOpen} onOpenChange={(_, d) => !isLoading && onClose()}>
      <DialogSurface style={{ maxWidth: '540px', width: '90vw' }}>
        <DialogBody>
          <DialogTitle
            action={
              <Button
                appearance="subtle"
                aria-label="Close"
                icon={<Dismiss24Regular />}
                onClick={onClose}
                disabled={isLoading}
              />
            }
          >
            Excel / Petroleum Results → Slide
          </DialogTitle>
          <DialogContent style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
            <div style={{ display: 'flex', gap: '6px', flexWrap: 'wrap' }}>
              <span style={{ fontSize: '11px', color: 'var(--colorNeutralForeground3)', alignSelf: 'center' }}>
                Insert Sample:
              </span>
              <Button
                size="small"
                appearance="subtle"
                onClick={() => setDataText(ESP_SAMPLE)}
              >
                ESP Performance
              </Button>
              <Button
                size="small"
                appearance="subtle"
                onClick={() => setDataText(DECLINE_SAMPLE)}
              >
                Decline Curve
              </Button>
            </div>

            <Field
              label="Spreadsheet or Petroleum Test Data"
              hint="Paste tabular data, CSV, or key well parameters"
            >
              <Textarea
                rows={9}
                placeholder="Paste numbers, ESP diagnostic metrics, or test tables..."
                value={dataText}
                onChange={(_, d) => setDataText(d.value)}
              />
            </Field>
          </DialogContent>
          <DialogActions>
            <Button appearance="secondary" onClick={onClose} disabled={isLoading}>
              Cancel
            </Button>
            <Button
              appearance="primary"
              icon={<Sparkle24Regular />}
              disabled={!dataText.trim() || isLoading}
              onClick={handleBuild}
              style={{ background: '#D24726', borderColor: '#D24726' }}
            >
              {isLoading ? 'Creating Results Slide...' : 'Create Results Slide'}
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
};
