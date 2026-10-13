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
  Input,
  Select,
  Field,
  Card,
  ProgressBar,
  Badge
} from '@fluentui/react-components';
import { Sparkle24Regular, Dismiss24Regular, Add24Regular, Delete24Regular } from '@fluentui/react-icons';
import { DeckOutline, pptFeaturesService } from '../services/pptFeaturesService';
import { AppSettings } from '../types/settings';

interface PowerPointDeckModalProps {
  isOpen: boolean;
  onClose: () => void;
  settings: AppSettings;
  onDeckBuilt: (slideCount: number) => void;
}

export const PowerPointDeckModal: React.FC<PowerPointDeckModalProps> = ({
  isOpen,
  onClose,
  settings,
  onDeckBuilt
}) => {
  const [reportText, setReportText] = useState('');
  const [slideCount, setSlideCount] = useState<number>(5);
  const [isGeneratingOutline, setIsGeneratingOutline] = useState(false);
  const [isBuildingDeck, setIsBuildingDeck] = useState(false);
  const [outline, setOutline] = useState<DeckOutline | null>(null);
  const [progressStatus, setProgressStatus] = useState<string>('');
  const [progressPercent, setProgressPercent] = useState<number>(0);

  const handleGenerateOutline = async () => {
    if (!reportText.trim()) return;
    setIsGeneratingOutline(true);
    try {
      const result = await pptFeaturesService.generateDeckOutline(reportText, slideCount, settings);
      setOutline(result);
    } catch (err) {
      alert(err instanceof Error ? err.message : 'Failed to generate outline.');
    } finally {
      setIsGeneratingOutline(false);
    }
  };

  const handleBuildDeck = async () => {
    if (!outline) return;
    setIsBuildingDeck(true);
    try {
      const res = await pptFeaturesService.buildDeckFromOutline(
        outline,
        (current, total, title) => {
          setProgressPercent(Math.round((current / total) * 100));
          setProgressStatus(`Building slide ${current} of ${total}: "${title}"...`);
        }
      );
      onDeckBuilt(res.createdCount);
      onClose();
      setOutline(null);
      setReportText('');
    } catch (err) {
      alert(err instanceof Error ? err.message : 'Failed to build presentation deck.');
    } finally {
      setIsBuildingDeck(false);
      setProgressPercent(0);
      setProgressStatus('');
    }
  };

  const updateSlideTitle = (idx: number, newTitle: string) => {
    if (!outline) return;
    const slides = [...outline.slides];
    slides[idx].title = newTitle;
    setOutline({ ...outline, slides });
  };

  const updateSlideBullet = (slideIdx: number, bulletIdx: number, text: string) => {
    if (!outline) return;
    const slides = [...outline.slides];
    slides[slideIdx].bulletPoints[bulletIdx] = text;
    setOutline({ ...outline, slides });
  };

  return (
    <Dialog open={isOpen} onOpenChange={(_, d) => !isBuildingDeck && !isGeneratingOutline && onClose()}>
      <DialogSurface style={{ maxWidth: '640px', width: '92vw', maxHeight: '85vh' }}>
        <DialogBody>
          <DialogTitle
            action={
              <Button
                appearance="subtle"
                aria-label="Close"
                icon={<Dismiss24Regular />}
                onClick={onClose}
                disabled={isBuildingDeck || isGeneratingOutline}
              />
            }
          >
            Report or PDF → Presentation Deck
          </DialogTitle>
          <DialogContent style={{ display: 'flex', flexDirection: 'column', gap: '14px', maxHeight: '60vh', overflowY: 'auto' }}>
            {!outline ? (
              <>
                <Field label="Source Report, Article, or PDF Text" hint="Paste the document text you want to transform into slides">
                  <Textarea
                    rows={8}
                    placeholder="Paste the technical report, executive summary, or PDF excerpt here..."
                    value={reportText}
                    onChange={(_, d) => setReportText(d.value)}
                  />
                </Field>
                <div style={{ display: 'flex', gap: '12px', alignItems: 'center' }}>
                  <Field label="Desired Slide Count">
                    <Select
                      value={String(slideCount)}
                      onChange={e => setSlideCount(Number(e.target.value))}
                    >
                      <option value="3">3 Slides (Quick Briefing)</option>
                      <option value="5">5 Slides (Executive Summary)</option>
                      <option value="7">7 Slides (Standard Deck)</option>
                      <option value="10">10 Slides (Comprehensive)</option>
                    </Select>
                  </Field>
                </div>
              </>
            ) : (
              <>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <div>
                    <span style={{ fontWeight: 600, fontSize: '15px' }}>{outline.presentationTitle}</span>
                    <div style={{ fontSize: '12px', color: 'var(--colorNeutralForeground3)' }}>
                      Review and edit the outline before building slides.
                    </div>
                  </div>
                  <Button size="small" appearance="subtle" onClick={() => setOutline(null)}>
                    Start Over
                  </Button>
                </div>

                {outline.slides.map((s, sIdx) => (
                  <Card key={s.slideNumber} style={{ padding: '10px', gap: '6px' }}>
                    <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                      <Badge color="brand">Slide {s.slideNumber}</Badge>
                      {s.visualSuggestion && (
                        <span style={{ fontSize: '11px', color: 'var(--colorNeutralForeground3)' }}>
                          💡 {s.visualSuggestion}
                        </span>
                      )}
                    </div>
                    <Input
                      value={s.title}
                      onChange={(_, d) => updateSlideTitle(sIdx, d.value)}
                      style={{ fontWeight: 600 }}
                    />
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                      {s.bulletPoints.map((b, bIdx) => (
                        <Input
                          key={bIdx}
                          size="small"
                          value={b}
                          onChange={(_, d) => updateSlideBullet(sIdx, bIdx, d.value)}
                        />
                      ))}
                    </div>
                    {s.notes && (
                      <div style={{ fontSize: '11px', color: 'var(--colorNeutralForeground3)', background: 'var(--colorNeutralBackground3)', padding: '4px 6px', borderRadius: '4px' }}>
                        <b>Speaker Notes:</b> {s.notes}
                      </div>
                    )}
                  </Card>
                ))}
              </>
            )}

            {isBuildingDeck && (
              <div style={{ display: 'flex', flexDirection: 'column', gap: '8px', padding: '12px 0' }}>
                <span style={{ fontSize: '12px', fontWeight: 600 }}>{progressStatus}</span>
                <ProgressBar value={progressPercent / 100} />
              </div>
            )}
          </DialogContent>
          <DialogActions>
            <Button
              appearance="secondary"
              onClick={onClose}
              disabled={isGeneratingOutline || isBuildingDeck}
            >
              Cancel
            </Button>
            {!outline ? (
              <Button
                appearance="primary"
                icon={<Sparkle24Regular />}
                disabled={!reportText.trim() || isGeneratingOutline}
                onClick={handleGenerateOutline}
                style={{ background: '#D24726', borderColor: '#D24726' }}
              >
                {isGeneratingOutline ? 'Generating Outline...' : 'Generate Outline'}
              </Button>
            ) : (
              <Button
                appearance="primary"
                icon={<Sparkle24Regular />}
                disabled={isBuildingDeck}
                onClick={handleBuildDeck}
                style={{ background: '#107C41', borderColor: '#107C41' }}
              >
                {isBuildingDeck ? 'Building Slides...' : 'Approve & Build Deck'}
              </Button>
            )}
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
};
