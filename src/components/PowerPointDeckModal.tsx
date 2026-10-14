import React, { useState } from "react";
import {
  Badge,
  Button,
  Checkbox,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
  Input,
  Label,
  Select,
  Spinner,
  Textarea,
} from "@fluentui/react-components";
import { Dismiss20Regular } from "@fluentui/react-icons";
import { ChatAttachment, SlideUndoSnapshot } from "../types/chat";
import {
  ConsistencyIssue,
  DeckOutlinePlan,
  DeckReviewReport,
  SlideBeautifyStyle,
  TextFitIssue,
} from "../types/ppt";
import { ZeeSettings } from "../types/settings";
import { fileToChatAttachment } from "./ChatInput";
import {
  applyConsistencyFixes,
  beautifySlideLayout,
  buildBackupSlidesFromQna,
  buildSlidesFromApprovedOutline,
  fixTextFitIssues,
  generateDeckOutlineFromReport,
  runManagementDeckReview,
  runPresentationConsistencyCheck,
  scanPresentationForTextFitIssues,
} from "../services/pptFeaturesService";
import { getActiveSlideIndex } from "../services/pptService";

export type DeckToolMode =
  | "report_to_deck"
  | "consistency"
  | "text_fit"
  | "deck_review"
  | "beautify"
  | null;

interface PowerPointDeckModalProps {
  mode: DeckToolMode;
  settings: ZeeSettings;
  onClose: () => void;
  onCompleted: (summaryMessage: string, undoSnapshots: SlideUndoSnapshot[]) => void;
}

export const PowerPointDeckModal: React.FC<PowerPointDeckModalProps> = ({
  mode,
  settings,
  onClose,
  onCompleted,
}) => {
  // Report -> Deck states
  const [reportText, setReportText] = useState("");
  const [targetSlideCount, setTargetSlideCount] = useState(6);
  const [attachments, setAttachments] = useState<ChatAttachment[]>([]);
  const [outlinePlan, setOutlinePlan] = useState<DeckOutlinePlan | null>(null);
  const [buildProgress, setBuildProgress] = useState("");

  // Consistency Checker states
  const [consistencyIssues, setConsistencyIssues] = useState<ConsistencyIssue[]>([]);
  const [consistencyScanned, setConsistencyScanned] = useState(false);

  // Text-Fit Fixer states
  const [textFitIssues, setTextFitIssues] = useState<TextFitIssue[]>([]);
  const [textFitScanned, setTextFitScanned] = useState(false);

  // Management Q&A & Deck Reviewer states
  const [reviewReport, setReviewReport] = useState<DeckReviewReport | null>(null);

  // Smart Slide Layout Beautifier states
  const [beautifySlideIdx, setBeautifySlideIdx] = useState<number>(1);
  const [beautifyStyle, setBeautifyStyle] =
    useState<SlideBeautifyStyle>("two_column");

  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");

  if (!mode) return null;

  const handleGenerateOutline = async () => {
    if (!reportText.trim() && attachments.length === 0) {
      setError("Please paste report text or attach a PDF/document first.");
      return;
    }
    setBusy(true);
    setError("");
    try {
      const plan = await generateDeckOutlineFromReport({
        reportText,
        attachments,
        targetSlideCount,
        settings,
      });
      setOutlinePlan(plan);
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setBusy(false);
    }
  };

  const handleApproveAndBuildDeck = async () => {
    if (!outlinePlan) return;
    setBusy(true);
    setError("");
    try {
      const { createdCount, undoSnapshots } = await buildSlidesFromApprovedOutline(
        outlinePlan,
        settings,
        (current, total, title) => {
          setBuildProgress(`Building slide ${current}/${total}: "${title}"...`);
        }
      );
      setBuildProgress("");
      onCompleted(
        `Built ${createdCount} slides from approved outline "${outlinePlan.deckTitle}".`,
        undoSnapshots
      );
      onClose();
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setBusy(false);
    }
  };

  const handleRunConsistencyScan = async () => {
    setBusy(true);
    setError("");
    try {
      const found = await runPresentationConsistencyCheck(settings);
      setConsistencyIssues(found);
      setConsistencyScanned(true);
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setBusy(false);
    }
  };

  const handleFixConsistency = async () => {
    setBusy(true);
    setError("");
    try {
      const { fixedCount, undoSnapshots } = await applyConsistencyFixes(
        consistencyIssues,
        settings
      );
      onCompleted(
        `Consistency Checker fixed ${fixedCount} issue(s) across the presentation.`,
        undoSnapshots
      );
      onClose();
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setBusy(false);
    }
  };

  const handleRunTextFitScan = async () => {
    setBusy(true);
    setError("");
    try {
      const found = await scanPresentationForTextFitIssues();
      setTextFitIssues(found);
      setTextFitScanned(true);
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setBusy(false);
    }
  };

  const handleFixTextFit = async () => {
    setBusy(true);
    setError("");
    try {
      const { fixedCount, undoSnapshots } = await fixTextFitIssues(
        textFitIssues,
        settings
      );
      onCompleted(
        `Text-Fit Fixer shortened and resized ${fixedCount} overflowing placeholder(s).`,
        undoSnapshots
      );
      onClose();
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setBusy(false);
    }
  };

  const handleRunDeckReview = async () => {
    setBusy(true);
    setError("");
    try {
      const report = await runManagementDeckReview(settings);
      setReviewReport(report);
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setBusy(false);
    }
  };

  const handleBuildBackupQna = async () => {
    if (!reviewReport) return;
    setBusy(true);
    setError("");
    try {
      const { createdCount, undoSnapshots } = await buildBackupSlidesFromQna(
        reviewReport.toughQuestions,
        settings
      );
      onCompleted(
        `Built ${createdCount} Backup Q&A slide(s) at the end of the presentation.`,
        undoSnapshots
      );
      onClose();
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setBusy(false);
    }
  };

  const handleApplyBeautifier = async () => {
    setBusy(true);
    setError("");
    try {
      const targetIdx =
        beautifySlideIdx > 0 ? beautifySlideIdx : await getActiveSlideIndex();
      const { summary, undoSnapshot } = await beautifySlideLayout({
        slideNumber: targetIdx,
        style: beautifyStyle,
        settings,
      });
      onCompleted(summary, [undoSnapshot]);
      onClose();
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setBusy(false);
    }
  };

  const modalTitle =
    mode === "report_to_deck"
      ? "📄 Report / PDF → Slide Deck (Outline & Approval)"
      : mode === "consistency"
        ? "✅ Presentation Consistency Checker"
        : mode === "text_fit"
          ? "📐 Slide Text-Fit Fixer"
          : mode === "deck_review"
            ? "🎯 Management Q&A & Deck Reviewer"
            : "🎨 Smart Slide Layout Beautifier";

  return (
    <Dialog open={Boolean(mode)} onOpenChange={(_, d) => !d.open && onClose()}>
      <DialogSurface style={{ maxWidth: 460, width: "95vw", maxHeight: "90vh" }}>
        <DialogBody>
          <DialogTitle
            action={
              <Button
                appearance="subtle"
                icon={<Dismiss20Regular />}
                onClick={onClose}
                aria-label="Close"
              />
            }
          >
            {modalTitle}
          </DialogTitle>

          <DialogContent
            style={{
              display: "flex",
              flexDirection: "column",
              gap: 10,
              overflowY: "auto",
              maxHeight: "70vh",
              fontSize: 12,
            }}
          >
            {error && (
              <div
                style={{
                  padding: 8,
                  borderRadius: 6,
                  background: "rgba(209, 52, 56, 0.12)",
                  color: "#d13438",
                }}
              >
                ⚠️ {error}
              </div>
            )}

            {/* MODE 1: REPORT / PDF -> DECK */}
            {mode === "report_to_deck" && (
              <>
                {!outlinePlan ? (
                  <>
                    <div>
                      <Label size="small">
                        Step 1: Paste a Word/Engineering Report or attach a PDF
                      </Label>
                      <Textarea
                        rows={6}
                        style={{ width: "100%", marginTop: 4 }}
                        placeholder="Paste report sections, executive summary, well surveillance notes, or attach a PDF below..."
                        value={reportText}
                        onChange={(_, d) => setReportText(d.value)}
                      />
                    </div>

                    <div
                      style={{
                        display: "flex",
                        alignItems: "center",
                        justifyContent: "space-between",
                        gap: 8,
                        flexWrap: "wrap",
                      }}
                    >
                      <div>
                        <input
                          type="file"
                          id="zee-report-file"
                          accept="application/pdf,.pdf,.txt,.md,image/*"
                          style={{ display: "none" }}
                          onChange={async (e) => {
                            if (e.target.files?.length) {
                              const att = await fileToChatAttachment(e.target.files[0]);
                              setAttachments([att]);
                            }
                          }}
                        />
                        <Button
                          size="small"
                          appearance="outline"
                          onClick={() =>
                            document.getElementById("zee-report-file")?.click()
                          }
                        >
                          📎 Attach PDF / Report File
                        </Button>
                        {attachments.map((a) => (
                          <Badge
                            key={a.id}
                            size="small"
                            appearance="tint"
                            style={{ marginLeft: 6 }}
                          >
                            {a.name}
                          </Badge>
                        ))}
                      </div>

                      <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                        <Label size="small">Target Slides:</Label>
                        <Input
                          size="small"
                          type="number"
                          style={{ width: 64 }}
                          value={String(targetSlideCount)}
                          onChange={(_, d) =>
                            setTargetSlideCount(
                              Math.max(3, Math.min(15, Number(d.value) || 6))
                            )
                          }
                        />
                      </div>
                    </div>
                  </>
                ) : (
                  <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                    <div
                      style={{
                        padding: 8,
                        borderRadius: 6,
                        background: "rgba(183, 54, 24, 0.08)",
                        fontWeight: 600,
                      }}
                    >
                      Step 2: Review &amp; Edit Slide Outline ({outlinePlan.slides.length} slides)
                    </div>

                    {outlinePlan.slides.map((sl, idx) => (
                      <div
                        key={idx}
                        style={{
                          border: "1px solid var(--colorNeutralStroke2, #e1dfdd)",
                          borderRadius: 6,
                          padding: 8,
                          display: "flex",
                          flexDirection: "column",
                          gap: 4,
                        }}
                      >
                        <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                          <Badge size="small" color="brand">
                            #{sl.slideNumber}
                          </Badge>
                          <Input
                            size="small"
                            style={{ flex: 1, fontWeight: 600 }}
                            value={sl.title}
                            onChange={(_, d) => {
                              const next = [...outlinePlan.slides];
                              next[idx] = { ...next[idx], title: d.value };
                              setOutlinePlan({ ...outlinePlan, slides: next });
                            }}
                          />
                        </div>
                        <Textarea
                          size="small"
                          rows={2}
                          value={sl.bullets.join("\n")}
                          onChange={(_, d) => {
                            const next = [...outlinePlan.slides];
                            next[idx] = {
                              ...next[idx],
                              bullets: d.value.split(/\r?\n/).filter(Boolean),
                            };
                            setOutlinePlan({ ...outlinePlan, slides: next });
                          }}
                        />
                      </div>
                    ))}
                  </div>
                )}
              </>
            )}

            {/* MODE 2: CONSISTENCY CHECKER */}
            {mode === "consistency" && (
              <>
                <div style={{ opacity: 0.85 }}>
                  Checks all slides for font family mismatches, font sizes, title punctuation,
                  engineering units (<code>bbl/d</code>, <code>scf/stb</code>), and protected
                  terms (<code>ESP</code>, <code>TDH</code>, <code>PIP</code>, <code>GOR</code>,{" "}
                  <code>BHP</code>).
                </div>

                {!consistencyScanned ? (
                  <Button
                    appearance="primary"
                    disabled={busy}
                    onClick={handleRunConsistencyScan}
                    style={{ backgroundColor: "#b73618", alignSelf: "flex-start" }}
                  >
                    🔍 Scan All Slides Now
                  </Button>
                ) : consistencyIssues.length === 0 ? (
                  <div style={{ padding: 12, color: "#107c10", fontWeight: 600 }}>
                    ✓ All slides are consistent! No font, unit, or protected-term issues found.
                  </div>
                ) : (
                  <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                    {consistencyIssues.map((iss, idx) => (
                      <div
                        key={iss.id}
                        style={{
                          border: "1px solid var(--colorNeutralStroke2, #e1dfdd)",
                          borderRadius: 6,
                          padding: "6px 8px",
                          display: "flex",
                          alignItems: "flex-start",
                          gap: 6,
                        }}
                      >
                        <Checkbox
                          checked={iss.selected}
                          onChange={(_, d) => {
                            const next = [...consistencyIssues];
                            next[idx] = { ...next[idx], selected: Boolean(d.checked) };
                            setConsistencyIssues(next);
                          }}
                        />
                        <div style={{ flex: 1 }}>
                          <div style={{ fontWeight: 600 }}>{iss.description}</div>
                          <div style={{ fontSize: 11, opacity: 0.8 }}>
                            Current: <code>{iss.currentValue}</code> → Suggested:{" "}
                            <code>{iss.suggestedValue}</code>
                          </div>
                        </div>
                      </div>
                    ))}
                  </div>
                )}
              </>
            )}

            {/* MODE 3: TEXT-FIT FIXER */}
            {mode === "text_fit" && (
              <>
                <div style={{ opacity: 0.85 }}>
                  Finds placeholders where bullet points overflow the text box and shortens/resizes
                  them while preserving all engineering numbers and protected terms.
                </div>

                {!textFitScanned ? (
                  <Button
                    appearance="primary"
                    disabled={busy}
                    onClick={handleRunTextFitScan}
                    style={{ backgroundColor: "#b73618", alignSelf: "flex-start" }}
                  >
                    📐 Scan Slides for Overflow
                  </Button>
                ) : textFitIssues.length === 0 ? (
                  <div style={{ padding: 12, color: "#107c10", fontWeight: 600 }}>
                    ✓ All slide placeholders fit cleanly! No overflowing text boxes detected.
                  </div>
                ) : (
                  <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                    {textFitIssues.map((iss, idx) => (
                      <div
                        key={iss.id}
                        style={{
                          border: "1px solid var(--colorNeutralStroke2, #e1dfdd)",
                          borderRadius: 6,
                          padding: "6px 8px",
                          display: "flex",
                          alignItems: "flex-start",
                          gap: 6,
                        }}
                      >
                        <Checkbox
                          checked={iss.selected}
                          onChange={(_, d) => {
                            const next = [...textFitIssues];
                            next[idx] = { ...next[idx], selected: Boolean(d.checked) };
                            setTextFitIssues(next);
                          }}
                        />
                        <div style={{ flex: 1 }}>
                          <div style={{ fontWeight: 600 }}>
                            Slide {iss.slideIndex} ({iss.shapeName}) — {iss.wordCount} words (
                            {Math.round(iss.estimatedOverflowRatio * 100)}% capacity)
                          </div>
                          <div style={{ fontSize: 11, opacity: 0.8 }}>
                            {iss.originalText.slice(0, 140)}...
                          </div>
                        </div>
                      </div>
                    ))}
                  </div>
                )}
              </>
            )}

            {/* MODE 4: MANAGEMENT Q&A & DECK REVIEWER */}
            {mode === "deck_review" && (
              <>
                <div style={{ opacity: 0.85 }}>
                  Audits the entire presentation for cross-slide numerical contradictions (same well
                  or metric with conflicting numbers), scores executive readiness, and generates
                  tough Management Q&amp;A with one-click Backup Slides.
                </div>

                {!reviewReport ? (
                  <Button
                    appearance="primary"
                    disabled={busy}
                    onClick={handleRunDeckReview}
                    style={{ backgroundColor: "#b73618", alignSelf: "flex-start" }}
                  >
                    🎯 Run Management &amp; Technical Audit
                  </Button>
                ) : (
                  <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                    <div
                      style={{
                        padding: 8,
                        borderRadius: 6,
                        background: "rgba(183, 54, 24, 0.08)",
                        display: "flex",
                        alignItems: "center",
                        justifyContent: "space-between",
                        gap: 8,
                      }}
                    >
                      <span style={{ fontWeight: 700 }}>
                        Executive Readiness Score: {reviewReport.overallScore}/100
                      </span>
                      <Badge
                        color={reviewReport.overallScore >= 80 ? "success" : "warning"}
                        size="small"
                      >
                        {reviewReport.contradictions.length} Contradiction(s)
                      </Badge>
                    </div>

                    <div style={{ fontSize: 11, opacity: 0.9 }}>
                      {reviewReport.executiveSummary}
                    </div>

                    {reviewReport.contradictions.length > 0 && (
                      <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
                        <Label size="small" style={{ color: "#d13438", fontWeight: 700 }}>
                          ⚠️ Cross-Slide Numerical Contradictions Found:
                        </Label>
                        {reviewReport.contradictions.map((c) => (
                          <div
                            key={c.id}
                            style={{
                              padding: "6px 8px",
                              borderRadius: 6,
                              background: "rgba(209, 52, 56, 0.08)",
                              border: "1px solid rgba(209, 52, 56, 0.25)",
                              fontSize: 11,
                            }}
                          >
                            {c.description}
                          </div>
                        ))}
                      </div>
                    )}

                    <Label size="small" style={{ fontWeight: 700, marginTop: 4 }}>
                       Anticipated Management Q&amp;A (Select to build Backup Slides):
                    </Label>
                    {reviewReport.toughQuestions.map((q, idx) => (
                      <div
                        key={q.id}
                        style={{
                          border: "1px solid var(--colorNeutralStroke2, #e1dfdd)",
                          borderRadius: 6,
                          padding: "6px 8px",
                          display: "flex",
                          alignItems: "flex-start",
                          gap: 6,
                        }}
                      >
                        <Checkbox
                          checked={q.selected}
                          onChange={(_, d) => {
                            const next = [...reviewReport.toughQuestions];
                            next[idx] = { ...next[idx], selected: Boolean(d.checked) };
                            setReviewReport({
                              ...reviewReport,
                              toughQuestions: next,
                            });
                          }}
                        />
                        <div style={{ flex: 1 }}>
                          <div style={{ fontWeight: 600 }}>
                            <Badge size="small" appearance="outline" style={{ marginRight: 6 }}>
                              {q.backupSlideTitle}
                            </Badge>
                            {q.question}
                          </div>
                          <div style={{ fontSize: 11, opacity: 0.85, marginTop: 2 }}>
                            <b>Answer:</b> {q.recommendedAnswer}
                          </div>
                        </div>
                      </div>
                    ))}
                  </div>
                )}
              </>
            )}

            {/* MODE 5: SMART SLIDE LAYOUT BEAUTIFIER */}
            {mode === "beautify" && (
              <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                <div style={{ opacity: 0.85 }}>
                  Transforms a dense bullet slide into a clean 2-Column comparison split, a
                  3-Card executive summary grid, or aligned company-template typography (with
                  one-click Undo).
                </div>

                <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
                  <div style={{ flex: 1 }}>
                    <Label size="small">Target Slide Number (1-based)</Label>
                    <Input
                      size="small"
                      type="number"
                      style={{ width: "100%", marginTop: 4 }}
                      value={String(beautifySlideIdx)}
                      onChange={(_, d) =>
                        setBeautifySlideIdx(Math.max(1, Number(d.value) || 1))
                      }
                    />
                  </div>
                  <div style={{ flex: 2 }}>
                    <Label size="small">Beautifier Layout Style</Label>
                    <Select
                      size="small"
                      style={{ width: "100%", marginTop: 4 }}
                      value={beautifyStyle}
                      onChange={(_, d) =>
                        setBeautifyStyle(d.value as SlideBeautifyStyle)
                      }
                    >
                      <option value="two_column">
                        2-Column Split (Comparison / Before vs After)
                      </option>
                      <option value="three_cards">
                        3-Card Callout Grid (Key Takeaways / Pillars)
                      </option>
                      <option value="align_grid">
                        Snap &amp; Align Typography to Company Template
                      </option>
                    </Select>
                  </div>
                </div>
              </div>
            )}

            {busy && (
              <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                <Spinner size="tiny" />
                <span>{buildProgress || "Working..."}</span>
              </div>
            )}
          </DialogContent>

          <DialogActions>
            {mode === "report_to_deck" && !outlinePlan && (
              <Button
                appearance="primary"
                disabled={busy}
                onClick={handleGenerateOutline}
                style={{ backgroundColor: "#b73618" }}
              >
                1. Generate Slide Outline
              </Button>
            )}

            {mode === "report_to_deck" && outlinePlan && (
              <>
                <Button
                  appearance="outline"
                  disabled={busy}
                  onClick={() => setOutlinePlan(null)}
                >
                  ← Edit Source
                </Button>
                <Button
                  appearance="primary"
                  disabled={busy}
                  onClick={handleApproveAndBuildDeck}
                  style={{ backgroundColor: "#b73618" }}
                >
                  ✓ Approve &amp; Build Slides ({outlinePlan.slides.length})
                </Button>
              </>
            )}

            {mode === "consistency" &&
              consistencyScanned &&
              consistencyIssues.some((i) => i.selected) && (
                <Button
                  appearance="primary"
                  disabled={busy}
                  onClick={handleFixConsistency}
                  style={{ backgroundColor: "#b73618" }}
                >
                  ✓ Fix Selected ({consistencyIssues.filter((i) => i.selected).length})
                </Button>
              )}

            {mode === "text_fit" &&
              textFitScanned &&
              textFitIssues.some((i) => i.selected) && (
                <Button
                  appearance="primary"
                  disabled={busy}
                  onClick={handleFixTextFit}
                  style={{ backgroundColor: "#b73618" }}
                >
                  ✓ Shorten &amp; Fit Selected ({textFitIssues.filter((i) => i.selected).length})
                </Button>
              )}

            {mode === "deck_review" &&
              reviewReport &&
              reviewReport.toughQuestions.some((q) => q.selected) && (
                <Button
                  appearance="primary"
                  disabled={busy}
                  onClick={handleBuildBackupQna}
                  style={{ backgroundColor: "#b73618" }}
                >
                  ✓ Build Selected Backup Q&amp;A Slides (
                  {reviewReport.toughQuestions.filter((q) => q.selected).length})
                </Button>
              )}

            {mode === "beautify" && (
              <Button
                appearance="primary"
                disabled={busy}
                onClick={handleApplyBeautifier}
                style={{ backgroundColor: "#b73618" }}
              >
                🎨 Apply Layout Beautifier
              </Button>
            )}

            <Button appearance="subtle" onClick={onClose}>
              Cancel
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
};
