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
  Spinner,
  Textarea,
} from "@fluentui/react-components";
import { Dismiss20Regular } from "@fluentui/react-icons";
import { ChatAttachment, SlideUndoSnapshot } from "../types/chat";
import {
  ConsistencyIssue,
  DeckOutlinePlan,
  TextFitIssue,
} from "../types/ppt";
import { ZeeSettings } from "../types/settings";
import { fileToChatAttachment } from "./ChatInput";
import {
  applyConsistencyFixes,
  buildSlidesFromApprovedOutline,
  fixTextFitIssues,
  generateDeckOutlineFromReport,
  runPresentationConsistencyCheck,
  scanPresentationForTextFitIssues,
} from "../services/pptFeaturesService";

export type DeckToolMode = "report_to_deck" | "consistency" | "text_fit" | null;

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

  const modalTitle =
    mode === "report_to_deck"
      ? "📄 Report / PDF → Slide Deck (Outline & Approval)"
      : mode === "consistency"
        ? "✅ Presentation Consistency Checker"
        : "📐 Slide Text-Fit Fixer";

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

            <Button appearance="subtle" onClick={onClose}>
              Cancel
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
};
