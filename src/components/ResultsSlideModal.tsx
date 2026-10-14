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
  Switch,
  Textarea,
} from "@fluentui/react-components";
import { Dismiss20Regular } from "@fluentui/react-icons";
import { SlideUndoSnapshot } from "../types/chat";
import { ExcelResultsPayload, PetroleumResultCategory } from "../types/ppt";
import { ZeeSettings } from "../types/settings";
import { convertPetroleumUnit } from "../services/arabicService";
import {
  buildBatchWellBookSlides,
  buildExcelResultsSlide,
  parseMultiWellTable,
  prepareExcelResultsSlidePayload,
} from "../services/pptFeaturesService";
import { fileToChatAttachment } from "./ChatInput";

interface ResultsSlideModalProps {
  open: boolean;
  settings: ZeeSettings;
  onClose: () => void;
  onCompleted: (summaryMessage: string, undoSnapshots: SlideUndoSnapshot[]) => void;
}

const SAMPLE_TEMPLATES: Record<PetroleumResultCategory, string> = {
  decline_curve: `Parameter\tWELL-04\tWELL-12
Initial Rate Qi (stb/d)\t1,450\t980
Decline Rate Di (%/yr)\t14.2%\t11.8%
Arps b-factor\t0.45\t0.38
EUR (MMstb)\t2.85\t1.92
Remaining Reserves (Mstb)\t740\t510`,
  esp_analysis: `Parameter\tMeasured\tDesign Target
Pump Intake Pressure PIP (psi)\t1,120\t1,080
Total Dynamic Head TDH (ft)\t5,400\t5,350
Bottomhole Pressure BHP (psia)\t1,640\t1,600
Gas-Oil Ratio GOR (scf/stb)\t420\t400
Liquid Rate (bbl/d)\t2,450\t2,500
Operating Frequency (Hz)\t56\t58`,
  production_qc: `Well\tLiquid (bbl/d)\tOil (stb/d)\tWC (%)\tGOR (scf/stb)\tBHP (psia)\tQC Status
WELL-01\t1,820\t1,310\t28.0%\t390\t1,720\tValid
WELL-04\t2,450\t1,690\t31.0%\t420\t1,580\tValid
WELL-09\t940\t410\t56.4%\t890\t1,190\tHigh GOR / WC Flag`,
  custom: `Metric\tCurrent\tTarget
Field Oil Rate (stb/d)\t8,450\t9,000
Average BHP (psia)\t1,620\t1,650`,
};

const SAMPLE_WELLBOOK_TABLE = `Well\tLiquid (bbl/d)\tOil (stb/d)\tWC (%)\tGOR (scf/stb)\tPIP (psi)\tTDH (ft)\tBHP (psia)\tStatus\tRecommended Action
WELL-01\t1,820\t1,310\t28.0%\t390\t1,180\t5,150\t1,720\tValid\tMaintain 55 Hz ESP speed
WELL-04\t2,450\t1,690\t31.0%\t420\t1,120\t5,400\t1,580\tValid\tStep up VSD +2 Hz for +140 stb/d
WELL-09\t940\t410\t56.4%\t890\t860\t5,620\t1,190\tHigh GOR Flag\tCheck gas separator & choke back
WELL-12\t1,540\t1,140\t26.0%\t360\t1,240\t4,980\t1,810\tValid\tCandidate for drawdown optimization`;

export const ResultsSlideModal: React.FC<ResultsSlideModalProps> = ({
  open,
  settings,
  onClose,
  onCompleted,
}) => {
  const [tabMode, setTabMode] = useState<"single_slide" | "batch_wellbook">("single_slide");

  // Single Slide states
  const [category, setCategory] = useState<PetroleumResultCategory>("esp_analysis");
  const [wellOrFieldName, setWellOrFieldName] = useState("WELL-04 / North Field");
  const [rawInputText, setRawInputText] = useState(SAMPLE_TEMPLATES.esp_analysis);
  const [autoPlotChart, setAutoPlotChart] = useState(true);
  const [chartBase64, setChartBase64] = useState<string | undefined>(undefined);
  const [chartFileName, setChartFileName] = useState<string>("");
  const [previewPayload, setPreviewPayload] = useState<ExcelResultsPayload | null>(null);

  // Batch Well-Book states
  const [batchFieldName, setBatchFieldName] = useState("North Field Q3 Well-Book");
  const [batchTableText, setBatchTableText] = useState(SAMPLE_WELLBOOK_TABLE);
  const [includeRollupSlide, setIncludeRollupSlide] = useState(true);
  const [includeBatchCharts, setIncludeBatchCharts] = useState(true);
  const [batchProgress, setBatchProgress] = useState("");

  // Built-in unit converter quick helper state
  const [unitValue, setUnitValue] = useState("2450");
  const [unitFrom, setUnitFrom] = useState("bbl/d");

  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");

  if (!open) return null;

  const convertedUnit = convertPetroleumUnit(Number(unitValue), unitFrom);
  const parsedBatch = parseMultiWellTable(batchTableText);

  const handlePrepareSlide = async () => {
    if (!rawInputText.trim()) {
      setError("Please paste your Excel results table or metrics first.");
      return;
    }
    setBusy(true);
    setError("");
    try {
      const payload = await prepareExcelResultsSlidePayload({
        category,
        wellOrFieldName,
        rawInputText,
        chartImageBase64: chartBase64,
        autoGenerateChart: autoPlotChart,
        settings,
      });
      setPreviewPayload(payload);
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setBusy(false);
    }
  };

  const handleInsertResultsSlide = async () => {
    if (!previewPayload) return;
    setBusy(true);
    setError("");
    try {
      const res = await buildExcelResultsSlide(previewPayload, settings);
      onCompleted(
        `Created Results Slide #${res.slideIndex}: "${previewPayload.slideTitle}" (with ${
          previewPayload.chartImageBase64 ? "Auto-Plotted Chart & Table" : "Summary Table"
        }).`,
        [res.undoSnapshot]
      );
      onClose();
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setBusy(false);
    }
  };

  const handleBuildWellBook = async () => {
    if (parsedBatch.wells.length === 0) {
      setError("Please paste a multi-well table with at least 1 well row.");
      return;
    }
    setBusy(true);
    setError("");
    try {
      const { createdCount, undoSnapshots } = await buildBatchWellBookSlides({
        fieldName: batchFieldName,
        rawTableText: batchTableText,
        includeRollupSlide,
        includeAutoCharts: includeBatchCharts,
        settings,
        onProgress: (curr, total, label) => {
          setBatchProgress(`Building slide ${curr}/${total}: ${label}...`);
        },
      });
      setBatchProgress("");
      onCompleted(
        `Batch Well-Book Generator built ${createdCount} slides (${parsedBatch.wells.length} well slides${
          includeRollupSlide ? " + 1 Field Roll-Up slide" : ""
        }).`,
        undoSnapshots
      );
      onClose();
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setBusy(false);
    }
  };

  return (
    <Dialog open={open} onOpenChange={(_, d) => !d.open && onClose()}>
      <DialogSurface style={{ maxWidth: 470, width: "95vw", maxHeight: "90vh" }}>
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
            📊 Excel Results &amp; Batch Well-Book Generator
          </DialogTitle>

          <DialogContent
            style={{
              display: "flex",
              flexDirection: "column",
              gap: 10,
              overflowY: "auto",
              maxHeight: "72vh",
              fontSize: 12,
            }}
          >
            {/* Mode Switcher */}
            <div style={{ display: "flex", gap: 6 }}>
              <Button
                size="small"
                appearance={tabMode === "single_slide" ? "primary" : "outline"}
                onClick={() => {
                  setTabMode("single_slide");
                  setError("");
                }}
                style={{
                  flex: 1,
                  backgroundColor: tabMode === "single_slide" ? "#b73618" : undefined,
                }}
              >
                📈 Single Slide + Auto-Chart
              </Button>
              <Button
                size="small"
                appearance={tabMode === "batch_wellbook" ? "primary" : "outline"}
                onClick={() => {
                  setTabMode("batch_wellbook");
                  setError("");
                }}
                style={{
                  flex: 1,
                  backgroundColor: tabMode === "batch_wellbook" ? "#b73618" : undefined,
                }}
              >
                📚 Batch Well-Book ({parsedBatch.wells.length} Wells)
              </Button>
            </div>

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

            {/* TAB 1: SINGLE RESULTS SLIDE + AUTO-PLOT CHART */}
            {tabMode === "single_slide" && (
              <>
                {!previewPayload ? (
                  <>
                    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                      <div>
                        <Label size="small">Analysis Type (Zee AI Excel)</Label>
                        <Select
                          size="small"
                          value={category}
                          onChange={(_, d) => {
                            const cat = d.value as PetroleumResultCategory;
                            setCategory(cat);
                            setRawInputText(SAMPLE_TEMPLATES[cat]);
                          }}
                        >
                          <option value="esp_analysis">ESP Analysis &amp; Pump Curve</option>
                          <option value="decline_curve">Decline Curve (Arps Plot)</option>
                          <option value="production_qc">Production QC &amp; KPI Strip</option>
                          <option value="custom">Custom Excel Table</option>
                        </Select>
                      </div>

                      <div>
                        <Label size="small">Well / Field Identifier</Label>
                        <Input
                          size="small"
                          style={{ width: "100%" }}
                          value={wellOrFieldName}
                          onChange={(_, d) => setWellOrFieldName(d.value)}
                        />
                      </div>
                    </div>

                    <div>
                      <Label size="small">
                        Paste Excel Table / Results (Tab, CSV, or Markdown)
                      </Label>
                      <Textarea
                        rows={5}
                        style={{
                          width: "100%",
                          fontFamily: "Consolas, monospace",
                          fontSize: 11,
                        }}
                        value={rawInputText}
                        onChange={(_, d) => setRawInputText(d.value)}
                      />
                    </div>

                    <div
                      style={{
                        display: "flex",
                        alignItems: "center",
                        justifyContent: "space-between",
                        flexWrap: "wrap",
                        gap: 6,
                      }}
                    >
                      <Switch
                        checked={autoPlotChart}
                        onChange={(_, d) => setAutoPlotChart(d.checked)}
                        label="Auto-Plot Vector Chart / KPI Cards onto Slide"
                      />

                      <div>
                        <input
                          type="file"
                          id="zee-excel-chart"
                          accept="image/*"
                          style={{ display: "none" }}
                          onChange={async (e) => {
                            if (e.target.files?.length) {
                              const att = await fileToChatAttachment(e.target.files[0]);
                              setChartBase64(att.base64Data);
                              setChartFileName(att.name);
                            }
                          }}
                        />
                        <Button
                          size="small"
                          appearance="outline"
                          onClick={() =>
                            document.getElementById("zee-excel-chart")?.click()
                          }
                        >
                          🖼️ Custom Image
                        </Button>
                        {chartFileName && (
                          <Badge size="small" appearance="tint" color="brand">
                            {chartFileName}
                          </Badge>
                        )}
                      </div>
                    </div>

                    {/* Built-in Petroleum Unit Converter Helper */}
                    <div
                      style={{
                        border: "1px solid var(--colorNeutralStroke2, #e1dfdd)",
                        borderRadius: 6,
                        padding: 8,
                        display: "flex",
                        flexDirection: "column",
                        gap: 6,
                      }}
                    >
                      <div style={{ fontWeight: 600, fontSize: 11 }}>
                        🛢️ Quick Petroleum Unit Converter
                      </div>
                      <div
                        style={{
                          display: "flex",
                          alignItems: "center",
                          gap: 6,
                          flexWrap: "wrap",
                        }}
                      >
                        <Input
                          size="small"
                          style={{ width: 90 }}
                          value={unitValue}
                          onChange={(_, d) => setUnitValue(d.value)}
                        />
                        <Select
                          size="small"
                          value={unitFrom}
                          onChange={(_, d) => setUnitFrom(d.value)}
                        >
                          <option value="bbl/d">bbl/d → m³/d</option>
                          <option value="stb/d">stb/d → sm³/d</option>
                          <option value="psi">psi → bar</option>
                          <option value="ft">ft → m</option>
                          <option value="scf/stb">scf/stb → sm³/sm³</option>
                          <option value="degF">°F → °C</option>
                        </Select>
                        {convertedUnit && (
                          <Badge size="small" appearance="filled" color="informative">
                            = {convertedUnit.outputValue.toLocaleString("en-US")}{" "}
                            {convertedUnit.toUnit}
                          </Badge>
                        )}
                      </div>
                    </div>
                  </>
                ) : (
                  <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                    <div style={{ fontWeight: 600 }}>
                      Preview Results Slide: {previewPayload.slideTitle}
                    </div>
                    {previewPayload.kpis.length > 0 && (
                      <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
                        {previewPayload.kpis.map((k, idx) => (
                          <Badge key={idx} size="medium" appearance="tint" color="brand">
                            {k.label}: {k.value}
                          </Badge>
                        ))}
                      </div>
                    )}
                    {previewPayload.chartImageBase64 && (
                      <div
                        style={{
                          border: "1px solid var(--colorNeutralStroke2, #e1dfdd)",
                          borderRadius: 6,
                          padding: 6,
                          textAlign: "center",
                          background: "#fff",
                        }}
                      >
                        <div style={{ fontSize: 11, fontWeight: 600, marginBottom: 4, color: "#242424" }}>
                          ✓ Auto-Plotted Engineering Chart Ready to Insert
                        </div>
                        <img
                          src={`data:image/png;base64,${previewPayload.chartImageBase64}`}
                          alt="Auto-Plotted Chart"
                          style={{ maxWidth: "100%", maxHeight: 165, borderRadius: 4 }}
                        />
                      </div>
                    )}
                    <div
                      style={{
                        border: "1px solid var(--colorNeutralStroke2, #e1dfdd)",
                        borderRadius: 6,
                        padding: 8,
                      }}
                    >
                      <div style={{ fontWeight: 600, marginBottom: 4 }}>Key Takeaways:</div>
                      {previewPayload.takeaways.map((t, idx) => (
                        <div key={idx}>• {t}</div>
                      ))}
                    </div>
                  </div>
                )}
              </>
            )}

            {/* TAB 2: BATCH WELL-BOOK GENERATOR */}
            {tabMode === "batch_wellbook" && (
              <>
                <div style={{ opacity: 0.85 }}>
                  Paste a multi-well Excel table (1 row per well) to automatically generate a{" "}
                  <b>Field Roll-Up Summary Slide</b> plus <b>1 standardized slide per well</b>{" "}
                  with tables, notes, and auto-plotted charts!
                </div>

                <div>
                  <Label size="small">Field / Well-Book Title</Label>
                  <Input
                    size="small"
                    style={{ width: "100%" }}
                    value={batchFieldName}
                    onChange={(_, d) => setBatchFieldName(d.value)}
                  />
                </div>

                <div>
                  <Label size="small">
                    Multi-Well Table (First column = Well Name, e.g. WELL-01, WELL-04)
                  </Label>
                  <Textarea
                    rows={6}
                    style={{
                      width: "100%",
                      fontFamily: "Consolas, monospace",
                      fontSize: 11,
                    }}
                    value={batchTableText}
                    onChange={(_, d) => setBatchTableText(d.value)}
                  />
                </div>

                <div style={{ display: "flex", flexWrap: "wrap", gap: 10 }}>
                  <Checkbox
                    checked={includeRollupSlide}
                    onChange={(_, d) => setIncludeRollupSlide(Boolean(d.checked))}
                    label="Include Executive Field Roll-Up Slide first"
                  />
                  <Checkbox
                    checked={includeBatchCharts}
                    onChange={(_, d) => setIncludeBatchCharts(Boolean(d.checked))}
                    label="Auto-plot ESP / KPI chart on every well slide"
                  />
                </div>

                <div
                  style={{
                    padding: "6px 8px",
                    borderRadius: 6,
                    background: "rgba(183, 54, 24, 0.08)",
                    fontSize: 11,
                  }}
                >
                  Detected <b>{parsedBatch.wells.length}</b> well(s):{" "}
                  {parsedBatch.wells.map((w) => w.wellName).join(", ") || "None"} → Will build{" "}
                  <b>{parsedBatch.wells.length + (includeRollupSlide ? 1 : 0)}</b> slides total.
                </div>
              </>
            )}

            {busy && (
              <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                <Spinner size="tiny" />
                <span>{batchProgress || "Generating slide & auto-plotting chart..."}</span>
              </div>
            )}
          </DialogContent>

          <DialogActions>
            {tabMode === "single_slide" && !previewPayload && (
              <Button
                appearance="primary"
                disabled={busy}
                onClick={handlePrepareSlide}
                style={{ backgroundColor: "#b73618" }}
              >
                1. Analyze &amp; Auto-Plot Preview
              </Button>
            )}

            {tabMode === "single_slide" && previewPayload && (
              <>
                <Button
                  appearance="outline"
                  disabled={busy}
                  onClick={() => setPreviewPayload(null)}
                >
                  ← Edit Data
                </Button>
                <Button
                  appearance="primary"
                  disabled={busy}
                  onClick={handleInsertResultsSlide}
                  style={{ backgroundColor: "#b73618" }}
                >
                  ✓ Insert Results Slide into Deck
                </Button>
              </>
            )}

            {tabMode === "batch_wellbook" && (
              <Button
                appearance="primary"
                disabled={busy || parsedBatch.wells.length === 0}
                onClick={handleBuildWellBook}
                style={{ backgroundColor: "#b73618" }}
              >
                📚 Build {parsedBatch.wells.length + (includeRollupSlide ? 1 : 0)} Well-Book Slides
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
