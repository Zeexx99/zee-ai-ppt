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
  Spinner,
  Textarea,
} from "@fluentui/react-components";
import { Dismiss20Regular } from "@fluentui/react-icons";
import { SlideUndoSnapshot } from "../types/chat";
import { ExcelResultsPayload, PetroleumResultCategory } from "../types/ppt";
import { ZeeSettings } from "../types/settings";
import { convertPetroleumUnit } from "../services/arabicService";
import {
  buildExcelResultsSlide,
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

export const ResultsSlideModal: React.FC<ResultsSlideModalProps> = ({
  open,
  settings,
  onClose,
  onCompleted,
}) => {
  const [category, setCategory] = useState<PetroleumResultCategory>("esp_analysis");
  const [wellOrFieldName, setWellOrFieldName] = useState("WELL-04 / North Field");
  const [rawInputText, setRawInputText] = useState(SAMPLE_TEMPLATES.esp_analysis);
  const [chartBase64, setChartBase64] = useState<string | undefined>(undefined);
  const [chartFileName, setChartFileName] = useState<string>("");
  const [previewPayload, setPreviewPayload] = useState<ExcelResultsPayload | null>(null);

  // Built-in unit converter quick helper state
  const [unitValue, setUnitValue] = useState("2450");
  const [unitFrom, setUnitFrom] = useState("bbl/d");

  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");

  if (!open) return null;

  const convertedUnit = convertPetroleumUnit(Number(unitValue), unitFrom);

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
        `Created Results Slide #${res.slideIndex}: "${previewPayload.slideTitle}".`,
        [res.undoSnapshot]
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
            📊 Excel Results → PowerPoint Results Slide
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

            {!previewPayload ? (
              <>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                  <div>
                    <Label size="small">Analysis Type (from Zee AI Excel)</Label>
                    <Select
                      size="small"
                      value={category}
                      onChange={(_, d) => {
                        const cat = d.value as PetroleumResultCategory;
                        setCategory(cat);
                        setRawInputText(SAMPLE_TEMPLATES[cat]);
                      }}
                    >
                      <option value="esp_analysis">ESP Analysis &amp; Diagnosis</option>
                      <option value="decline_curve">Decline Curve (Arps Forecast)</option>
                      <option value="production_qc">Production QC &amp; Well Test</option>
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
                    rows={6}
                    style={{ width: "100%", fontFamily: "Consolas, monospace", fontSize: 11 }}
                    value={rawInputText}
                    onChange={(_, d) => setRawInputText(d.value)}
                  />
                </div>

                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
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
                    onClick={() => document.getElementById("zee-excel-chart")?.click()}
                  >
                    🖼️ Attach Excel Chart Image (Optional)
                  </Button>
                  {chartFileName && (
                    <Badge size="small" appearance="tint" color="brand">
                      {chartFileName}
                    </Badge>
                  )}
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
                  <div style={{ display: "flex", alignItems: "center", gap: 6, flexWrap: "wrap" }}>
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
                        = {convertedUnit.outputValue.toLocaleString("en-US")} {convertedUnit.toUnit}
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

            {busy && (
              <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                <Spinner size="tiny" />
                <span>Building Results Slide...</span>
              </div>
            )}
          </DialogContent>

          <DialogActions>
            {!previewPayload ? (
              <Button
                appearance="primary"
                disabled={busy}
                onClick={handlePrepareSlide}
                style={{ backgroundColor: "#b73618" }}
              >
                1. Analyze &amp; Preview Slide
              </Button>
            ) : (
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
            <Button appearance="subtle" onClick={onClose}>
              Cancel
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
};
