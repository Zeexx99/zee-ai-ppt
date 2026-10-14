export interface DeclineCurveParams {
  wellName: string;
  qi: number; // initial oil rate (stb/d)
  diAnnualPercent: number; // nominal annual decline % (e.g. 14.2)
  bFactor: number; // Arps b-factor (0 = exponential, 0..1 = hyperbolic, 1 = harmonic)
  months?: number; // forecast horizon (default 36)
  econLimit?: number; // economic limit stb/d (default 50)
  primaryColor?: string;
}

export interface DeclineCurvePoint {
  month: number;
  rateStbPerDay: number;
  cumOilMstb: number;
}

export interface EspCurveParams {
  wellName: string;
  operatingRateBblPerDay: number;
  operatingTdhFt: number;
  pipPsi: number;
  frequencyHz?: number;
  designRateBblPerDay?: number;
  primaryColor?: string;
}

export interface KpiCardItem {
  label: string;
  value: string;
  status?: "normal" | "warning" | "critical";
}

function escapeXml(str: string): string {
  return String(str || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

/**
 * Computes monthly Arps decline curve points and cumulative production (MSTB).
 */
export function computeArpsDeclineCurve(params: DeclineCurveParams): {
  points: DeclineCurvePoint[];
  eurMstb: number;
} {
  const qi = Math.max(10, Number(params.qi) || 1000);
  const diMonthly = Math.max(0.001, (Number(params.diAnnualPercent) || 12) / 100 / 12);
  const b = Math.max(0, Math.min(1.2, Number(params.bFactor) ?? 0.4));
  const months = Math.max(6, Math.min(120, params.months || 36));
  const econLimit = Math.max(1, params.econLimit ?? 50);

  const points: DeclineCurvePoint[] = [];
  let cumBarrels = 0;

  for (let m = 0; m <= months; m++) {
    const rate =
      b < 0.001
        ? qi * Math.exp(-diMonthly * m)
        : qi / Math.pow(1 + b * diMonthly * m, 1 / b);

    if (m > 0) {
      const prevRate = points[m - 1].rateStbPerDay;
      cumBarrels += ((prevRate + rate) / 2) * 30.4375;
    }

    points.push({
      month: m,
      rateStbPerDay: Number(Math.max(econLimit * 0.5, rate).toFixed(1)),
      cumOilMstb: Number((cumBarrels / 1000).toFixed(1)),
    });
  }

  return {
    points,
    eurMstb: points[points.length - 1]?.cumOilMstb || 0,
  };
}

/**
 * Renders a crisp SVG chart for an Arps Decline Curve Forecast.
 */
export function renderDeclineCurveSvg(params: DeclineCurveParams): string {
  const { points, eurMstb } = computeArpsDeclineCurve(params);
  const width = 680;
  const height = 380;
  const padLeft = 68;
  const padRight = 32;
  const padTop = 76;
  const padBottom = 52;
  const plotW = width - padLeft - padRight;
  const plotH = height - padTop - padBottom;

  const maxRate = Math.max(...points.map((p) => p.rateStbPerDay), 100) * 1.1;
  const maxMonth = points[points.length - 1]?.month || 36;
  const econLimit = params.econLimit ?? 50;
  const brandColor = params.primaryColor || "#B73618";

  const toX = (m: number) => padLeft + (m / maxMonth) * plotW;
  const toY = (r: number) => padTop + plotH - Math.min(1, Math.max(0, r / maxRate)) * plotH;

  const polylinePoints = points
    .map((p) => `${toX(p.month).toFixed(1)},${toY(p.rateStbPerDay).toFixed(1)}`)
    .join(" ");

  const areaPoints = [
    `${toX(0).toFixed(1)},${(padTop + plotH).toFixed(1)}`,
    polylinePoints,
    `${toX(maxMonth).toFixed(1)},${(padTop + plotH).toFixed(1)}`,
  ].join(" ");

  const yTicks = [0, 0.25, 0.5, 0.75, 1].map((t) => Math.round(maxRate * t));
  const econY = toY(econLimit);

  return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">
  <rect width="${width}" height="${height}" rx="10" fill="#FFFFFF" stroke="#E1DFDD" stroke-width="1.5"/>
  <text x="24" y="32" font-family="Segoe UI, Arial, sans-serif" font-size="17" font-weight="700" fill="${escapeXml(brandColor)}">${escapeXml(params.wellName)} — Decline Curve Forecast (Arps)</text>
  <text x="24" y="54" font-family="Segoe UI, Arial, sans-serif" font-size="12" fill="#484644">Qi: ${params.qi.toLocaleString("en-US")} stb/d   |   Di: ${params.diAnnualPercent}%/yr   |   b: ${params.bFactor}   |   Forecast Cum: ${eurMstb.toLocaleString("en-US")} MSTB</text>
  ${yTicks
    .map((val) => {
      const y = toY(val).toFixed(1);
      return `<line x1="${padLeft}" y1="${y}" x2="${width - padRight}" y2="${y}" stroke="#EDEBE9" stroke-width="1"/>
  <text x="${padLeft - 8}" y="${Number(y) + 4}" text-anchor="end" font-family="Segoe UI, Arial, sans-serif" font-size="11" fill="#605E5C">${val.toLocaleString("en-US")}</text>`;
    })
    .join("\n  ")}
  <polygon points="${areaPoints}" fill="${escapeXml(brandColor)}" fill-opacity="0.12"/>
  <polyline fill="none" stroke="${escapeXml(brandColor)}" stroke-width="3" points="${polylinePoints}"/>
  <line x1="${padLeft}" y1="${econY.toFixed(1)}" x2="${width - padRight}" y2="${econY.toFixed(1)}" stroke="#D13438" stroke-width="1.5" stroke-dasharray="5,4"/>
  <text x="${width - padRight - 6}" y="${(econY - 6).toFixed(1)}" text-anchor="end" font-family="Segoe UI, Arial, sans-serif" font-size="10" font-weight="600" fill="#D13438">Econ Limit (${econLimit} stb/d)</text>
  <line x1="${padLeft}" y1="${padTop}" x2="${padLeft}" y2="${padTop + plotH}" stroke="#8A8886" stroke-width="1.5"/>
  <line x1="${padLeft}" y1="${padTop + plotH}" x2="${width - padRight}" y2="${padTop + plotH}" stroke="#8A8886" stroke-width="1.5"/>
  <text x="${padLeft + plotW / 2}" y="${height - 14}" text-anchor="middle" font-family="Segoe UI, Arial, sans-serif" font-size="12" font-weight="600" fill="#323130">Forecast Time (Months: 0 → ${maxMonth})</text>
  <text x="18" y="${padTop + plotH / 2}" text-anchor="middle" transform="rotate(-90 18 ${padTop + plotH / 2})" font-family="Segoe UI, Arial, sans-serif" font-size="12" font-weight="600" fill="#323130">Oil Rate (stb/d)</text>
</svg>`;
}

/**
 * Renders an ESP Total Dynamic Head (TDH) vs Liquid Rate performance curve with operating envelope.
 */
export function renderEspCurveSvg(params: EspCurveParams): string {
  const width = 680;
  const height = 380;
  const padLeft = 68;
  const padRight = 32;
  const padTop = 76;
  const padBottom = 52;
  const plotW = width - padLeft - padRight;
  const plotH = height - padTop - padBottom;

  const opRate = Math.max(100, Number(params.operatingRateBblPerDay) || 2450);
  const opTdh = Math.max(500, Number(params.operatingTdhFt) || 5400);
  const bepRate = params.designRateBblPerDay || opRate * 1.05;
  const shutInHead = opTdh * 1.32;
  const maxRateAxis = bepRate * 1.55;
  const maxHeadAxis = shutInHead * 1.12;
  const brandColor = params.primaryColor || "#0F6CBD";

  const toX = (q: number) => padLeft + Math.min(1, Math.max(0, q / maxRateAxis)) * plotW;
  const toY = (h: number) => padTop + plotH - Math.min(1, Math.max(0, h / maxHeadAxis)) * plotH;

  const curvePts: string[] = [];
  for (let i = 0; i <= 24; i++) {
    const q = (maxRateAxis * 0.92 * i) / 24;
    const ratio = q / bepRate;
    const head = Math.max(0, shutInHead - (shutInHead - opTdh) * Math.pow(ratio, 1.85));
    curvePts.push(`${toX(q).toFixed(1)},${toY(head).toFixed(1)}`);
  }

  const minOpX = toX(bepRate * 0.72);
  const maxOpX = toX(bepRate * 1.22);
  const opX = toX(opRate);
  const opY = toY(opTdh);

  return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">
  <rect width="${width}" height="${height}" rx="10" fill="#FFFFFF" stroke="#E1DFDD" stroke-width="1.5"/>
  <text x="24" y="32" font-family="Segoe UI, Arial, sans-serif" font-size="17" font-weight="700" fill="#B73618">${escapeXml(params.wellName)} — ESP Pump Performance Curve</text>
  <text x="24" y="54" font-family="Segoe UI, Arial, sans-serif" font-size="12" fill="#484644">Rate: ${opRate.toLocaleString("en-US")} bbl/d   |   TDH: ${opTdh.toLocaleString("en-US")} ft   |   PIP: ${params.pipPsi.toLocaleString("en-US")} psi   |   Freq: ${params.frequencyHz || 56} Hz</text>
  <rect x="${minOpX.toFixed(1)}" y="${padTop}" width="${(maxOpX - minOpX).toFixed(1)}" height="${plotH}" fill="#107C10" fill-opacity="0.08"/>
  <line x1="${minOpX.toFixed(1)}" y1="${padTop}" x2="${minOpX.toFixed(1)}" y2="${padTop + plotH}" stroke="#107C10" stroke-width="1.2" stroke-dasharray="4,4"/>
  <line x1="${maxOpX.toFixed(1)}" y1="${padTop}" x2="${maxOpX.toFixed(1)}" y2="${padTop + plotH}" stroke="#107C10" stroke-width="1.2" stroke-dasharray="4,4"/>
  <text x="${((minOpX + maxOpX) / 2).toFixed(1)}" y="${padTop + 16}" text-anchor="middle" font-family="Segoe UI, Arial, sans-serif" font-size="10" font-weight="600" fill="#107C10">Recommended Operating Range</text>
  <polyline fill="none" stroke="${escapeXml(brandColor)}" stroke-width="3" points="${curvePts.join(" ")}"/>
  <circle cx="${opX.toFixed(1)}" cy="${opY.toFixed(1)}" r="7" fill="#B73618" stroke="#FFFFFF" stroke-width="2"/>
  <text x="${(opX + 12).toFixed(1)}" y="${(opY - 8).toFixed(1)}" font-family="Segoe UI, Arial, sans-serif" font-size="11" font-weight="700" fill="#B73618">Operating Point (${opRate} bbl/d, ${opTdh} ft)</text>
  <line x1="${padLeft}" y1="${padTop}" x2="${padLeft}" y2="${padTop + plotH}" stroke="#8A8886" stroke-width="1.5"/>
  <line x1="${padLeft}" y1="${padTop + plotH}" x2="${width - padRight}" y2="${padTop + plotH}" stroke="#8A8886" stroke-width="1.5"/>
  <text x="${padLeft + plotW / 2}" y="${height - 14}" text-anchor="middle" font-family="Segoe UI, Arial, sans-serif" font-size="12" font-weight="600" fill="#323130">Liquid Rate (bbl/d)</text>
  <text x="18" y="${padTop + plotH / 2}" text-anchor="middle" transform="rotate(-90 18 ${padTop + plotH / 2})" font-family="Segoe UI, Arial, sans-serif" font-size="12" font-weight="600" fill="#323130">Total Dynamic Head — TDH (ft)</text>
</svg>`;
}

/**
 * Renders a row of executive KPI summary cards as a crisp SVG banner.
 */
export function renderKpiCardsSvg(
  kpis: KpiCardItem[],
  primaryColor: string = "#B73618"
): string {
  const items = kpis.slice(0, 4);
  const width = 680;
  const height = 110;
  const count = Math.max(1, items.length);
  const gap = 12;
  const cardWidth = Math.floor((width - 24 - gap * (count - 1)) / count);

  const statusAccent = (status?: KpiCardItem["status"]) => {
    if (status === "critical") return "#D13438";
    if (status === "warning") return "#D83B01";
    return primaryColor;
  };

  return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">
  <rect width="${width}" height="${height}" fill="none"/>
  ${items
    .map((item, idx) => {
      const x = 12 + idx * (cardWidth + gap);
      const accent = statusAccent(item.status);
      return `<g transform="translate(${x}, 8)">
    <rect width="${cardWidth}" height="92" rx="8" fill="#FAF9F8" stroke="#E1DFDD" stroke-width="1.5"/>
    <rect width="${cardWidth}" height="6" rx="3" fill="${escapeXml(accent)}"/>
    <text x="14" y="34" font-family="Segoe UI, Arial, sans-serif" font-size="11" font-weight="600" fill="#605E5C">${escapeXml(item.label.slice(0, 26))}</text>
    <text x="14" y="68" font-family="Segoe UI, Arial, sans-serif" font-size="20" font-weight="700" fill="#242424">${escapeXml(item.value.slice(0, 20))}</text>
  </g>`;
    })
    .join("\n  ")}
</svg>`;
}

/**
 * Rasterizes an SVG string to a PNG Base64 payload in browser/WebView2,
 * falling back to UTF-8 Base64 encoding in headless unit test environments.
 */
export async function svgToPngBase64(
  svgMarkup: string,
  width: number = 680,
  height: number = 380
): Promise<string> {
  if (
    typeof window === "undefined" ||
    typeof document === "undefined" ||
    typeof Image === "undefined"
  ) {
    return Buffer.from(svgMarkup, "utf-8").toString("base64");
  }

  return new Promise<string>((resolve) => {
    try {
      const svgBase64 = window.btoa(unescape(encodeURIComponent(svgMarkup)));
      const img = new Image();
      img.onload = () => {
        try {
          const canvas = document.createElement("canvas");
          canvas.width = width * 2; // 2x retina DPI for crisp PowerPoint slides
          canvas.height = height * 2;
          const ctx = canvas.getContext("2d");
          if (!ctx) {
            resolve(svgBase64);
            return;
          }
          ctx.scale(2, 2);
          ctx.fillStyle = "#FFFFFF";
          ctx.fillRect(0, 0, width, height);
          ctx.drawImage(img, 0, 0, width, height);
          const dataUrl = canvas.toDataURL("image/png");
          resolve(dataUrl.replace(/^data:image\/png;base64,/, ""));
        } catch {
          resolve(svgBase64);
        }
      };
      img.onerror = () => resolve(svgBase64);
      img.src = `data:image/svg+xml;base64,${svgBase64}`;
    } catch {
      resolve("");
    }
  });
}

export interface WellboreSchematicParams {
  wellName: string;
  surfaceCasingDepthFt: number;
  productionCasingDepthFt: number;
  tubingSizeInch?: string;
  espPumpDepthFt: number;
  packerDepthFt?: number;
  perforationTopFt: number;
  perforationBottomFt: number;
  reservoirZoneName?: string;
  pipPsi?: number;
  frequencyHz?: number;
  primaryColor?: string;
}

export interface ProcessFlowNode {
  id: string;
  label: string;
  subLabel?: string;
  metric?: string;
  status?: "normal" | "warning" | "critical";
}

export interface ProcessFlowchartParams {
  title: string;
  subtitle?: string;
  nodes: ProcessFlowNode[];
  primaryColor?: string;
}

export interface TimelineMilestone {
  phase: string;
  task: string;
  durationOrDate: string;
  owner?: string;
  status?: "completed" | "active" | "planned";
}

export interface TimelineDiagramParams {
  title: string;
  subtitle?: string;
  milestones: TimelineMilestone[];
  primaryColor?: string;
}

/**
 * Renders an engineering Wellbore & ESP Completion Schematic as a crisp SVG diagram.
 */
export function renderWellboreSchematicSvg(params: WellboreSchematicParams): string {
  const width = 680;
  const height = 400;
  const brand = params.primaryColor || "#B73618";
  const surfDepth = Math.max(500, params.surfaceCasingDepthFt || 1800);
  const prodDepth = Math.max(surfDepth + 1500, params.productionCasingDepthFt || 8500);
  const espDepth = Math.min(prodDepth - 500, Math.max(surfDepth + 300, params.espPumpDepthFt || 6800));
  const packerDepth = params.packerDepthFt || Math.min(prodDepth - 350, espDepth + 280);
  const perfTop = Math.min(prodDepth - 200, Math.max(packerDepth + 150, params.perforationTopFt || 7900));
  const perfBot = Math.min(prodDepth - 40, Math.max(perfTop + 80, params.perforationBottomFt || 8250));
  const tubing = params.tubingSizeInch || '3-1/2" EUE';
  const zoneName = params.reservoirZoneName || "Target Reservoir Sand";

  const topY = 74;
  const botY = 368;
  const depthToY = (d: number) =>
    topY + Math.min(1, Math.max(0, d / prodDepth)) * (botY - topY);

  const surfY = depthToY(surfDepth);
  const espY = depthToY(espDepth);
  const packerY = depthToY(packerDepth);
  const perfTopY = depthToY(perfTop);
  const perfBotY = depthToY(perfBot);
  const centerX = 250;

  return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">
  <rect width="${width}" height="${height}" rx="10" fill="#FFFFFF" stroke="#E1DFDD" stroke-width="1.5"/>
  <text x="24" y="30" font-family="Segoe UI, Arial, sans-serif" font-size="16" font-weight="700" fill="${escapeXml(brand)}">${escapeXml(params.wellName)} — Wellbore &amp; ESP Completion Schematic</text>
  <text x="24" y="50" font-family="Segoe UI, Arial, sans-serif" font-size="11.5" fill="#484644">Tubing: ${escapeXml(tubing)}  |  ESP Intake: ${espDepth.toLocaleString()} ft  |  Perforations: ${perfTop.toLocaleString()}–${perfBot.toLocaleString()} ft  |  TD: ${prodDepth.toLocaleString()} ft</text>

  <!-- Ground Surface Line & Wellhead Christmas Tree -->
  <line x1="60" y1="${topY}" x2="440" y2="${topY}" stroke="#8A8886" stroke-width="2" stroke-dasharray="6,3"/>
  <text x="66" y="${topY - 6}" font-family="Segoe UI, Arial, sans-serif" font-size="10" font-weight="600" fill="#605E5C">Surface (0 ft RKB)</text>
  <rect x="${centerX - 24}" y="${topY - 20}" width="48" height="20" rx="3" fill="${escapeXml(brand)}"/>
  <text x="${centerX}" y="${topY - 7}" text-anchor="middle" font-family="Segoe UI, Arial, sans-serif" font-size="9.5" font-weight="700" fill="#FFFFFF">WELLHEAD</text>

  <!-- Reservoir Pay Zone Band -->
  <rect x="90" y="${perfTopY - 8}" width="320" height="${Math.max(26, perfBotY - perfTopY + 16)}" rx="4" fill="#FFF4CE" stroke="#F7B500" stroke-width="1" stroke-dasharray="4,3"/>
  <text x="100" y="${(perfTopY + perfBotY) / 2 + 4}" font-family="Segoe UI, Arial, sans-serif" font-size="10.5" font-weight="700" fill="#8A5A00">${escapeXml(zoneName)}</text>

  <!-- Cement Annulus & Surface Casing -->
  <rect x="${centerX - 66}" y="${topY}" width="14" height="${surfY - topY}" fill="#D2D0CE"/>
  <rect x="${centerX + 52}" y="${topY}" width="14" height="${surfY - topY}" fill="#D2D0CE"/>
  <line x1="${centerX - 52}" y1="${topY}" x2="${centerX - 52}" y2="${surfY}" stroke="#323130" stroke-width="3"/>
  <line x1="${centerX + 52}" y1="${topY}" x2="${centerX + 52}" y2="${surfY}" stroke="#323130" stroke-width="3"/>
  <polygon points="${centerX - 52},${surfY} ${centerX - 64},${surfY} ${centerX - 52},${surfY - 12}" fill="#323130"/>
  <polygon points="${centerX + 52},${surfY} ${centerX + 64},${surfY} ${centerX + 52},${surfY - 12}" fill="#323130"/>

  <!-- Production Casing -->
  <line x1="${centerX - 36}" y1="${topY}" x2="${centerX - 36}" y2="${botY}" stroke="#201F1E" stroke-width="2.5"/>
  <line x1="${centerX + 36}" y1="${topY}" x2="${centerX + 36}" y2="${botY}" stroke="#201F1E" stroke-width="2.5"/>
  <line x1="${centerX - 36}" y1="${botY}" x2="${centerX + 36}" y2="${botY}" stroke="#201F1E" stroke-width="3"/>

  <!-- Production Tubing String -->
  <line x1="${centerX - 10}" y1="${topY}" x2="${centerX - 10}" y2="${espY - 22}" stroke="#0F6CBD" stroke-width="2.5"/>
  <line x1="${centerX + 10}" y1="${topY}" x2="${centerX + 10}" y2="${espY - 22}" stroke="#0F6CBD" stroke-width="2.5"/>

  <!-- ESP Power Cable along Tubing -->
  <line x1="${centerX + 15}" y1="${topY}" x2="${centerX + 15}" y2="${espY}" stroke="#D83B01" stroke-width="2" stroke-dasharray="4,2"/>

  <!-- ESP Assembly (Pump + Intake + Protector + Motor) -->
  <rect x="${centerX - 16}" y="${espY - 22}" width="32" height="44" rx="4" fill="${escapeXml(brand)}" stroke="#FFFFFF" stroke-width="1.5"/>
  <text x="${centerX}" y="${espY - 4}" text-anchor="middle" font-family="Segoe UI, Arial, sans-serif" font-size="9.5" font-weight="700" fill="#FFFFFF">ESP</text>
  <text x="${centerX}" y="${espY + 10}" text-anchor="middle" font-family="Segoe UI, Arial, sans-serif" font-size="8.5" fill="#FFFFFF">PUMP</text>

  <!-- Production Packer -->
  <rect x="${centerX - 35}" y="${packerY - 6}" width="24" height="12" fill="#323130"/>
  <rect x="${centerX + 11}" y="${packerY - 6}" width="24" height="12" fill="#323130"/>

  <!-- Perforation Intervals -->
  <line x1="${centerX - 36}" y1="${perfTopY}" x2="${centerX - 54}" y2="${perfTopY - 5}" stroke="#D13438" stroke-width="2.2"/>
  <line x1="${centerX + 36}" y1="${perfTopY}" x2="${centerX + 54}" y2="${perfTopY - 5}" stroke="#D13438" stroke-width="2.2"/>
  <line x1="${centerX - 36}" y1="${(perfTopY + perfBotY) / 2}" x2="${centerX - 56}" y2="${(perfTopY + perfBotY) / 2}" stroke="#D13438" stroke-width="2.2"/>
  <line x1="${centerX + 36}" y1="${(perfTopY + perfBotY) / 2}" x2="${centerX + 56}" y2="${(perfTopY + perfBotY) / 2}" stroke="#D13438" stroke-width="2.2"/>
  <line x1="${centerX - 36}" y1="${perfBotY}" x2="${centerX - 54}" y2="${perfBotY + 5}" stroke="#D13438" stroke-width="2.2"/>
  <line x1="${centerX + 36}" y1="${perfBotY}" x2="${centerX + 54}" y2="${perfBotY + 5}" stroke="#D13438" stroke-width="2.2"/>

  <!-- Right-side Engineering Depth Callouts -->
  <line x1="${centerX + 64}" y1="${surfY}" x2="430" y2="${surfY}" stroke="#A19F9D" stroke-width="1"/>
  <text x="436" y="${surfY + 4}" font-family="Segoe UI, Arial, sans-serif" font-size="11" font-weight="600" fill="#323130">Surface Casing Shoe: ${surfDepth.toLocaleString()} ft</text>

  <line x1="${centerX + 18}" y1="${espY}" x2="430" y2="${espY}" stroke="${escapeXml(brand)}" stroke-width="1.2"/>
  <text x="436" y="${espY - 2}" font-family="Segoe UI, Arial, sans-serif" font-size="11" font-weight="700" fill="${escapeXml(brand)}">ESP Pump Intake: ${espDepth.toLocaleString()} ft</text>
  <text x="436" y="${espY + 13}" font-family="Segoe UI, Arial, sans-serif" font-size="10" fill="#605E5C">${params.pipPsi ? `PIP: ${params.pipPsi} psi` : "Active VSD Monitoring"}${params.frequencyHz ? ` | ${params.frequencyHz} Hz` : ""}</text>

  <line x1="${centerX + 36}" y1="${packerY}" x2="430" y2="${packerY}" stroke="#605E5C" stroke-width="1"/>
  <text x="436" y="${packerY + 4}" font-family="Segoe UI, Arial, sans-serif" font-size="10.5" fill="#323130">Production Packer: ${packerDepth.toLocaleString()} ft</text>

  <line x1="${centerX + 56}" y1="${(perfTopY + perfBotY) / 2}" x2="430" y2="${(perfTopY + perfBotY) / 2}" stroke="#D13438" stroke-width="1.2"/>
  <text x="436" y="${(perfTopY + perfBotY) / 2 + 4}" font-family="Segoe UI, Arial, sans-serif" font-size="11" font-weight="700" fill="#D13438">Perforations: ${perfTop.toLocaleString()}–${perfBot.toLocaleString()} ft</text>

  <line x1="${centerX + 36}" y1="${botY}" x2="430" y2="${botY}" stroke="#323130" stroke-width="1"/>
  <text x="436" y="${botY + 4}" font-family="Segoe UI, Arial, sans-serif" font-size="11" font-weight="600" fill="#323130">Total Depth (TD): ${prodDepth.toLocaleString()} ft</text>
</svg>`;
}

/**
 * Renders a Surface Facility / Engineering Process Flowchart as a crisp SVG diagram.
 */
export function renderProcessFlowchartSvg(params: ProcessFlowchartParams): string {
  const width = 680;
  const height = 360;
  const brand = params.primaryColor || "#B73618";
  const nodes = (params.nodes || []).slice(0, 6);
  const count = Math.max(1, nodes.length);

  const cols = count <= 3 ? count : 3;
  const boxW = 184;
  const boxH = 88;
  const gapX = 36;
  const gapY = 48;
  const startX = 28;
  const startY = 86;

  const statusColor = (s?: ProcessFlowNode["status"]) => {
    if (s === "critical") return "#D13438";
    if (s === "warning") return "#D83B01";
    return brand;
  };

  const renderedNodes: string[] = [];
  const renderedArrows: string[] = [];

  nodes.forEach((node, idx) => {
    const row = Math.floor(idx / cols);
    const colInRow = idx % cols;
    // Snake layout: row 0 left-to-right, row 1 right-to-left for continuous flow
    const col = row % 2 === 0 ? colInRow : cols - 1 - colInRow;
    const x = startX + col * (boxW + gapX);
    const y = startY + row * (boxH + gapY);
    const accent = statusColor(node.status);

    renderedNodes.push(`<g transform="translate(${x}, ${y})">
      <rect width="${boxW}" height="${boxH}" rx="8" fill="#FAF9F8" stroke="${escapeXml(accent)}" stroke-width="2"/>
      <rect width="${boxW}" height="24" rx="6" fill="${escapeXml(accent)}"/>
      <text x="10" y="16" font-family="Segoe UI, Arial, sans-serif" font-size="10.5" font-weight="700" fill="#FFFFFF">STEP ${idx + 1}</text>
      <text x="12" y="45" font-family="Segoe UI, Arial, sans-serif" font-size="12" font-weight="700" fill="#201F1E">${escapeXml(node.label.slice(0, 24))}</text>
      ${node.subLabel ? `<text x="12" y="62" font-family="Segoe UI, Arial, sans-serif" font-size="10.5" fill="#605E5C">${escapeXml(node.subLabel.slice(0, 28))}</text>` : ""}
      ${node.metric ? `<text x="12" y="78" font-family="Segoe UI, Arial, sans-serif" font-size="10.5" font-weight="700" fill="${escapeXml(accent)}">${escapeXml(node.metric.slice(0, 26))}</text>` : ""}
    </g>`);

    if (idx < nodes.length - 1) {
      const nextRow = Math.floor((idx + 1) / cols);
      const nextColInRow = (idx + 1) % cols;
      const nextCol = nextRow % 2 === 0 ? nextColInRow : cols - 1 - nextColInRow;
      const nx = startX + nextCol * (boxW + gapX);
      const ny = startY + nextRow * (boxH + gapY);

      if (nextRow === row) {
        if (nextCol > col) {
          const ax1 = x + boxW;
          const ax2 = nx;
          const ay = y + boxH / 2;
          renderedArrows.push(
            `<line x1="${ax1}" y1="${ay}" x2="${ax2}" y2="${ay}" stroke="#323130" stroke-width="2.2"/>`,
            `<polygon points="${ax2},${ay} ${ax2 - 8},${ay - 5} ${ax2 - 8},${ay + 5}" fill="#323130"/>`
          );
        } else {
          const ax1 = x;
          const ax2 = nx + boxW;
          const ay = y + boxH / 2;
          renderedArrows.push(
            `<line x1="${ax1}" y1="${ay}" x2="${ax2}" y2="${ay}" stroke="#323130" stroke-width="2.2"/>`,
            `<polygon points="${ax2},${ay} ${ax2 + 8},${ay - 5} ${ax2 + 8},${ay + 5}" fill="#323130"/>`
          );
        }
      } else {
        const ax = x + boxW / 2;
        const ay1 = y + boxH;
        const ay2 = ny;
        renderedArrows.push(
          `<line x1="${ax}" y1="${ay1}" x2="${ax}" y2="${ay2}" stroke="#323130" stroke-width="2.2"/>`,
          `<polygon points="${ax},${ay2} ${ax - 5},${ay2 - 8} ${ax + 5},${ay2 - 8}" fill="#323130"/>`
        );
      }
    }
  });

  return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">
  <rect width="${width}" height="${height}" rx="10" fill="#FFFFFF" stroke="#E1DFDD" stroke-width="1.5"/>
  <text x="24" y="32" font-family="Segoe UI, Arial, sans-serif" font-size="16" font-weight="700" fill="${escapeXml(brand)}">${escapeXml(params.title)}</text>
  <text x="24" y="52" font-family="Segoe UI, Arial, sans-serif" font-size="11.5" fill="#484644">${escapeXml(params.subtitle || "Process & Surface Facility Flow Diagram")}</text>
  ${renderedArrows.join("\n  ")}
  ${renderedNodes.join("\n  ")}
</svg>`;
}

/**
 * Renders a Workover / Rig Schedule or Project Execution Timeline as a crisp SVG diagram.
 */
export function renderWorkoverTimelineSvg(params: TimelineDiagramParams): string {
  const width = 680;
  const height = 360;
  const brand = params.primaryColor || "#B73618";
  const milestones = (params.milestones || []).slice(0, 6);
  const count = Math.max(1, milestones.length);

  const rowH = Math.min(44, Math.floor(245 / count));
  const startY = 82;

  const badgeFill = (status?: TimelineMilestone["status"]) => {
    if (status === "completed") return "#107C10";
    if (status === "active") return brand;
    return "#0F6CBD";
  };

  return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">
  <rect width="${width}" height="${height}" rx="10" fill="#FFFFFF" stroke="#E1DFDD" stroke-width="1.5"/>
  <text x="24" y="32" font-family="Segoe UI, Arial, sans-serif" font-size="16" font-weight="700" fill="${escapeXml(brand)}">${escapeXml(params.title)}</text>
  <text x="24" y="52" font-family="Segoe UI, Arial, sans-serif" font-size="11.5" fill="#484644">${escapeXml(params.subtitle || "Engineering Execution & Workover Schedule")}</text>
  <line x1="128" y1="${startY - 8}" x2="128" y2="${startY + count * (rowH + 8)}" stroke="${escapeXml(brand)}" stroke-width="2.5"/>
  ${milestones
    .map((m, idx) => {
      const y = startY + idx * (rowH + 8);
      const color = badgeFill(m.status);
      const barLeft = 148 + idx * 45;
      const barW = Math.max(180, 480 - idx * 45);
      return `<g transform="translate(0, ${y})">
        <text x="24" y="22" font-family="Segoe UI, Arial, sans-serif" font-size="11" font-weight="700" fill="#323130">${escapeXml(m.phase.slice(0, 14))}</text>
        <circle cx="128" cy="18" r="6" fill="${escapeXml(color)}" stroke="#FFFFFF" stroke-width="2"/>
        <rect x="${barLeft}" y="2" width="${barW}" height="${rowH - 6}" rx="6" fill="#FAF9F8" stroke="${escapeXml(color)}" stroke-width="1.5"/>
        <rect x="${barLeft}" y="2" width="8" height="${rowH - 6}" rx="3" fill="${escapeXml(color)}"/>
        <text x="${barLeft + 16}" y="18" font-family="Segoe UI, Arial, sans-serif" font-size="11.5" font-weight="700" fill="#201F1E">${escapeXml(m.task.slice(0, 44))}</text>
        <text x="${barLeft + 16}" y="32" font-family="Segoe UI, Arial, sans-serif" font-size="10" fill="#605E5C">${escapeXml(m.durationOrDate)}${m.owner ? `  •  Owner: ${escapeXml(m.owner)}` : ""}</text>
      </g>`;
    })
    .join("\n  ")}
</svg>`;
}

