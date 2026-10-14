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
