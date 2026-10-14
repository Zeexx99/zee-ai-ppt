import { readFileSync, writeFileSync, existsSync, copyFileSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const rootDir = join(__dirname, "..");

const devManifestPath = join(rootDir, "manifest.xml");
const prodManifestPath = join(rootDir, "manifest.prod.xml");
const distDir = join(rootDir, "dist");

const HOSTED_BASE_URL =
  process.env.ZEE_PPT_HOST_URL || "https://zeexx99.github.io/zee-ai-ppt";

function validateManifestXml(xml, label) {
  const errors = [];
  if (!/<Id>[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}<\/Id>/.test(xml)) {
    errors.push(`${label}: Missing or invalid GUID <Id>`);
  }
  if (!/<Host\s+Name="Presentation"\s*\/>/.test(xml)) {
    errors.push(`${label}: Missing <Host Name="Presentation"/>`);
  }
  if (!/<Host\s+xsi:type="Presentation"\s*>/.test(xml)) {
    errors.push(`${label}: Missing <Host xsi:type="Presentation"> in VersionOverrides`);
  }
  if (!/<OfficeTab\s+id="TabHome"\s*>/.test(xml)) {
    errors.push(`${label}: Missing <OfficeTab id="TabHome">`);
  }
  if (!/<TaskpaneId>ZeeAI\.PPT\.Taskpane<\/TaskpaneId>/.test(xml)) {
    errors.push(`${label}: TaskpaneId must be ZeeAI.PPT.Taskpane`);
  }
  if (errors.length > 0) {
    throw new Error(errors.join("\n"));
  }
}

const devXml = readFileSync(devManifestPath, "utf-8");
validateManifestXml(devXml, "manifest.xml");

const normalizedBase = HOSTED_BASE_URL.replace(/\/+$/, "");
const prodXml = devXml
  .replace(/<AppDomain>https:\/\/localhost:3300<\/AppDomain>\s*/g, "")
  .replace(/https:\/\/localhost:3300/g, normalizedBase);

validateManifestXml(prodXml, "manifest.prod.xml");

writeFileSync(prodManifestPath, prodXml, "utf-8");
console.log(`Generated ${prodManifestPath} -> ${normalizedBase}/`);

if (existsSync(distDir)) {
  copyFileSync(prodManifestPath, join(distDir, "manifest.xml"));
  copyFileSync(prodManifestPath, join(distDir, "manifest.prod.xml"));
  console.log("Copied manifest.prod.xml to dist/manifest.xml and dist/manifest.prod.xml");
}
