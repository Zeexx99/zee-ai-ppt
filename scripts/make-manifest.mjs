import fs from 'node:fs';
import path from 'node:path';

const devManifestPath = path.resolve('manifest.xml');
const prodManifestPath = path.resolve('manifest.prod.xml');
const distDir = path.resolve('dist');

if (!fs.existsSync(devManifestPath)) {
  console.error('Error: manifest.xml not found!');
  process.exit(1);
}

let content = fs.readFileSync(devManifestPath, 'utf8');

// Replace local dev URLs with production GitHub Pages URL
content = content
  .replaceAll('https://localhost:3300/', 'https://zeexx99.github.io/zee-ai-ppt/')
  .replaceAll('https://localhost:3300', 'https://zeexx99.github.io/zee-ai-ppt');

// Add zeexx99.github.io to AppDomains if needed
if (!content.includes('https://zeexx99.github.io')) {
  content = content.replace(
    '</AppDomains>',
    '  <AppDomain>https://zeexx99.github.io</AppDomain>\n  </AppDomains>'
  );
}

fs.writeFileSync(prodManifestPath, content, 'utf8');
console.log(`Created ${prodManifestPath}`);

if (fs.existsSync(distDir)) {
  const distManifestPath = path.join(distDir, 'manifest.xml');
  fs.writeFileSync(distManifestPath, content, 'utf8');
  console.log(`Copied to ${distManifestPath}`);
}
