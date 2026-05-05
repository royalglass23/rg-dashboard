// deploy.js — Royal Glass Apps Script deploy tool
// Usage: node deploy.js staging
//        node deploy.js live

const fs   = require('fs');
const path = require('path');
const { execSync } = require('child_process');

const ENV_OPTIONS = ['staging', 'live'];
const env = process.argv[2];

if (!env || !ENV_OPTIONS.includes(env)) {
  console.error('\n❌  Usage: node deploy.js staging|live\n');
  process.exit(1);
}

const sourceFile = path.resolve(`.clasp.${env}.json`);
const targetFile = path.resolve('.clasp.json');

if (!fs.existsSync(sourceFile)) {
  console.error(`\n❌  Missing config file: .clasp.${env}.json\n`);
  process.exit(1);
}

console.log(`\n🚀  Deploying to ${env.toUpperCase()}...`);

// Swap the active clasp config
fs.copyFileSync(sourceFile, targetFile);

try {
  execSync('clasp push --force', { stdio: 'inherit' });
  console.log(`\n✅  Deployed to ${env.toUpperCase()} successfully.\n`);
} catch (e) {
  console.error(`\n❌  Deploy failed. Check clasp output above.\n`);
  process.exit(1);
}
