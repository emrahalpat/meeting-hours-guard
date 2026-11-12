// Builds the GitHub Pages site (docs/) and the production manifest (manifest.prod.xml).
// Usage: node build.js https://<user>.github.io/<repo>
const fs = require("fs");
const path = require("path");

const base = (process.argv[2] || "").replace(/\/$/, "");
if (!/^https:\/\//.test(base)) {
  console.error("Usage: node build.js https://host/path");
  process.exit(1);
}
const origin = new URL(base).origin;

const out = path.join(__dirname, "docs");
fs.rmSync(out, { recursive: true, force: true });
fs.mkdirSync(path.join(out, "assets"), { recursive: true });
for (const f of fs.readdirSync(path.join(__dirname, "src"))) fs.copyFileSync(path.join(__dirname, "src", f), path.join(out, f));
for (const f of fs.readdirSync(path.join(__dirname, "assets"))) fs.copyFileSync(path.join(__dirname, "assets", f), path.join(out, "assets", f));
fs.writeFileSync(path.join(out, ".nojekyll"), "");

let manifest = fs.readFileSync(path.join(__dirname, "manifest.xml"), "utf8");
manifest = manifest
  .replace("<AppDomain>https://localhost:3000</AppDomain>", `<AppDomain>${origin}</AppDomain>`)
  .replaceAll("https://localhost:3000", base);
fs.writeFileSync(path.join(__dirname, "manifest.prod.xml"), manifest);
console.log(`docs/ built; manifest.prod.xml points at ${base}`);
