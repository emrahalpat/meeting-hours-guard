// Builds the hosted files and the production manifest (manifest.prod.xml).
// Usage: node build.js https://<your-host>/<path> [outputFolder]   (output folder defaults to docs/)
const fs = require("fs");
const path = require("path");

const base = (process.argv[2] || "").replace(/\/$/, "");
if (!/^https:\/\//.test(base)) {
  console.error("Usage: node build.js https://host/path");
  process.exit(1);
}
const origin = new URL(base).origin;

const out = path.resolve(__dirname, process.argv[3] || "docs");
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
fs.writeFileSync(path.join(out, "manifest.xml"), manifest);
console.log(`${path.relative(__dirname, out) || "."}/ built; manifest.prod.xml points at ${base}`);
