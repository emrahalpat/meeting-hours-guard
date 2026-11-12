// Minimal HTTPS static server for local add-in development (https://localhost:3000).
// Uses the trusted localhost certificate created by `npm run certs`.
const https = require("https");
const fs = require("fs");
const path = require("path");
const os = require("os");

const PORT = 3000;
const CERT_DIR = path.join(os.homedir(), ".office-addin-dev-certs");
const ROOTS = { "/assets/": path.join(__dirname, "assets"), "/": path.join(__dirname, "src") };
const TYPES = { ".html": "text/html; charset=utf-8", ".js": "text/javascript; charset=utf-8", ".png": "image/png", ".xml": "application/xml" };

let key, cert;
try {
  key = fs.readFileSync(path.join(CERT_DIR, "localhost.key"));
  cert = fs.readFileSync(path.join(CERT_DIR, "localhost.crt"));
} catch {
  console.error("No dev certificate found. Run: npm run certs");
  process.exit(1);
}

https
  .createServer({ key, cert }, (req, res) => {
    const url = decodeURIComponent(new URL(req.url, "https://localhost").pathname);
    const prefix = url.startsWith("/assets/") ? "/assets/" : "/";
    const root = ROOTS[prefix];
    const file = path.normalize(path.join(root, url.slice(prefix.length) || "commands.html"));
    if (!file.startsWith(root)) { res.writeHead(403).end(); return; }
    fs.readFile(file, (err, data) => {
      if (err) { res.writeHead(404).end("Not found"); return; }
      res.writeHead(200, {
        "Content-Type": TYPES[path.extname(file)] || "application/octet-stream",
        "Cache-Control": "no-store",
        "Access-Control-Allow-Origin": "*",
      });
      res.end(data);
    });
    console.log(new Date().toISOString(), req.method, url);
  })
  .listen(PORT, () => console.log(`Meeting Hours Guard served at https://localhost:${PORT}`));
