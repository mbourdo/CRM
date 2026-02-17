const http = require("http");
const fs = require("fs");
const path = require("path");
const { URL } = require("url");

const PORT = process.env.PORT || 3000;
const ROOT = __dirname;
const DATA_DIR = path.join(ROOT, "data");
const DATA_FILE = path.join(DATA_DIR, "leads.json");

const MIME_TYPES = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".svg": "image/svg+xml",
  ".ico": "image/x-icon",
};

ensureDataStore();

const server = http.createServer(async (req, res) => {
  const requestUrl = new URL(req.url, `http://${req.headers.host}`);

  if (requestUrl.pathname === "/api/leads") {
    if (req.method === "GET") {
      return handleGetLeads(res);
    }
    if (req.method === "PUT") {
      return handlePutLeads(req, res);
    }
    return send(res, 405, { error: "Method not allowed" });
  }

  if (req.method !== "GET") {
    return send(res, 405, "Method not allowed", "text/plain; charset=utf-8");
  }

  return serveStatic(requestUrl.pathname, res);
});

server.listen(PORT, () => {
  // eslint-disable-next-line no-console
  console.log(`CRM running at http://localhost:${PORT}`);
});

function ensureDataStore() {
  if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
  if (!fs.existsSync(DATA_FILE)) fs.writeFileSync(DATA_FILE, "[]", "utf8");
}

function readLeads() {
  try {
    const content = fs.readFileSync(DATA_FILE, "utf8");
    const parsed = JSON.parse(content);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function writeLeads(leads) {
  fs.writeFileSync(DATA_FILE, JSON.stringify(leads, null, 2), "utf8");
}

function handleGetLeads(res) {
  const leads = readLeads();
  return send(res, 200, leads);
}

function handlePutLeads(req, res) {
  let body = "";
  req.on("data", (chunk) => {
    body += chunk.toString();
    if (body.length > 20 * 1024 * 1024) {
      req.destroy();
    }
  });
  req.on("end", () => {
    try {
      const parsed = JSON.parse(body || "[]");
      if (!Array.isArray(parsed)) {
        return send(res, 400, { error: "Expected an array of leads" });
      }
      writeLeads(parsed);
      return send(res, 200, { ok: true, count: parsed.length });
    } catch {
      return send(res, 400, { error: "Invalid JSON" });
    }
  });
  req.on("error", () => send(res, 500, { error: "Request failed" }));
}

function serveStatic(rawPath, res) {
  const pathname = rawPath === "/" ? "index.html" : rawPath.replace(/^\/+/, "");
  const filePath = path.join(ROOT, path.normalize(pathname));
  if (!filePath.startsWith(ROOT)) {
    return send(res, 403, "Forbidden", "text/plain; charset=utf-8");
  }

  fs.readFile(filePath, (err, content) => {
    if (err) {
      return send(res, 404, "Not found", "text/plain; charset=utf-8");
    }
    const ext = path.extname(filePath).toLowerCase();
    const contentType = MIME_TYPES[ext] || "application/octet-stream";
    send(res, 200, content, contentType);
  });
}

function send(res, status, payload, contentType = "application/json; charset=utf-8") {
  const body =
    typeof payload === "string" || Buffer.isBuffer(payload)
      ? payload
      : JSON.stringify(payload);
  res.writeHead(status, {
    "Content-Type": contentType,
    "Cache-Control": "no-store",
  });
  res.end(body);
}
