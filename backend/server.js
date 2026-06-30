import express from "express";
import cors from "cors";
import multer from "multer";
import fs from "fs";
import path from "path";
import PDFMerger from "pdf-merger-js";
import mammoth from "mammoth";
import htmlToDocx from "html-to-docx";
import { spawn } from "child_process";

// hello
const app = express();

// Request logging for production debugging
app.use((req, res, next) => {
  console.log(`[${new Date().toISOString()}] ${req.method} ${req.originalUrl}`);
  next();
});

app.use(cors({ exposedHeaders: ["Content-Disposition"] }));
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

const UPLOAD_DIR = path.resolve("./uploads");
const OUTPUT_DIR = path.resolve("./output");

if (!fs.existsSync(UPLOAD_DIR)) fs.mkdirSync(UPLOAD_DIR, { recursive: true });
if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR, { recursive: true });

const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, UPLOAD_DIR),
  filename: (req, file, cb) => cb(null, `${Date.now()}-${file.originalname}`),
});
const upload = multer({ storage });

function cleanupFiles(files) {
  for (const f of files) {
    try { if (fs.existsSync(f.path)) fs.unlinkSync(f.path); } catch (e) {}
  }
}

app.get("/", (req, res) => res.json({ ok: true }));

// ---------------------------------------------------------------------------
// PDF / DOCX merge
// ---------------------------------------------------------------------------
app.post("/merge", upload.array("files"), async (req, res) => {
  const files = req.files || [];
  const type = (req.body.type || "pdf").toLowerCase();
  console.log(`[/merge] Type: ${type}, Files: ${files.length}`);

  if (!files.length) return res.status(400).json({ error: "No files uploaded" });

  const outExt = type === "pdf-docx" ? "docx" : type;
  const outName = `merged-${Date.now()}.${outExt}`;
  const outPath = path.join(OUTPUT_DIR, outName);

  try {
    if (type === "pdf") {
      const merger = new PDFMerger();
      for (const f of files) await merger.add(f.path);
      await merger.save(outPath);
      cleanupFiles(files);
      return res.download(outPath, outName, () => {
        try { if (fs.existsSync(outPath)) fs.unlinkSync(outPath); } catch (e) {}
      });
    }

    if (type === "docx" || type === "pdf-docx") {
      const filePaths = files.map(f => f.path);
      await new Promise((resolve, reject) => {
        const py = spawn("python", [path.resolve("./merge.py"), outPath, ...filePaths]);
        let stderr = "";
        py.stderr.on("data", d => { stderr += d.toString(); });
        py.on("close", code => code !== 0 ? reject(new Error(stderr || `Exit ${code}`)) : resolve());
        py.on("error", reject);
      });
      cleanupFiles(files);
      return res.download(outPath, outName, () => {
        try { if (fs.existsSync(outPath)) fs.unlinkSync(outPath); } catch (e) {}
      });
    }

    throw new Error(`Unsupported type: ${type}`);
  } catch (err) {
    cleanupFiles(files);
    return res.status(500).json({ error: err.message });
  }
});



// ---------------------------------------------------------------------------
// AMCP Dossier generation — always generates, always downloads
// ---------------------------------------------------------------------------
app.post("/dossier/generate", upload.array("files"), async (req, res) => {
  const files = req.files || [];
  console.log(`[/dossier/generate] Received ${files.length} files`);

  let configPath = null;
  let outPath = null;

  try {
    const metadata = JSON.parse(req.body.metadata || "{}");
    const sectionMappings = JSON.parse(req.body.mappings || "{}");

    const resolvedMappings = {};
    for (const [sectionKey, originalName] of Object.entries(sectionMappings)) {
      if (!originalName) continue;
      const fileObj = files.find(f => f.originalname === originalName);
      if (fileObj) resolvedMappings[sectionKey] = fileObj.path;
    }

    const configData = {
      metadata,
      sections: resolvedMappings,
      dossierType: metadata.dossierType ? metadata.dossierType.toLowerCase() : "format-a",
    };

    configPath = path.join(OUTPUT_DIR, `config-${Date.now()}.json`);
    fs.writeFileSync(configPath, JSON.stringify(configData, null, 2));

    const outName = `dossier-${Date.now()}.docx`;
    outPath = path.join(OUTPUT_DIR, outName);

    await new Promise((resolve, reject) => {
      const py = spawn("python", [path.resolve("./dossier_builder.py"), configPath, outPath]);
      let stdout = "";
      let stderr = "";
      py.stdout.on("data", d => { stdout += d.toString(); });
      py.stderr.on("data", d => { stderr += d.toString(); });
      py.on("close", code => {
        console.log("[Python Dossier STDOUT]:\n", stdout);
        if (code !== 0) {
          console.error("[Python Dossier STDERR]:\n", stderr);
          reject(new Error(stderr.trim() || `Python exited with code ${code}`));
        } else {
          resolve();
        }
      });
      py.on("error", reject);
    });

    try { if (fs.existsSync(configPath)) fs.unlinkSync(configPath); } catch (e) {}
    cleanupFiles(files);

    return res.download(outPath, outName, () => {
      try { if (fs.existsSync(outPath)) fs.unlinkSync(outPath); } catch (e) {}
    });

  } catch (err) {
    console.error("[/dossier/generate] Error:", err.message || err);
    try { if (configPath && fs.existsSync(configPath)) fs.unlinkSync(configPath); } catch (e) {}
    cleanupFiles(files);
    return res.status(500).json({ error: err.message || "Failed to generate dossier" });
  }
});

// Catch-all 404 handler to debug unmatched routes
app.use((req, res) => {
  console.warn(`[404] Unmatched Route: ${req.method} ${req.originalUrl}`);
  res.status(404).json({ error: "Route not found" });
});

const port = process.env.PORT || 3000;
app.listen(port, () => console.log(`Document merging backend running on http://localhost:${port}`));
