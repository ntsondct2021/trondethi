import express from "express";
import { createServer as createViteServer } from "vite";
import path from "path";
import { fileURLToPath } from "url";
import cors from "cors";
import multer from "multer";
import Database from "better-sqlite3";
import fs from "fs";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Ensure uploads directory exists
const uploadsDir = path.join(process.cwd(), "uploads");
if (!fs.existsSync(uploadsDir)) {
  fs.mkdirSync(uploadsDir);
}

// Multer configuration
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, "uploads/");
  },
  filename: (req, file, cb) => {
    const uniqueSuffix = Date.now() + "-" + Math.round(Math.random() * 1e9);
    cb(null, uniqueSuffix + path.extname(file.originalname));
  },
});
const upload = multer({ storage });

async function startServer() {
  const app = express();
  const PORT = Number(process.env.PORT) || 3000;

  app.use(cors());
  app.use(express.json({ limit: "100mb" }));
  app.use(express.urlencoded({ limit: "100mb", extended: true }));
  app.use("/uploads", express.static("uploads"));

  // Database setup
  const db = new Database("exam_bank.db");
  db.exec(`
    CREATE TABLE IF NOT EXISTS questions (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      type TEXT NOT NULL DEFAULT 'SINGLE_CHOICE',
      content TEXT NOT NULL,
      option_a TEXT,
      option_b TEXT,
      option_c TEXT,
      option_d TEXT,
      correct_answer TEXT NOT NULL,
      topic TEXT,
      difficulty TEXT,
      image_url TEXT,
      explanation TEXT,
      created_at DATETIME DEFAULT CURRENT_TIMESTAMP
    );

    CREATE TABLE IF NOT EXISTS exam_structures (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      name TEXT NOT NULL,
      total_questions INTEGER NOT NULL,
      difficulty_config TEXT, -- JSON string
      topic_config TEXT -- JSON string
    );

    CREATE TABLE IF NOT EXISTS math_assets (
      id TEXT PRIMARY KEY,
      original_xml TEXT NOT NULL,
      ole_bin_base64 TEXT,
      image_bin_base64 TEXT,
      image_ext TEXT,
      latex_fallback TEXT
    );
  `);

  // Migration: Add type column if it doesn't exist
  try {
    db.prepare("ALTER TABLE questions ADD COLUMN type TEXT NOT NULL DEFAULT 'SINGLE_CHOICE'").run();
  } catch (e) {
    // Column already exists
  }

  // API Routes
  app.get("/api/health", (req, res) => {
    try {
      const count = db.prepare("SELECT COUNT(*) as count FROM questions").get().count;
      res.json({ status: "ok", database: "connected", questionCount: count });
    } catch (error) {
      res.status(500).json({ status: "error", message: error instanceof Error ? error.message : "Database error" });
    }
  });

  app.get("/api/questions", (req, res) => {
    try {
      const questions = db.prepare("SELECT * FROM questions ORDER BY created_at DESC").all();
      res.json(questions);
    } catch (error) {
      console.error("Error fetching questions:", error);
      res.status(500).json({ error: "Failed to fetch questions" });
    }
  });

  app.post("/api/questions", (req, res) => {
    try {
      const { type, content, option_a, option_b, option_c, option_d, correct_answer, topic, difficulty, image_url, explanation } = req.body;
      const info = db.prepare(`
        INSERT INTO questions (type, content, option_a, option_b, option_c, option_d, correct_answer, topic, difficulty, image_url, explanation)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      `).run(type || 'SINGLE_CHOICE', content, option_a, option_b, option_c, option_d, correct_answer, topic, difficulty, image_url, explanation);
      res.json({ id: info.lastInsertRowid });
    } catch (error) {
      console.error("Error adding question:", error);
      res.status(500).json({ error: error instanceof Error ? error.message : "Failed to add question" });
    }
  });

  app.post("/api/questions/bulk", (req, res) => {
    const questions = req.body;
    const insert = db.prepare(`
      INSERT INTO questions (type, content, option_a, option_b, option_c, option_d, correct_answer, topic, difficulty, image_url, explanation)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `);
    
    const insertMany = db.transaction((qs) => {
      for (const q of qs) {
        insert.run(q.type || 'SINGLE_CHOICE', q.content, q.option_a, q.option_b, q.option_c, q.option_d, q.correct_answer, q.topic, q.difficulty, q.image_url, q.explanation);
      }
    });

    insertMany(questions);
    res.json({ success: true, count: questions.length });
  });

  app.delete("/api/questions/:id", (req, res) => {
    db.prepare("DELETE FROM questions WHERE id = ?").run(req.params.id);
    res.json({ success: true });
  });

  app.post("/api/questions/delete-bulk", (req, res) => {
    const { ids } = req.body;
    if (!Array.isArray(ids) || ids.length === 0) {
      return res.status(400).json({ error: "Invalid IDs" });
    }

    const deleteStmt = db.prepare("DELETE FROM questions WHERE id = ?");
    const deleteMany = db.transaction((idsToDelete) => {
      for (const id of idsToDelete) {
        deleteStmt.run(id);
      }
    });

    deleteMany(ids);
    res.json({ success: true, count: ids.length });
  });

  app.put("/api/questions/:id", (req, res) => {
    const { type, content, option_a, option_b, option_c, option_d, correct_answer, topic, difficulty, image_url, explanation } = req.body;
    db.prepare(`
      UPDATE questions 
      SET type = ?, content = ?, option_a = ?, option_b = ?, option_c = ?, option_d = ?, correct_answer = ?, topic = ?, difficulty = ?, image_url = ?, explanation = ?
      WHERE id = ?
    `).run(type || 'SINGLE_CHOICE', content, option_a, option_b, option_c, option_d, correct_answer, topic, difficulty, image_url, explanation, req.params.id);
    res.json({ success: true });
  });

  app.get("/api/stats", (req, res) => {
    const total = db.prepare("SELECT COUNT(*) as count FROM questions").get().count;
    const difficultyStats = db.prepare("SELECT difficulty, COUNT(*) as count FROM questions GROUP BY difficulty").all();
    const topicStats = db.prepare("SELECT topic as subject, COUNT(*) as count FROM questions GROUP BY topic").all();
    res.json({ total, difficultyStats, topicStats });
  });

  app.post("/api/upload", upload.single("image"), (req: any, res) => {
    if (!req.file) {
      return res.status(400).json({ error: "No file uploaded" });
    }
    const url = `/uploads/${req.file.filename}`;
    res.json({ url });
  });

  // Math Assets Endpoints
  app.post("/api/math-assets", (req, res) => {
    try {
      const assets = req.body;
      if (!Array.isArray(assets)) {
        return res.status(400).json({ error: "Invalid body, expected array" });
      }
      const insert = db.prepare(`
        INSERT OR REPLACE INTO math_assets (id, original_xml, ole_bin_base64, image_bin_base64, image_ext, latex_fallback)
        VALUES (?, ?, ?, ?, ?, ?)
      `);
      const insertMany = db.transaction((list) => {
        for (const item of list) {
          insert.run(
            item.id,
            item.originalXml,
            item.oleBinBase64 || null,
            item.imageBinBase64 || null,
            item.imageExt || null,
            item.latexFallback || null
          );
        }
      });
      insertMany(assets);
      res.json({ success: true, count: assets.length });
    } catch (error) {
      console.error("Error saving math assets:", error);
      res.status(500).json({ error: "Failed to save math assets" });
    }
  });

  app.get("/api/math-assets/:id", (req, res) => {
    try {
      const asset = db.prepare("SELECT * FROM math_assets WHERE id = ?").get(req.params.id);
      if (!asset) {
        return res.status(404).json({ error: "Asset not found" });
      }
      res.json({
        id: asset.id,
        originalXml: asset.original_xml,
        oleBinBase64: asset.ole_bin_base64,
        imageBinBase64: asset.image_bin_base64,
        imageExt: asset.image_ext,
        latexFallback: asset.latex_fallback
      });
    } catch (error) {
      console.error("Error fetching math asset:", error);
      res.status(500).json({ error: "Failed to get math asset" });
    }
  });

  app.get("/api/math-assets/:id/image", (req, res) => {
    try {
      const asset = db.prepare("SELECT image_bin_base64, image_ext FROM math_assets WHERE id = ?").get(req.params.id);
      if (!asset || !asset.image_bin_base64) {
        return res.status(404).send("Not found");
      }
      const buffer = Buffer.from(asset.image_bin_base64, "base64");
      const ext = (asset.image_ext || "png").toLowerCase();
      let contentType = "image/png";
      if (ext === "wmf") contentType = "image/x-wmf";
      else if (ext === "emf") contentType = "image/x-emf";
      else if (ext === "gif") contentType = "image/gif";
      else if (ext === "jpg" || ext === "jpeg") contentType = "image/jpeg";

      res.setHeader("Content-Type", contentType);
      res.end(buffer);
    } catch (error) {
      console.error("Error serving asset image:", error);
      res.status(500).send("Server error");
    }
  });

  app.post("/api/math-assets/get-bulk", (req, res) => {
    try {
      const { ids } = req.body;
      if (!Array.isArray(ids) || ids.length === 0) {
        return res.json([]);
      }
      const stmt = db.prepare("SELECT * FROM math_assets WHERE id = ?");
      const results: any[] = [];
      for (const id of ids) {
        const asset = stmt.get(id);
        if (asset) {
          results.push({
            id: asset.id,
            originalXml: asset.original_xml,
            oleBinBase64: asset.ole_bin_base64,
            imageBinBase64: asset.image_bin_base64,
            imageExt: asset.image_ext,
            latexFallback: asset.latex_fallback
          });
        }
      }
      res.json(results);
    } catch (error) {
      console.error("Error fetching bulk math assets:", error);
      res.status(500).json({ error: "Failed to fetch math assets" });
    }
  });

  // Vite middleware for development
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", (req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();
