// server.js (Node >= 18, package.json has "type": "module")
import express from "express";
import bodyParser from "body-parser";
import cors from "cors";
import morgan from "morgan";
import multer from "multer";
import fs from "fs";
import path from "path";
import odbc from "odbc";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// -------------------- App setup --------------------
const app = express();
app.use(cors());
app.use(bodyParser.json());
app.use(morgan("dev"));

// -------------------- Access ODBC --------------------
const dbPath = "C:\\codes\\Student_db\\StudentResult.accdb";
const connectionString = `Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=${dbPath};`;

// Helper: escape string literals for Access SQL
function esc(v) {
  if (v === null || v === undefined) return "NULL";
  return `'${String(v).replace(/'/g, "''")}'`;
}

// Run raw SQL (no params) – avoids Access ODBC parameter metadata issue
async function execSql(sql) {
  const conn = await odbc.connect(connectionString);
  try {
    return await conn.query(sql);
  } finally {
    await conn.close();
  }
}

// -------------------- Routes --------------------

// Health
app.get("/", (_req, res) => {
  res.send("Student API running. Try GET /students");
});

// Optional: verify uploads folder exists
app.get("/health/uploads", (_req, res) => {
  try {
    const dir = path.join(__dirname, "Marksheets");
    fs.mkdirSync(dir, { recursive: true });
    res.json({ ok: true, path: dir });
  } catch (e) {
    res.status(500).json({ ok: false, error: String(e) });
  }
});

// GET all students
app.get("/students", async (_req, res) => {
  try {
    const rows = await execSql("SELECT * FROM Students ORDER BY RollNumber");

    // Normalize MarksFilePath to a browser-safe form
    const normalized = rows.map((r) => {
      let p = r.MarksFilePath || "";
      if (p) {
        p = p.replace(/\\/g, "/"); // backslashes -> slashes
        // ensure leading slash for the static route
        if (!p.startsWith("/")) {
          if (p.toLowerCase().startsWith("marksheets/")) p = `/${p}`;
        }
      }
      return { ...r, MarksFilePath: p };
    });

    res.json(normalized);
  } catch (e) {
    console.error("SELECT error:", e);
    res.status(500).json({ error: "Error fetching students" });
  }
});

// Create student (no params; use escaping)
app.post("/students", async (req, res) => {
  try {
    const { firstName, lastName, marksFilePath } = req.body;
    if (!firstName || !lastName) {
      return res.status(400).json({ error: "FirstName and LastName are required" });
    }

    const sql = `
      INSERT INTO Students (FirstName, LastName, MarksFilePath)
      VALUES (${esc(firstName)}, ${esc(lastName)}, ${esc(marksFilePath ?? "")})
    `;
    await execSql(sql);

    const rows = await execSql("SELECT TOP 1 * FROM Students ORDER BY RollNumber DESC");
    res.json({ success: true, student: rows[0] });
  } catch (e) {
    console.error("INSERT error:", e);
    res.status(500).json({ error: String(e) });
  }
});

// -------------------- Uploads --------------------
function safeName(original) {
  // Keep only safe filename chars
  const base = path.basename(original).replace(/[^a-zA-Z0-9._-]/g, "_");
  return base || `upload_${Date.now()}.xlsx`;
}

const storage = multer.diskStorage({
  destination: (req, _file, cb) => {
    const folder = path.join(__dirname, "Marksheets", String(req.params.rollNumber));
    fs.mkdirSync(folder, { recursive: true });
    cb(null, folder);
  },
  filename: (_req, file, cb) => cb(null, safeName(file.originalname)),
});

const upload = multer({
  storage,
  limits: { fileSize: 5 * 1024 * 1024 }, // 5 MB
  fileFilter: (_req, file, cb) => {
    const ok =
      file.mimetype === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" ||
      file.mimetype === "application/vnd.ms-excel" ||
      file.originalname.toLowerCase().endsWith(".xlsx") ||
      file.originalname.toLowerCase().endsWith(".xls");
    cb(ok ? null : new Error("Only Excel files are allowed"), ok);
  },
});

// Upload marks for a roll number
app.post("/upload/:rollNumber", upload.single("file"), async (req, res) => {
  try {
    const { rollNumber } = req.params;
    if (!req.file) return res.status(400).json({ error: "No file uploaded" });

    const relativeFsPath = path
      .join("Marksheets", String(rollNumber), req.file.originalname)
      .replace(/\\/g, "/");

    const sql = `
      UPDATE Students
      SET MarksFilePath = ${esc(relativeFsPath)}
      WHERE RollNumber = ${Number(rollNumber)}
    `;
    await execSql(sql);

    // Web path the UI can open via the static route
    res.json({
      success: true,
      message: "File uploaded and path saved",
      path: `/marksheets/${rollNumber}/${req.file.originalname}`,
    });
  } catch (e) {
    console.error("UPLOAD error:", e);
    res.status(500).json({ error: "Upload failed" });
  }
});

// Serve uploaded files
app.use("/marksheets", express.static(path.join(__dirname, "Marksheets")));

// 404
app.use((_req, res) => res.status(404).json({ error: "Not found" }));

// -------------------- Start --------------------
const PORT = 5000;
app.listen(PORT, () => {
  console.log(`Server running at http://localhost:${PORT}`);
});
