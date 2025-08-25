const express = require("express");
const multer = require("multer");
const fs = require("fs");
const mammoth = require("mammoth");
const pdf = require("pdf-parse");
const officegen = require("officegen");
const unzipper = require("unzipper");
const { Document, Packer, Paragraph } = require("docx");
const { PDFDocument, StandardFonts, PDFFont } = require("pdf-lib");
const fontkit = require("fontkit");
const path = require("path");

const app = express();
const upload = multer({ dest: "uploads/" });

app.use(express.static("public"));

// Extract text from PPTX
async function extractTextFromPptx(filePath) {
  const texts = [];
  return new Promise((resolve, reject) => {
    fs.createReadStream(filePath)
      .pipe(unzipper.Parse())
      .on("entry", function (entry) {
        if (
          entry.path.includes("ppt/slides/slide") &&
          entry.path.endsWith(".xml")
        ) {
          let data = "";
          entry.on("data", (chunk) => (data += chunk.toString()));
          entry.on("end", () => {
            const matches = data.match(/<a:t[^>]*>(.*?)<\/a:t>/g) || [];
            matches.forEach((t) => {
              const content = t.replace(/<\/?a:t[^>]*>/g, "");
              texts.push(content);
            });
          });
        } else {
          entry.autodrain();
        }
      })
      .on("close", () => resolve(texts))
      .on("error", reject);
  });
}

app.post("/upload", upload.single("file"), async (req, res) => {
  const filePath = req.file.path;
  const ext = path.extname(req.file.originalname).toLowerCase();
  const baseName = path.basename(req.file.originalname, ext);
  const outputDir = "outputs";
  const uploadDir = "uploads";
  if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir);
  if (!fs.existsSync(uploadDir)) fs.mkdirSync(uploadDir);
  const outputFile = `outputs/${baseName}_UPPER${ext}`;

  try {
    if (ext === ".docx") {
      // Extract & rebuild DOCX
      const result = await mammoth.extractRawText({ path: filePath });
      const upperText = result.value.toUpperCase();

      const doc = new Document({
        sections: [{ properties: {}, children: [new Paragraph(upperText)] }],
      });

      const buffer = await Packer.toBuffer(doc);
      fs.writeFileSync(outputFile, buffer);
    } else if (ext === ".pdf") {
      // Extract & rebuild PDF
      const buffer = fs.readFileSync(filePath);
      const result = await pdf(buffer);
      const upperText = result.text.toUpperCase();

      const pdfDoc = await PDFDocument.create();
      pdfDoc.registerFontkit(fontkit);
      const page = pdfDoc.addPage([600, 800]);
      const fontBytes = fs.readFileSync(path.join(__dirname, "NotoSans.ttf"));
      const font = await pdfDoc.embedFont(fontBytes);

      page.drawText(upperText, { x: 50, y: 750, size: 12, font });

      const pdfBytes = await pdfDoc.save();
      fs.writeFileSync(outputFile, pdfBytes);
    } else if (ext === ".pptx") {
      // Extract & rebuild PPTX
      const texts = await extractTextFromPptx(filePath);
      const upperTexts = texts.map((t) => t.toUpperCase());

      const pptx = officegen("pptx");
      upperTexts.forEach((t) => {
        const slide = pptx.makeNewSlide();
        slide.addText(t, { x: 1, y: 1, font_size: 18 });
      });

      const out = fs.createWriteStream(outputFile);
      pptx.generate(out);
      await new Promise((resolve) => out.on("finish", resolve));
    } else {
      return res.status(400).send("Unsupported file format");
    }

    res.download(outputFile, (err) => {
      if (err) console.error(err);
      fs.unlinkSync(filePath); // cleanup
    });
  } catch (err) {
    console.error(err);
    res.status(500).send("Error processing file");
  }
});

// UI
app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "views", "index.html"));
});

app.listen(3000, () => console.log("Server running on http://localhost:3000"));
