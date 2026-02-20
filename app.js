const fileInput = document.getElementById("fileInput");
const fileMeta = document.getElementById("fileMeta");
const resumeText = document.getElementById("resumeText");
const nameLine = document.getElementById("nameLine");
const headingKnown = document.getElementById("headingKnown");
const headingAllCaps = document.getElementById("headingAllCaps");
const headingShort = document.getElementById("headingShort");
const headingCustom = document.getElementById("headingCustom");
const stripBulletMarkers = document.getElementById("stripBulletMarkers");
const alignDates = document.getElementById("alignDates");
const useAi = document.getElementById("useAi");
const aiModel = document.getElementById("aiModel");
const generateBtn = document.getElementById("generateBtn");
const clearBtn = document.getElementById("clearBtn");
const status = document.getElementById("status");

const commonHeadings = [
  "Professional Summary",
  "Technical Skills",
  "Education And Training",
  "Education",
  "Certifications",
  "Professional Development",
  "Professional Experience",
  "Work Experience",
  "Experience",
  "Skills",
  "Projects",
  "Summary"
];

fileInput.addEventListener("change", async (event) => {
  const file = event.target.files[0];
  if (!file) {
    return;
  }
  fileMeta.textContent = `${file.name} (${Math.round(file.size / 1024)} KB)`;

  if (file.name.toLowerCase().endsWith(".txt")) {
    const text = await file.text();
    resumeText.value = text;
    status.textContent = "Text loaded.";
    return;
  }

  if (file.name.toLowerCase().endsWith(".docx")) {
    const buffer = await file.arrayBuffer();
    try {
      const result = await window.mammoth.extractRawText({ arrayBuffer: buffer });
      resumeText.value = result.value || "";
      status.textContent = "Docx text loaded.";
    } catch (err) {
      status.textContent = "Failed to read docx file.";
    }
    return;
  }

  status.textContent = "Unsupported file type. Use .txt or .docx.";
});

clearBtn.addEventListener("click", () => {
  resumeText.value = "";
  fileInput.value = "";
  fileMeta.textContent = "No file selected";
  status.textContent = "";
});

generateBtn.addEventListener("click", async () => {
  let text = resumeText.value;
  if (!text.trim()) {
    status.textContent = "Add resume content first.";
    return;
  }

  if (useAi.checked) {
    status.textContent = "Running local AI (Ollama)...";
    try {
      text = await formatWithAi(text);
    } catch (err) {
      status.textContent = "AI formatting failed. Check Ollama is running.";
      return;
    }
  }

  const doc = buildDocument(text);
  status.textContent = "Generating .docx...";

  try {
    const blob = await window.docx.Packer.toBlob(doc);
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = "formatted-resume.docx";
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
    status.textContent = "Download ready.";
  } catch (err) {
    status.textContent = "Failed to generate .docx.";
  }
});

async function formatWithAi(content) {
  const prompt = getFormattingPrompt();
  const response = await fetch("http://localhost:11434/api/generate", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: aiModel.value || "llama3.1:8b",
      prompt: content,
      system: prompt,
      stream: false
    })
  });

  if (!response.ok) {
    throw new Error("Ollama request failed");
  }

  const data = await response.json();
  return (data.response || "").trim();
}

function getFormattingPrompt() {
  return `
You are a formatting engine. Your task is to reformat the provided resume content into a clean, professional, client-ready Word document.

CRITICAL RULES (Must Follow Strictly)
1. DO NOT change the content in any way
- Do not remove text
- Do not add text
- Do not reword, summarize, or paraphrase
- Do not correct grammar or spelling
- Do not reorder bullets or sections
- Do not infer or insert missing information
2. Formatting ONLY
- Your responsibility is limited to layout, spacing, alignment, fonts, and structure
- The words must remain exactly as provided
3. Output must be suitable for Microsoft Word
- Assume standard Word rendering
- No Markdown symbols in the final output
- No emojis, icons, tables, or graphics
- No columns

Page Setup
- Page size: A4
- Orientation: Portrait
- Margins: Top 1 inch, Bottom 1 inch, Left 1 inch, Right 1 inch
- Line spacing: Single
- Paragraph spacing: Before 0 pt, After 6 pt (unless specified otherwise)

Font Standards (Apply Consistently)
- Primary font: Calibri
- Text color: Black
- Body text size: 10.5 pt
- Section headers: 12 pt, Bold
- Candidate name: 14 pt, Bold

Header (Candidate Name)
- Candidate name appears at the very top
- Alignment: Center
- Font: Calibri, 14 pt, Bold
- No underline
- No extra text before or after the name
- Add one blank line after the name

Section Headings
- Font: Calibri
- Size: 12 pt
- Bold
- Left-aligned
- Capitalization: Title Case (as provided — do not modify text)
- Spacing: One blank line before the section header
- No blank line between header and its content

Body Text (Paragraphs)
- Font: Calibri
- Size: 10.5 pt
- Left-aligned
- Single-spaced
- Paragraph spacing after: 6 pt
- No indentation
- Maintain original paragraph breaks exactly as provided

Bullet Points
- Bullet style: Standard round bullet (Word default)
- Font: Calibri, 10.5 pt
- Alignment: Left
- Indentation: Bullet indent 0.25 inch, Text indent 0.5 inch
- Spacing: No blank line between bullets, 6 pt after final bullet in a group
- Preserve bullet text exactly as provided

Job Experience Entries
- Company line: Company Name, Location Start Date – End Date
- Company name bold, dates regular
- Dates right-aligned on same line
- Role title on next line, bold
- Preserve bullets/paragraphs as received

Skills Sections
- Keep category labels exactly as provided
- Category labels bold, same line as content
- Content regular, comma-separated lists unchanged
- No tables or columns

Certifications & Education
- Each item on its own line
- Maintain order
- No bullets unless present in source
- Preserve dates and separators exactly

Final Output Requirements
- Clean, consistent, recruiter-standard resume
- No decorative elements or horizontal lines
- No headers or footers
- No page numbers
- No commentary or analysis

Return only the formatted resume content, ready to be pasted directly into Microsoft Word.
`.trim();
}

function buildDocument(rawText) {
  const {
    Document,
    Paragraph,
    TextRun,
    AlignmentType,
    TabStopType
  } = window.docx;

  const lines = rawText.replace(/\r\n/g, "\n").split("\n");
  const trimmedLines = lines.map((line) => line.replace(/\s+$/g, ""));
  const contentLines = trimmedLines;

  const nameLineNumber = Math.max(parseInt(nameLine.value || "1", 10), 1);
  const nameIndex = findNonEmptyLineIndex(contentLines, nameLineNumber);

  const customHeadingList = headingCustom.value
    .split("|")
    .map((entry) => entry.trim())
    .filter(Boolean);

  let currentSection = "";

  const paragraphs = [];

  for (let i = 0; i < contentLines.length; i += 1) {
    const line = contentLines[i];
    const isEmpty = !line.trim();

    if (isEmpty) {
      paragraphs.push(blankParagraph(Paragraph));
      continue;
    }

    if (i === nameIndex) {
      paragraphs.push(
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 0, line: 240 },
          children: [
            new TextRun({
              text: line,
              bold: true,
              size: 28,
              font: "Calibri",
              color: "000000"
            })
          ]
        })
      );
      paragraphs.push(blankParagraph(Paragraph));
      continue;
    }

    if (isHeading(line, i, contentLines, customHeadingList)) {
      currentSection = line.trim();
      paragraphs.push(blankParagraph(Paragraph));
      paragraphs.push(
        new Paragraph({
          alignment: AlignmentType.LEFT,
          spacing: { after: 0, line: 240 },
          children: [
            new TextRun({
              text: line,
              bold: true,
              size: 24,
              font: "Calibri",
              color: "000000"
            })
          ]
        })
      );
      continue;
    }

    const bulletInfo = detectBullet(line);
    const nextLine = contentLines[i + 1] || "";
    const nextIsBullet = detectBullet(nextLine).isBullet;
    const afterSpacing = nextIsBullet ? 0 : 120;

    if (bulletInfo.isBullet && stripBulletMarkers.checked) {
      paragraphs.push(
        new Paragraph({
          alignment: AlignmentType.LEFT,
          spacing: { after: afterSpacing, line: 240 },
          bullet: { level: 0 },
          indent: { left: 720, hanging: 360 },
          children: [
            new TextRun({
              text: bulletInfo.text,
              size: 21,
              font: "Calibri",
              color: "000000"
            })
          ]
        })
      );
      continue;
    }

    if (
      alignDates.checked &&
      currentSection.toLowerCase() === "professional experience"
    ) {
      const splitLine = splitCompanyLine(line);
      if (splitLine) {
        paragraphs.push(
          new Paragraph({
            alignment: AlignmentType.LEFT,
            spacing: { after: 120, line: 240 },
            tabStops: [
              {
                type: TabStopType.RIGHT,
                position: 9026
              }
            ],
            children: [
              new TextRun({
                text: splitLine.company,
                bold: true,
                size: 21,
                font: "Calibri",
                color: "000000"
              }),
              new TextRun({
                text: "\t" + splitLine.dates,
                size: 21,
                font: "Calibri",
                color: "000000"
              })
            ]
          })
        );
        continue;
      }
    }

    paragraphs.push(
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { after: 120, line: 240 },
        children: [
          new TextRun({
            text: line,
            size: 21,
            font: "Calibri",
            color: "000000"
          })
        ]
      })
    );
  }

  return new Document({
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: 1440,
              right: 1440,
              bottom: 1440,
              left: 1440
            },
            size: {
              width: 11906,
              height: 16838,
              orientation: "portrait"
            }
          }
        },
        children: paragraphs
      }
    ]
  });
}

function findNonEmptyLineIndex(lines, lineNumber) {
  let count = 0;
  for (let i = 0; i < lines.length; i += 1) {
    if (lines[i].trim()) {
      count += 1;
      if (count === lineNumber) {
        return i;
      }
    }
  }
  return 0;
}

function isHeading(line, index, lines, customList) {
  const trimmed = line.trim();
  if (!trimmed) {
    return false;
  }

  const normalized = trimmed.toLowerCase();

  if (customList.some((entry) => entry.toLowerCase() === normalized)) {
    return true;
  }

  if (headingKnown.checked) {
    const match = commonHeadings.some(
      (heading) => heading.toLowerCase() === normalized
    );
    if (match) {
      return true;
    }
  }

  if (headingAllCaps.checked) {
    const noLetters = !/[a-z]/.test(trimmed);
    const hasLetters = /[A-Z]/.test(trimmed);
    if (noLetters && hasLetters && trimmed.length <= 40) {
      return true;
    }
  }

  if (trimmed.endsWith(":") && trimmed.length <= 40) {
    return true;
  }

  if (headingShort.checked) {
    const prev = lines[index - 1] || "";
    const next = lines[index + 1] || "";
    const isStandalone = !prev.trim() && !next.trim();
    if (trimmed.length <= 40 && isStandalone) {
      return true;
    }
  }

  return false;
}

function detectBullet(line) {
  const trimmed = line.trim();
  if (!trimmed) {
    return { isBullet: false, text: trimmed };
  }
  const markers = ["•", "-", "*", "·"];
  for (const marker of markers) {
    if (trimmed.startsWith(marker + " ")) {
      return { isBullet: true, text: trimmed.slice(2) };
    }
    if (trimmed.startsWith(marker)) {
      return { isBullet: true, text: trimmed.slice(1).trimStart() };
    }
  }
  return { isBullet: false, text: trimmed };
}

function splitCompanyLine(line) {
  const separators = [" – ", " - "];
  for (const sep of separators) {
    const idx = line.lastIndexOf(sep);
    if (idx > 0 && idx < line.length - sep.length) {
      const company = line.slice(0, idx).trimEnd();
      const dates = line.slice(idx + 1).trimStart();
      if (/\d/.test(dates)) {
        return { company, dates };
      }
    }
  }
  return null;
}

function blankParagraph(Paragraph) {
  return new Paragraph({
    children: [new window.docx.TextRun({ text: "" })],
    spacing: { after: 120, line: 240 }
  });
}
