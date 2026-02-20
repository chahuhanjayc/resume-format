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
  const text = resumeText.value;
  if (!text.trim()) {
    status.textContent = "Add resume content first.";
    return;
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
