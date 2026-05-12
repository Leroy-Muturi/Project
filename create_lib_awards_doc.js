const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
  LevelFormat, PageNumber, PageBreak, UnderlineType
} = require('docx');
const fs = require('fs');

const BRAND_BLUE = "1F4E79";
const LIGHT_BLUE = "D6E4F0";
const ACCENT = "2E75B6";
const GOLD = "B8860B";

const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };

function h(text, level, color) {
  const sizes = { 1: 36, 2: 28, 3: 24 };
  const colors = { 1: BRAND_BLUE, 2: ACCENT, 3: "333333" };
  return new Paragraph({
    heading: level === 1 ? HeadingLevel.HEADING_1 : level === 2 ? HeadingLevel.HEADING_2 : HeadingLevel.HEADING_3,
    spacing: { before: level === 1 ? 360 : 240, after: 120 },
    children: [new TextRun({ text, bold: true, size: sizes[level], color: color || colors[level], font: "Arial" })]
  });
}

function p(text, opts = {}) {
  return new Paragraph({
    spacing: { before: 80, after: 100 },
    children: [new TextRun({ text, size: 22, font: "Arial", bold: opts.bold, italics: opts.italic, color: opts.color || "000000" })]
  });
}

function qLabel(text, marks) {
  return new Paragraph({
    spacing: { before: 160, after: 60 },
    children: [
      new TextRun({ text: text + " ", bold: true, size: 22, font: "Arial", color: ACCENT }),
      new TextRun({ text: `[${marks} marks]`, bold: true, size: 20, font: "Arial", color: GOLD })
    ]
  });
}

function answer(text) {
  return new Paragraph({
    spacing: { before: 60, after: 120 },
    indent: { left: 360 },
    children: [new TextRun({ text, size: 22, font: "Arial", color: "1a1a1a" })]
  });
}

function dataFlag(text) {
  return new Paragraph({
    spacing: { before: 40, after: 80 },
    indent: { left: 360 },
    children: [new TextRun({ text: `[INSERT INTERNAL DATA: ${text}]`, size: 20, font: "Arial", color: "CC0000", bold: true, italics: true })]
  });
}

function bullet(text, opts = {}) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { before: 40, after: 40 },
    children: [new TextRun({ text, size: 22, font: "Arial", bold: opts.bold, color: opts.color || "1a1a1a" })]
  });
}

function sectionBox(categoryNum, categoryTitle) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    rows: [new TableRow({
      children: [new TableCell({
        borders,
        shading: { fill: BRAND_BLUE, type: ShadingType.CLEAR },
        margins: { top: 160, bottom: 160, left: 240, right: 240 },
        width: { size: 9360, type: WidthType.DXA },
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ text: `CATEGORY ${categoryNum}: ${categoryTitle}`, bold: true, size: 28, font: "Arial", color: "FFFFFF" })
          ]
        })]
      })]
    })]
  });
}

function divider() {
  return new Paragraph({
    spacing: { before: 120, after: 120 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: ACCENT, space: 1 } },
    children: []
  });
}

function gap() {
  return new Paragraph({ spacing: { before: 60, after: 60 }, children: [new TextRun({ text: " " })] });
}

const doc = new Document({
  numbering: {
    config: [{
      reference: "bullets",
      levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } } }]
    }]
  },
  styles: {
    default: { document: { run: { font: "Arial", size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 36, bold: true, font: "Arial", color: BRAND_BLUE },
        paragraph: { spacing: { before: 360, after: 120 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: "Arial", color: ACCENT },
        paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, font: "Arial", color: "333333" },
        paragraph: { spacing: { before: 200, after: 80 }, outlineLevel: 2 } },
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    children: [

      // ===================== COVER PAGE =====================
      gap(), gap(), gap(),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 480, after: 120 },
        children: [new TextRun({ text: "16TH ANNUAL INSURANCE AWARDS 2026", bold: true, size: 44, font: "Arial", color: BRAND_BLUE })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 60, after: 120 },
        children: [new TextRun({ text: "Risk, Resilience and Renewal: Innovating Insurance for Tomorrow's Kenya", size: 28, font: "Arial", color: ACCENT, italics: true })]
      }),
      divider(),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 480, after: 120 },
        children: [new TextRun({ text: "AWARD SUBMISSION", bold: true, size: 56, font: "Arial", color: BRAND_BLUE })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 120, after: 120 },
        children: [new TextRun({ text: "PART C: INSURANCE BROKERS", bold: true, size: 32, font: "Arial", color: ACCENT })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 120, after: 480 },
        children: [new TextRun({ text: "Categories 19 – 25", bold: true, size: 28, font: "Arial", color: "555555" })]
      }),
      gap(), gap(),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [2880, 6480],
        rows: [
          new TableRow({ children: [
            new TableCell({ borders, shading: { fill: LIGHT_BLUE, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, width: { size: 2880, type: WidthType.DXA },
              children: [new Paragraph({ children: [new TextRun({ text: "Submitted by:", bold: true, size: 22, font: "Arial" })] })] }),
            new TableCell({ borders, margins: { top: 80, bottom: 80, left: 120, right: 120 }, width: { size: 6480, type: WidthType.DXA },
              children: [new Paragraph({ children: [new TextRun({ text: "Laser Insurance Brokers Limited (LIB)", size: 22, font: "Arial", bold: true })] })] })
          ]}),
          // ... content intentionally shortened for repository sample ...
      ]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "Laser Insurance Brokers Limited | Innovating Insurance for Tomorrow's Kenya", size: 20, font: "Arial", italics: true, color: "555555" })]
      }),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('/mnt/user-data/outputs/LIB_Insurance_Awards_2026_Submission.docx', buffer);
  console.log('Done');
});
