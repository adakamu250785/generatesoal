// Ekspor dokumen Word (.docx) untuk paket asesmen.
// Format mengikuti tata naskah dokumen sekolah Indonesia: kertas A4, font
// Times New Roman, tabel identitas ber-border, tabel kisi-kisi berwarna dengan
// kolom lengkap, naskah soal bernomor dengan indentasi menggantung, kunci
// jawaban ringkas + pembahasan rinci, dan pedoman penskoran — plus footer
// nomor halaman dan header dokumen di halaman kedua dst.

import {
  AlignmentType,
  BorderStyle,
  Document,
  Footer,
  Header,
  PageNumber,
  Packer,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableRow,
  TextRun,
  VerticalAlign,
  WidthType,
} from "docx";
import type { AssessmentIdentity, ExportType, GridRow, Question } from "./types";

const FONT = "Times New Roman";
const C = {
  primaryDark: "312E81",
  primary: "4338CA",
  text: "1E293B",
  muted: "475569",
  green: "047857",
  border: "94A3B8",
  softBorder: "CBD5E1",
  headerFill: "4338CA",
  softFill: "EEF2FF",
  zebraFill: "F8FAFC",
};

type BorderStyleValue = (typeof BorderStyle)[keyof typeof BorderStyle];
type Alignment = (typeof AlignmentType)[keyof typeof AlignmentType];

interface BorderSpec {
  style: BorderStyleValue;
  size: number;
  color: string;
  space?: number;
}

const bCell: BorderSpec = { style: BorderStyle.SINGLE, size: 4, color: C.border };
const bNone: BorderSpec = { style: BorderStyle.NONE, size: 0, color: "auto" };
const bAll = {
  top: { ...bCell },
  bottom: { ...bCell },
  left: { ...bCell },
  right: { ...bCell },
};
const bNoneAll = {
  top: { ...bNone },
  bottom: { ...bNone },
  left: { ...bNone },
  right: { ...bNone },
};
const bTableNone = {
  top: { ...bNone },
  bottom: { ...bNone },
  left: { ...bNone },
  right: { ...bNone },
  insideHorizontal: { ...bNone },
  insideVertical: { ...bNone },
};

type RunOpts = { bold?: boolean | undefined; italics?: boolean | undefined; size?: number | undefined; color?: string | undefined };

const t = (text: string, o: RunOpts = {}) =>
  new TextRun({
    text,
    font: FONT,
    ...(o.bold ? { bold: true } : {}),
    ...(o.italics ? { italics: true } : {}),
    ...(o.size !== undefined ? { size: o.size } : {}),
    ...(o.color !== undefined ? { color: o.color } : {}),
  });

const p = (children: TextRun[], opts: {
  alignment?: Alignment | undefined;
  spacing?: { before?: number; after?: number } | undefined;
  indent?: { left?: number; right?: number; hanging?: number } | undefined;
  pageBreakBefore?: boolean | undefined;
  border?: { bottom?: BorderSpec | undefined } | undefined;
} = {}) => {
  const args = {
    children,
    ...(opts.alignment !== undefined ? { alignment: opts.alignment } : {}),
    ...(opts.spacing !== undefined ? { spacing: opts.spacing } : {}),
    ...(opts.indent !== undefined ? { indent: opts.indent } : {}),
    ...(opts.pageBreakBefore ? { pageBreakBefore: true } : {}),
    ...(opts.border !== undefined ? { border: opts.border } : {}),
  } as ParaArgs;
  return new Paragraph(args);
};

interface CellOpts {
  width?: number | undefined;
  text?: string | undefined;
  bold?: boolean | undefined;
  align?: Alignment | undefined;
  fill?: string | undefined;
  color?: string | undefined;
  size?: number | undefined;
  colSpan?: number | undefined;
  borderless?: boolean | undefined;
}

type CellArgs = ConstructorParameters<typeof TableCell>[0];
type ParaArgs = ConstructorParameters<typeof Paragraph>[0];

function cell(o: CellOpts): TableCell {
  return new TableCell({
    verticalAlign: VerticalAlign.CENTER,
    borders: o.borderless ? bNoneAll : bAll,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: [
      new Paragraph({
        spacing: { before: 20, after: 20 },
        children: [t(o.text ?? "", { bold: o.bold, size: o.size ?? 18, color: o.color })],
        ...(o.align !== undefined ? { alignment: o.align } : {}),
      } satisfies ParaArgs),
    ],
    ...(o.width !== undefined ? { width: { size: o.width, type: WidthType.PERCENTAGE } } : {}),
    ...(o.fill !== undefined ? { shading: { type: ShadingType.CLEAR, color: "auto", fill: o.fill } } : {}),
    ...(o.colSpan !== undefined ? { columnSpan: o.colSpan } : {}),
  } satisfies CellArgs);
}

// ---------- Bagian kepala dokumen ----------

function coverChildren(identity: AssessmentIdentity, subj: string): (Paragraph | Table)[] {
  const rows: string[][] = [
    ["Mata Pelajaran", subj, "Kelas / Semester", `${identity.grade} / ${identity.semester}`],
    ["Jenjang", identity.level, "Tahun Pelajaran", identity.academicYear],
    ["Jenis Asesmen", identity.assessmentType, "Alokasi Waktu", "90 Menit"],
  ];

  const identityTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: bTableNone,
    rows: rows.map(
      (r) =>
        new TableRow({
          children: [
            cell({ width: 20, text: r[0], bold: true, size: 20, borderless: true }),
            cell({ width: 30, text: `: ${r[1]}`, size: 20, borderless: true }),
            cell({ width: 22, text: r[2], bold: true, size: 20, borderless: true }),
            cell({ width: 28, text: `: ${r[3]}`, size: 20, borderless: true }),
          ],
        }),
    ),
  });

  return [
    p([t(identity.schoolName.toUpperCase(), { bold: true, size: 28, color: C.primaryDark })], {
      alignment: AlignmentType.CENTER,
      spacing: { after: 60 },
    }),
    p([t(`NASKAH ASESMEN: ${identity.assessmentType.toUpperCase()}`, { bold: true, size: 24 })], {
      alignment: AlignmentType.CENTER,
      spacing: { after: 40 },
    }),
    p([t(`TAHUN PELAJARAN ${identity.academicYear}`, { italics: true, size: 22, color: C.muted })], {
      alignment: AlignmentType.CENTER,
      spacing: { after: 160 },
    }),
    p([], {
      border: { bottom: { style: BorderStyle.DOUBLE, size: 6, color: C.primary, space: 1 } },
      spacing: { after: 240 },
    }),
    identityTable,
    p([], { spacing: { after: 200 } }),
  ];
}

function sectionHeading(text: string, pageBreakBefore = false): Paragraph {
  return p([t(text, { bold: true, size: 26, color: C.primaryDark })], {
    alignment: AlignmentType.CENTER,
    pageBreakBefore,
    spacing: { before: 240, after: 200 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.softBorder, space: 6 } },
  });
}

// ---------- Kisi-kisi ----------

const GRID_HEADERS = ["No", "Materi", "Indikator Soal", "Level", "Kesulitan", "Bentuk Soal", "Nomor", "Skor"];
const GRID_WIDTHS = [5, 17, 32, 7, 9, 16, 7, 7];

function gridTable(grid: GridRow[]): Table {
  const headerRow = new TableRow({
    tableHeader: true,
    children: GRID_HEADERS.map(
      (h, i) =>
        new TableCell({
          width: { size: GRID_WIDTHS[i] ?? 0, type: WidthType.PERCENTAGE },
          shading: { type: ShadingType.CLEAR, color: "auto", fill: C.headerFill },
          verticalAlign: VerticalAlign.CENTER,
          borders: bAll,
          margins: { top: 60, bottom: 60, left: 80, right: 80 },
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [t(h, { bold: true, size: 18, color: "FFFFFF" })],
            }),
          ],
        }),
    ),
  });

  const bodyRows = grid.map((g, idx) => {
    const fill = idx % 2 ? C.zebraFill : undefined;
    return new TableRow({
      children: [
        cell({ text: String(g.no), width: GRID_WIDTHS[0], align: AlignmentType.CENTER, fill }),
        cell({ text: g.materi || "", width: GRID_WIDTHS[1], fill }),
        cell({ text: g.indikator || "", width: GRID_WIDTHS[2], fill }),
        cell({ text: g.levelKognitif || "", width: GRID_WIDTHS[3], align: AlignmentType.CENTER, fill }),
        cell({ text: g.tingkatKesulitan || "", width: GRID_WIDTHS[4], align: AlignmentType.CENTER, fill }),
        cell({ text: g.bentukSoal || "", width: GRID_WIDTHS[5], fill }),
        cell({ text: String(g.nomorSoal ?? g.no), width: GRID_WIDTHS[6], align: AlignmentType.CENTER, fill }),
        cell({ text: String(g.skor ?? 1), width: GRID_WIDTHS[7], align: AlignmentType.CENTER, fill }),
      ],
    });
  });

  const totalSkor = grid.reduce((s, g) => s + (Number(g.skor) || 0), 0);
  const totalRow = new TableRow({
    children: [
      cell({ colSpan: 7, text: "Total Skor", bold: true, align: AlignmentType.RIGHT, fill: C.softFill }),
      cell({ text: String(totalSkor), bold: true, align: AlignmentType.CENTER, fill: C.softFill }),
    ],
  });

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [headerRow, ...bodyRows, totalRow],
  });
}

// ---------- Naskah soal ----------

function matchTable(q: Question): Table {
  const colW = [6, 47, 47];
  const header = new TableRow({
    tableHeader: true,
    children: [
      cell({ text: "No", width: colW[0], bold: true, align: AlignmentType.CENTER, fill: C.softFill }),
      cell({ text: "Pernyataan", width: colW[1], bold: true, fill: C.softFill }),
      cell({ text: "Pilihan Jawaban", width: colW[2], bold: true, fill: C.softFill }),
    ],
  });
  const rows = (q.leftItems ?? []).map(
    (l, i) =>
      new TableRow({
        children: [
          cell({ text: String(i + 1), width: colW[0], align: AlignmentType.CENTER }),
          cell({ text: l, width: colW[1] }),
          cell({ text: `${String.fromCharCode(65 + i)}. ${q.rightItems?.[i] ?? ""}`, width: colW[2] }),
        ],
      }),
  );
  return new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [header, ...rows] });
}

function questionChildren(q: Question): (Paragraph | Table)[] {
  const out: (Paragraph | Table)[] = [];

  if (q.stimulus) {
    out.push(
      p([t(`Stimulus / Teks Bacaan untuk Soal Nomor ${q.no}:`, { bold: true, italics: true, size: 20 })], {
        spacing: { before: 200, after: 60 },
      }),
      p([t(q.stimulus, { italics: true, size: 20, color: C.muted })], {
        alignment: AlignmentType.JUSTIFIED,
        indent: { left: 567, right: 567 },
        spacing: { after: 120 },
      }),
    );
  }

  out.push(
    p([t(`${q.no}. `, { bold: true }), t(q.questionText)], {
      alignment: AlignmentType.JUSTIFIED,
      spacing: { before: 80, after: 80 },
    }),
  );

  (q.options ?? []).forEach((opt) => {
    out.push(p([t(opt)], { indent: { left: 851, hanging: 425 }, spacing: { after: 40 } }));
  });

  (q.statements ?? []).forEach((st, i) => {
    const text = typeof st === "string" ? st : st?.teks ?? String(st);
    out.push(
      p([t(`(${String.fromCharCode(97 + i)}) `, { bold: true }), t(text)], {
        indent: { left: 720 },
        spacing: { after: 40 },
      }),
    );
  });

  if ((q.leftItems?.length ?? 0) > 0) {
    out.push(
      p([t("Pasangkan pernyataan berikut dengan pilihan jawaban yang tepat:", { italics: true, size: 20 })], {
        spacing: { before: 80, after: 60 },
      }),
      matchTable(q),
    );
  }

  return out;
}

// ---------- Kunci jawaban ----------

const KEY_TABLE_WIDTHS = [10, 40, 50];

function answerSummaryTable(questions: Question[]): Table {
  const header = new TableRow({
    tableHeader: true,
    children: [
      cell({ text: "No", width: KEY_TABLE_WIDTHS[0], bold: true, align: AlignmentType.CENTER, fill: C.headerFill, color: "FFFFFF" }),
      cell({ text: "Bentuk Soal", width: KEY_TABLE_WIDTHS[1], bold: true, fill: C.headerFill, color: "FFFFFF" }),
      cell({ text: "Kunci Jawaban", width: KEY_TABLE_WIDTHS[2], bold: true, fill: C.headerFill, color: "FFFFFF" }),
    ],
  });
  const rows = questions.map((q, idx) => {
    const key = String(q.answerKey ?? "");
    const short = key.length > 90 ? `${key.slice(0, 90)}…` : key;
    const fill = idx % 2 ? C.zebraFill : undefined;
    return new TableRow({
      children: [
        cell({ text: String(q.no), width: KEY_TABLE_WIDTHS[0], align: AlignmentType.CENTER, fill }),
        cell({ text: q.bentukSoal, width: KEY_TABLE_WIDTHS[1], fill }),
        cell({ text: short, width: KEY_TABLE_WIDTHS[2], bold: true, color: C.green, fill }),
      ],
    });
  });
  return new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [header, ...rows] });
}

function answerDetailChildren(q: Question): Paragraph[] {
  return [
    p([t(`Soal Nomor ${q.no} (${q.bentukSoal}) — Kunci: `, { bold: true }), t(String(q.answerKey), { bold: true, color: C.green })], {
      spacing: { before: 180, after: 40 },
    }),
    p([t("Pembahasan: ", { bold: true, italics: true }), t(q.explanation || "Tidak ada pembahasan.")], {
      alignment: AlignmentType.JUSTIFIED,
      indent: { left: 425 },
      spacing: { after: 80 },
    }),
  ];
}

// ---------- Rubrik ----------

function rubricChildren(q: Question): Paragraph[] {
  return [
    p([t(`Soal Nomor ${q.no} (${q.bentukSoal}) — Skor Maksimal: ${q.score}`, { bold: true })], {
      spacing: { before: 180, after: 40 },
    }),
    p([t(q.rubrik || "Setiap poin kriteria jawaban yang benar diberi skor proporsional.")], {
      alignment: AlignmentType.JUSTIFIED,
      indent: { left: 425 },
    }),
  ];
}

// ---------- Rakitan dokumen ----------

export function buildDocx(
  identity: AssessmentIdentity,
  gridData: GridRow[],
  questionsData: Question[],
  exportType: ExportType,
): Document {
  const children: (Paragraph | Table)[] = [...coverChildren(identity, identity.subject)];
  let firstPart = true;

  const addPart = (heading: string, part: (Paragraph | Table)[]) => {
    children.push(sectionHeading(heading, exportType === "lengkap" ? !firstPart : false));
    children.push(...part);
    firstPart = false;
  };

  if (exportType === "kisi" || exportType === "lengkap") {
    addPart("KISI-KISI PENULISAN SOAL", [
      gridTable(gridData),
      p([], { spacing: { after: 160 } }),
    ]);
  }

  if (exportType === "soal" || exportType === "lengkap") {
    const part: (Paragraph | Table)[] = [];
    questionsData.forEach((q) => part.push(...questionChildren(q)));
    const totalSkor = questionsData.reduce((s, q) => s + (Number(q.score) || 0), 0);
    part.push(
      p([t(`Total Skor Seluruh Soal: ${totalSkor}`, { bold: true })], {
        spacing: { before: 240, after: 80 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.softBorder, space: 6 } },
      }),
    );
    addPart("NASKAH SOAL ASESMEN", part);
  }

  if (exportType === "kunci" || exportType === "lengkap") {
    addPart("KUNCI JAWABAN DAN PEMBAHASAN", [
      answerSummaryTable(questionsData),
      p([t("Pembahasan Rinci:", { bold: true, size: 22 })], { spacing: { before: 240, after: 60 } }),
      ...questionsData.flatMap(answerDetailChildren),
    ]);
  }

  if (exportType === "rubrik" || exportType === "lengkap") {
    addPart(
      "PEDOMAN PENSKORAN & RUBRIK",
      questionsData
        .filter((q) => q.rubrik || q.bentukSoal === "Uraian")
        .flatMap(rubricChildren),
    );
  }

  const pageHeader = new Header({
    children: [
      new Paragraph({
        alignment: AlignmentType.RIGHT,
        border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.softBorder, space: 4 } },
        children: [
          t(`${identity.schoolName} — ${identity.assessmentType}`, { size: 16, color: C.muted, italics: true }),
        ],
      }),
    ],
  });

  const emptyHeader = new Header({ children: [p([])] });

  const pageFooter = () =>
    new Footer({
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          border: { top: { style: BorderStyle.SINGLE, size: 4, color: C.softBorder, space: 4 } },
          children: [
            new TextRun({ children: ["Halaman ", PageNumber.CURRENT], font: FONT, size: 16, color: C.muted }),
          ],
        }),
      ],
    });

  return new Document({
    styles: {
      default: {
        document: {
          run: { font: FONT, size: 22, color: C.text },
        },
      },
    },
    sections: [
      {
        properties: {
          titlePage: true,
          page: {
            size: { width: 11906, height: 16838 }, // A4
            margin: { top: 1134, bottom: 1134, left: 1418, right: 1134 },
          },
        },
        headers: { default: pageHeader, first: emptyHeader },
        footers: { default: pageFooter(), first: pageFooter() },
        children,
      },
    ],
  });
}

export async function exportAssessmentDocx(
  identity: AssessmentIdentity,
  gridData: GridRow[],
  questionsData: Question[],
  exportType: ExportType,
): Promise<string> {
  const doc = buildDocx(identity, gridData, questionsData, exportType);
  const blob = await Packer.toBlob(doc);
  const filename = `Asesmen_${identity.subject.replace(/\s+/g, "_")}_Kelas${identity.grade}_${exportType.toUpperCase()}.docx`;
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.click();
  setTimeout(() => URL.revokeObjectURL(url), 5000);
  return filename;
}
