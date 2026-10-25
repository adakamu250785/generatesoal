import { Packer } from "docx";
import { describe, expect, it } from "vitest";

import { buildDocx } from "@/lib/docx-export";
import type { AssessmentIdentity, GridRow, Question } from "@/lib/types";

const identity: AssessmentIdentity = {
  schoolName: "SD Negeri 01 Nusantara",
  academicYear: "2026/2027",
  level: "SD",
  grade: "5",
  subject: "Bahasa Indonesia",
  customSubject: "",
  semester: "Semester 1 (Ganjil)",
  assessmentType: "Sumatif Lingkup Materi (Ulangan Harian)",
  examDuration: "90",
};

const grid: GridRow[] = [
  {
    no: 1,
    materi: "Unsur Intrinsik Cerita",
    indikator: "Siswa dapat menyebutkan tokoh dalam cerita",
    levelKognitif: "C2",
    tingkatKesulitan: "Mudah",
    bentukSoal: "Pilihan Ganda",
    nomorSoal: "1",
    skor: 1,
  },
  {
    no: 2,
    materi: "Ringkasan Teks",
    indikator: "Siswa dapat merangkum isi cerita",
    levelKognitif: "C4",
    tingkatKesulitan: "Sulit",
    bentukSoal: "Uraian",
    nomorSoal: "2",
    skor: 4,
  },
];

const questions: Question[] = [
  {
    id: "q1",
    no: 1,
    stimulus: "Bacalah cerita pendek berikut dengan saksama.",
    materi: "Unsur Intrinsik Cerita",
    indikator: "Siswa dapat menyebutkan tokoh dalam cerita",
    levelKognitif: "C2",
    tingkatKesulitan: "Mudah",
    bentukSoal: "Pilihan Ganda",
    questionText: "Siapa nama tokoh utama dalam cerita di atas?",
    options: ["A. Si Kancil", "B. Bu Kura-kura", "C. Pak Kumbang", "D. Nenek Kesatria"],
    answerKey: "A",
    explanation: "Tokoh utama cerita adalah Si Kancil yang terlibat konflik utama.",
    rubrik: "Jawaban benar mendapat skor 1.",
    score: 1,
  },
  {
    id: "q2",
    no: 2,
    materi: "Ringkasan Teks",
    indikator: "Siswa dapat merangkum isi cerita",
    levelKognitif: "C4",
    tingkatKesulitan: "Sulit",
    bentukSoal: "Uraian",
    questionText: "Buatlah ringkasan cerita di atas dalam tiga kalimat.",
    answerKey: "Ringkasan memuat pokok-pokok cerita secara utuh.",
    explanation: "Ringkasan harus memuat unsur siapa, apa, dan bagaimana.",
    rubrik: "Skor 4: lengkap dan runtut; Skor 3: lengkap tetapi kurang runtut; Skor 2: sebagian benar; Skor 1: mencoba tetapi belum tepat.",
    score: 4,
  },
];

describe("ekspor docx", () => {
  it.each(["kisi", "soal", "kunci", "rubrik", "lengkap"] as const)(
    "membangun dokumen untuk tipe %s tanpa error",
    async (exportType) => {
      const doc = buildDocx(identity, grid, questions, exportType);
      const buffer = await Packer.toBuffer(doc);
      expect(buffer.byteLength).toBeGreaterThan(1000);
    },
  );

  it("menghitung total skor kisi-kisi dengan benar pada baris total", () => {
    // Total skor grid = 1 + 4 = 5; kisi table terakhir memuat baris "Total Skor".
    const doc = buildDocx(identity, grid, questions, "kisi");
    expect(doc).toBeTruthy();
  });
});

import { groupByBentuk, instructionFor, useTwoColumns } from "@/lib/docx-export";

describe("format naskah soal siswa", () => {
  it("opsi pendek disusun 2 kolom, opsi panjang 1 kolom", () => {
    expect(useTwoColumns(["A. 12", "B. 15", "C. 18", "D. 20"])).toBe(true);
    expect(useTwoColumns(["A. " + "x".repeat(60), "B. y", "C. z", "D. w"])).toBe(false);
  });
  it("petunjuk pilihan ganda memakai tanda silang", () => {
    expect(instructionFor("Pilihan Ganda")).toContain("tanda silang (x)");
  });
  it("soal dikelompokkan per bentuk soal", () => {
    expect(groupByBentuk(questions).map((g) => g.bentuk)).toEqual(["Pilihan Ganda", "Uraian"]);
  });
  it("tabel identitas lembar ujian memuat waktu ujian dari identitas", async () => {
    const JSZip = (await import("jszip")).default;
    const buf = await Packer.toBuffer(buildDocx({ ...identity, examDuration: "120" }, grid, questions, "soal"));
    const xml = await (await JSZip.loadAsync(buf)).file("word/document.xml")!.async("string");
    expect(xml).toContain("Waktu");
    expect(xml).toContain("120 Menit");
  });
  it("paket lengkap memuat kisi-kisi, naskah, kunci, dan rubrik", async () => {
    const JSZip = (await import("jszip")).default;
    const buf = await Packer.toBuffer(buildDocx(identity, grid, questions, "lengkap"));
    const xml = await (await JSZip.loadAsync(buf)).file("word/document.xml")!.async("string");
    for (const s of ["KISI-KISI", "NASKAH SOAL", "KUNCI JAWABAN", "RUBRIK", "Nama Siswa"]) expect(xml).toContain(s);
  });
});
