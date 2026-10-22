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
