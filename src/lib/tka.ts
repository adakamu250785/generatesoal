// Panduan gaya soal TKA Kemendikdasmen: menyerupai TKA, tetap mengacu pada mata pelajaran.
const NUMERACY_KEYWORDS = ["matematika", "ipa", "fisika", "kimia", "biologi", "ekonomi", "akuntansi", "informatika", "geografi", "ipas"];

export function isNumeracySubject(subject: string): boolean {
  const s = subject.toLowerCase();
  return NUMERACY_KEYWORDS.some((k) => s.includes(k));
}

export function tkaInstruction(subject: string): string {
  const focus = isNumeracySubject(subject)
    ? "Mapel ini boleh mengintegrasikan literasi dan numerasi: stimulus dapat berupa data, tabel, grafik, atau situasi kuantitatif nyata yang relevan dengan materi."
    : "Mapel ini fokus pada LITERASI saja (JANGAN memaksakan angka/hitungan numerasi): stimulus berupa studi kasus, teks wacana faktual, dilema, berita, atau skenario kehidupan yang relevan dengan materi.";
  return `FORMAT SOAL BERSTANDAR TKA KEMENDIKDASMEN (WAJIB):
- Gaya soal menyerupai TKA, tetapi isi tetap mengacu penuh pada mata pelajaran ${subject} dan materi yang dipilih.
- Setiap soal (atau kelompok soal) diawali stimulus kontekstual yang bermakna dan diisi pada field "stimulus".
- Pertanyaan menuntut siswa memahami/menalar stimulus, bukan hafalan lepas.
- ${focus}
- Opsi pilihan ganda homogen dan masuk akal (pengecoh plausibel).
- Tingkat kesulitan dan level kognitif TETAP mengikuti pengaturan guru (tidak harus C4–C6).`;
}
