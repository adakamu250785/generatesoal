export interface AssessmentIdentity {
  schoolName: string;
  academicYear: string;
  level: string;
  grade: string;
  subject: string;
  customSubject: string;
  semester: string;
  assessmentType: string;
}

export interface GridRow {
  no: number;
  materi: string;
  indikator: string;
  levelKognitif: string;
  tingkatKesulitan: string;
  bentukSoal: string;
  nomorSoal?: string;
  skor: number;
}

export interface StatementItem {
  teks: string;
  jawaban: string;
}

export interface Question {
  id: string;
  no: number;
  stimulus?: string;
  materi?: string;
  indikator?: string;
  levelKognitif?: string;
  tingkatKesulitan?: string;
  bentukSoal: string;
  questionText: string;
  options?: string[];
  statements?: (string | StatementItem)[];
  leftItems?: string[];
  rightItems?: string[];
  matchingPairs?: unknown;
  answerKey: string;
  explanation?: string;
  rubrik?: string;
  score: number;
}

export interface QualityFinding {
  no: number;
  status: string;
  note: string;
}

export interface QualityReport {
  overallStatus: string;
  score: number;
  findings: QualityFinding[];
}

export interface AssessmentPackage {
  id: string;
  title: string;
  createdAt: string;
  identity: AssessmentIdentity;
  materials: string[];
  gridData: GridRow[];
  questionsData: Question[];
  qualityReport?: QualityReport | null;
}

export interface QuestionTypeConfig {
  enabled: boolean;
  count: number;
  weight: number;
}

export type QuestionTypes = Record<
  "pg" | "pgk" | "isian" | "uraian" | "bs" | "jodohkan",
  QuestionTypeConfig
>;

export type ExportType = "kisi" | "soal" | "kunci" | "rubrik" | "lengkap";

export const DEFAULT_IDENTITY: AssessmentIdentity = {
  schoolName: "SD/SMP/SMA Nusantara",
  academicYear: "2026/2027",
  level: "",
  grade: "",
  subject: "",
  customSubject: "",
  semester: "Semester 1 (Ganjil)",
  assessmentType: "Sumatif Lingkup Materi (Ulangan Harian)",
};

export const DEFAULT_QUESTION_TYPES: QuestionTypes = {
  pg: { enabled: true, count: 10, weight: 1 },
  pgk: { enabled: false, count: 5, weight: 2 },
  isian: { enabled: false, count: 5, weight: 2 },
  uraian: { enabled: true, count: 5, weight: 4 },
  bs: { enabled: false, count: 5, weight: 1 },
  jodohkan: { enabled: false, count: 5, weight: 2 },
};
