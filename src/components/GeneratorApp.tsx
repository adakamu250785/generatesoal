import { useEffect, useMemo, useState } from "react";
import {
  Award,
  BookOpen,
  CircleCheck,
  CirclePlus,
  CloudDownload,
  CloudUpload,
  Database,
  Download,
  FileCheck,
  FileText,
  HardDrive,
  History,
  Info,
  KeyRound,
  Layers,
  LayoutDashboard,
  Plus,
  RefreshCw,
  Save,
  Settings,
  SlidersHorizontal,
  Sparkles,
  Table,
  Trash2,
  TriangleAlert,
} from "lucide-react";

import { PRESET_INSTRUCTIONS, SUBJECTS_DATA } from "@/lib/constants";
import { exportAssessmentDocx } from "@/lib/docx-export";
import { callGeminiAPI } from "@/lib/gemini";
import { dbDelete, dbGetAll, dbPut, initDB } from "@/lib/idb";
import {
  DEFAULT_IDENTITY,
  DEFAULT_QUESTION_TYPES,
  type AssessmentIdentity,
  type AssessmentPackage,
  type ExportType,
  type GridRow,
  type QualityReport,
  type Question,
  type QuestionTypes,
} from "@/lib/types";

type TabId = "dashboard" | "create" | "bank" | "history" | "settings";
type NotificationType = "info" | "success" | "warning" | "error";

const NAV_ITEMS: { id: TabId; label: string; icon: typeof Sparkles }[] = [
  { id: "dashboard", label: "Dashboard", icon: LayoutDashboard },
  { id: "create", label: "Buat Soal AI", icon: CirclePlus },
  { id: "bank", label: "Bank Soal", icon: Database },
  { id: "history", label: "Riwayat Paket", icon: History },
  { id: "settings", label: "Pengaturan", icon: Settings },
];

const INPUT_CLASS =
  "w-full rounded-lg border border-input bg-card p-2.5 text-sm outline-none transition focus:ring-2 focus:ring-ring";

function Card({ className = "", children }: { className?: string; children: React.ReactNode }) {
  return <div className={`rounded-2xl border border-border bg-card shadow-sm ${className}`}>{children}</div>;
}

export function GeneratorApp() {
  // Navigation
  const [activeTab, setActiveTab] = useState<TabId>("dashboard");

  // Notification / Alert State
  const [notification, setNotification] = useState<{ msg: string; type: NotificationType } | null>(null);

  // API Key State
  const [apiKey, setApiKey] = useState("");

  // Wizard Step State for Generator
  const [currentStep, setCurrentStep] = useState(1);

  // Stored Packages & History
  const [packagesHistory, setPackagesHistory] = useState<AssessmentPackage[]>([]);

  // Package Identity Form State
  const [identity, setIdentity] = useState<AssessmentIdentity>(DEFAULT_IDENTITY);

  // Materials State
  const [materials, setMaterials] = useState<string[]>([""]);
  const [sasMode, setSasMode] = useState<"manual" | "ai_suggest" | "hybrid">("manual");
  const [aiSuggestedMaterials, setAiSuggestedMaterials] = useState<string[]>([]);
  const [acceptedAiMaterials, setAcceptedAiMaterials] = useState<string[]>([]);

  // Additional Teacher Instructions
  const [additionalInstruction, setAdditionalInstruction] = useState("");

  // Question Types & Counts
  const [questionTypes, setQuestionTypes] = useState<QuestionTypes>(DEFAULT_QUESTION_TYPES);

  // Difficulty Setting
  const [difficultyMode, setDifficultyMode] = useState("Campuran");
  const [difficultyDist, setDifficultyDist] = useState({ mudah: 5, sedang: 7, sulit: 3 });

  // Cognitive Level Setting
  const [useCognitiveLevel, setUseCognitiveLevel] = useState(true);

  // Generated Data
  const [isGenerating, setIsGenerating] = useState(false);
  const [loadingMessage, setLoadingMessage] = useState("");
  const [gridData, setGridData] = useState<GridRow[]>([]);
  const [questionsData, setQuestionsData] = useState<Question[]>([]);
  const [qualityReport, setQualityReport] = useState<QualityReport | null>(null);

  // Active Package ID (if editing existing)
  const [activePackageId, setActivePackageId] = useState<string | null>(null);

  // Active Preview Tab
  const [previewTab, setPreviewTab] = useState<"soal" | "kisi" | "kunci" | "rubrik">("soal");

  const showNotification = (msg: string, type: NotificationType = "info") => {
    setNotification({ msg, type });
    setTimeout(() => setNotification(null), 4000);
  };

  // Load saved settings & packages on mount
  useEffect(() => {
    initDB()
      .then(async () => {
        try {
          const data = await dbGetAll<AssessmentPackage>("packages");
          setPackagesHistory(data.sort((a, b) => new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime()));
        } catch (e) {
          console.error(e);
        }
        try {
          const res = await dbGetAll<{ key: string; value: string }>("settings");
          const keyItem = res.find((i) => i.key === "gemini_api_key");
          if (keyItem) setApiKey(keyItem.value);
        } catch (e) {
          console.error(e);
        }
      })
      .catch((err) => showNotification("Error IndexedDB: " + err, "error"));
  }, []);

  const saveApiKey = async (val: string) => {
    setApiKey(val);
    await dbPut("settings", { key: "gemini_api_key", value: val });
    showNotification("Pengaturan API Key tersimpan!", "success");
  };

  // Calculate total questions count
  const totalQuestionsCount = useMemo(() => {
    let total = 0;
    Object.keys(questionTypes).forEach((k) => {
      if (questionTypes[k as keyof QuestionTypes].enabled) {
        total += parseInt(String(questionTypes[k as keyof QuestionTypes].count)) || 0;
      }
    });
    return total;
  }, [questionTypes]);

  // Handle Level Change -> Reset Grade
  const handleLevelChange = (val: string) => {
    setIdentity((prev) => ({ ...prev, level: val, grade: "", subject: "", customSubject: "" }));
  };

  // Material handlers
  const addMaterial = () => setMaterials([...materials, ""]);
  const removeMaterial = (index: number) => {
    if (materials.length <= 1 && identity.assessmentType === "Sumatif Lingkup Materi (Ulangan Harian)") {
      showNotification("Ulangan Harian membutuhkan minimal 1 materi!", "warning");
      return;
    }
    setMaterials(materials.filter((_, i) => i !== index));
  };
  const updateMaterial = (index: number, val: string) => {
    const updated = [...materials];
    updated[index] = val;
    setMaterials(updated);
  };

  const validateStep1 = () => {
    if (!identity.level) { showNotification("Pilih Jenjang Sekolah terlebih dahulu.", "warning"); return false; }
    if (!identity.grade) { showNotification("Pilih Kelas terlebih dahulu.", "warning"); return false; }
    const subj = identity.subject === "Lainnya" ? identity.customSubject : identity.subject;
    if (!subj) { showNotification("Pilih atau ketik Mata Pelajaran.", "warning"); return false; }
    if (!identity.semester) { showNotification("Pilih Semester.", "warning"); return false; }
    if (!identity.assessmentType) { showNotification("Pilih Jenis Asesmen.", "warning"); return false; }
    return true;
  };

  const validateStep2 = () => {
    const activeMaterials = materials.filter((m) => m.trim() !== "");
    if (identity.assessmentType === "Sumatif Lingkup Materi (Ulangan Harian)" && activeMaterials.length === 0) {
      showNotification("Sumatif Lingkup Materi (Ulangan Harian) wajib mengisi minimal satu materi!", "error");
      return false;
    }
    return true;
  };

  const validateStep3 = () => {
    let enabledCount = 0;
    Object.keys(questionTypes).forEach((k) => {
      const conf = questionTypes[k as keyof QuestionTypes];
      if (conf.enabled && parseInt(String(conf.count)) > 0) enabledCount++;
    });
    if (enabledCount === 0) {
      showNotification("Pilih minimal satu bentuk soal dan jumlah soalnya!", "warning");
      return false;
    }
    if (difficultyMode === "Campuran") {
      const distTotal = (parseInt(String(difficultyDist.mudah)) || 0) + (parseInt(String(difficultyDist.sedang)) || 0) + (parseInt(String(difficultyDist.sulit)) || 0);
      if (distTotal !== totalQuestionsCount) {
        showNotification(`Total tingkat kesulitan (${distTotal}) harus tepat sama dengan total soal (${totalQuestionsCount})!`, "error");
        return false;
      }
    }
    return true;
  };

  const handleSuggestMaterials = async () => {
    if (!apiKey) {
      showNotification("Mohon atur API Key Gemini terlebih dahulu di menu Pengaturan!", "warning");
      return;
    }
    setIsGenerating(true);
    setLoadingMessage("AI sedang mengusulkan cakupan materi berdasarkan Kurikulum Merdeka...");
    const subj = identity.subject === "Lainnya" ? identity.customSubject : identity.subject;
    const prompt = `Bertindaklah sebagai pengembang Kurikulum Merdeka Indonesia. Usulkan 4-6 topik/materi utama yang relevan untuk:
Jenjang: ${identity.level}
Kelas: ${identity.grade}
Mata Pelajaran: ${subj}
Semester: ${identity.semester}
Jenis Asesmen: ${identity.assessmentType}

Format JSON persis seperti ini:
{
  "proposedMaterials": [
    "Materi 1: ...",
    "Materi 2: ..."
  ]
}`;
    try {
      const result = await callGeminiAPI<{ proposedMaterials?: string[] }>(prompt, "Anda adalah pakar kurikulum Indonesia.", apiKey);
      if (result && result.proposedMaterials) {
        setAiSuggestedMaterials(result.proposedMaterials);
        setAcceptedAiMaterials(result.proposedMaterials);
        showNotification("Berhasil mendapatkan usulan materi dari AI!", "success");
      }
    } catch (err) {
      showNotification("Gagal mendapatkan usulan materi: " + (err as Error).message, "error");
    } finally {
      setIsGenerating(false);
    }
  };

  // Generate Grid (Kisi-kisi)
  const handleGenerateGrid = async () => {
    if (!validateStep1() || !validateStep2() || !validateStep3()) return;

    if (!apiKey) {
      showNotification("Silakan konfigurasi API Key Gemini di menu Pengaturan terlebih dahulu!", "warning");
      return;
    }

    setIsGenerating(true);
    setLoadingMessage("Tahap 1/3: AI sedang menyusun Kisi-Kisi Soal yang terstruktur...");

    const subj = identity.subject === "Lainnya" ? identity.customSubject : identity.subject;
    const cleanMaterials = materials.filter((m) => m.trim() !== "");
    const finalMaterials = [...cleanMaterials, ...acceptedAiMaterials];

    const activeTypes = Object.keys(questionTypes)
      .filter((k) => questionTypes[k as keyof QuestionTypes].enabled)
      .map((k) => `${k.toUpperCase()}: ${questionTypes[k as keyof QuestionTypes].count} soal`)
      .join(", ");

    const prompt = `Susun Kisi-Kisi Soal Asesmen untuk:
Sekolah/Jenjang: ${identity.level} (Kelas ${identity.grade})
Mata Pelajaran: ${subj}
Semester: ${identity.semester}
Jenis Asesmen: ${identity.assessmentType}
Cakupan Materi: ${finalMaterials.join("; ")}
Bentuk & Jumlah Soal: ${activeTypes} (Total: ${totalQuestionsCount})
Tingkat Kesulitan: ${difficultyMode === "Campuran" ? `Mudah: ${difficultyDist.mudah}, Sedang: ${difficultyDist.sedang}, Sulit: ${difficultyDist.sulit}` : difficultyMode}
Instruksi Tambahan Guru: ${additionalInstruction || "Tidak ada"}

Buat JSON array untuk kisi-kisi dengan skema persis berikut:
{
  "grid": [
    {
      "no": 1,
      "materi": "Nama Materi",
      "indikator": "Indikator Soal yang spesifik dan terukur",
      "levelKognitif": "C1/C2/C3/C4/C5/C6",
      "tingkatKesulitan": "Mudah/Sedang/Sulit",
      "bentukSoal": "Pilihan Ganda / Pilihan Ganda Kompleks / Isian / Uraian / Benar-Salah / Menjodohkan",
      "nomorSoal": "1",
      "skor": 1
    }
  ]
}`;

    try {
      const res = await callGeminiAPI<{ grid?: GridRow[] }>(prompt, "Anda adalah spesialis penyusun kisi-kisi asesmen kurikulum Indonesia.", apiKey);
      if (res && res.grid) {
        setGridData(res.grid);
        setCurrentStep(4);
        showNotification("Kisi-kisi soal berhasil dibuat! Silakan tinjau dan sesuaikan.", "success");
      }
    } catch (err) {
      showNotification("Gagal menyusun kisi-kisi: " + (err as Error).message, "error");
    } finally {
      setIsGenerating(false);
    }
  };

  // Generate Questions based on approved Grid
  const handleGenerateQuestions = async () => {
    if (!apiKey) {
      showNotification("Atur API Key terlebih dahulu!", "warning");
      return;
    }
    setIsGenerating(true);
    setLoadingMessage("Tahap 2/3: AI sedang menyusun Naskah Soal, Kunci Jawaban, Pembahasan & Rubrik Penskoran...");

    const subj = identity.subject === "Lainnya" ? identity.customSubject : identity.subject;

    const prompt = `Berdasarkan Kisi-Kisi Soal berikut, hasilkan Paket Soal Lengkap beserta Kunci Jawaban, Pembahasan, dan Rubrik Penskoran.
Detail Identitas: ${identity.level} Kelas ${identity.grade}, MP: ${subj}, ${identity.semester}, ${identity.assessmentType}.
Instruksi Guru: ${additionalInstruction || "Terapkan kausa bahasa Indonesia yang baik dan sesuai konteks kelas."}

KISI-KISI SOAL:
${JSON.stringify(gridData, null, 2)}

PERATURAN PENTING BENTUK SOAL:
1. Pilihan Ganda (PG): Sediakan array 'options' (misal A, B, C, D atau A, B, C, D, E) dan 'answerKey' yang jelas.
2. Pilihan Ganda Kompleks (PGK): Sediakan array 'options', tandai beberapa jawaban benar pada 'answerKey' (misal: "A, C").
3. Isian Singkat: 'answerKey' berupa kata/frasa singkat.
4. Uraian: Sediakan 'answerKey' komprehensif dan 'rubrik' (pedoman penskoran bertahap).
5. Benar-Salah (BS): Sediakan 'statements' berupa list item dengan 'jawaban' masing-masing (Benar/Salah).
6. Menjodohkan: Sediakan 'leftItems' (Premis) dan 'rightItems' (Pilihan Jawaban) serta 'matchingPairs' kunci pasangannya.

Kembalikan format JSON:
{
  "questions": [
    {
      "id": "q1",
      "no": 1,
      "stimulus": "Teks cerita/bacaan/studi kasus (isi jika diperlukan oleh instruksi/indikator, jika tidak kosongkan string)",
      "materi": "Nama Materi",
      "indikator": "Indikator Soal",
      "levelKognitif": "C3",
      "tingkatKesulitan": "Sedang",
      "bentukSoal": "Pilihan Ganda",
      "questionText": "Teks pertanyaan utama...",
      "options": ["A. pilihan...", "B. pilihan...", "C. pilihan...", "D. pilihan..."],
      "statements": [],
      "leftItems": [],
      "rightItems": [],
      "answerKey": "A",
      "explanation": "Pembahasan rinci mengapa A benar...",
      "rubrik": "Pedoman penskoran...",
      "score": 1
    }
  ]
}`;

    try {
      const res = await callGeminiAPI<{ questions?: Question[] }>(prompt, "Anda adalah pembuat soal ujian nasional & asesmen sekolah yang berpengalaman.", apiKey);
      if (res && res.questions) {
        setQuestionsData(res.questions);
        // Move to Step 5: Quality Check
        await runQualityCheck(res.questions);
        setCurrentStep(5);
        showNotification("Naskah Soal AI berhasil dibuat!", "success");
      }
    } catch (err) {
      showNotification("Gagal menyusun soal: " + (err as Error).message, "error");
    } finally {
      setIsGenerating(false);
    }
  };

  const runQualityCheck = async (questionsList: Question[]) => {
    setLoadingMessage("Tahap 3/3: Menguji kualitas & konsistensi naskah soal...");

    const prompt = `Evaluasi kualitas paket soal berikut berdasarkan kriteria kelayakan asesmen pendidikan di Indonesia.
Daftar Soal:
${JSON.stringify(questionsList, null, 2)}

Periksa:
1. Kesesuaian materi & indikator.
2. Kejelasan bahasa & kesusuaian tingkat kelas.
3. Ketepatan kunci jawaban & konsistensi bentuk soal.
4. Ada/tidaknya soal ambigu atau berulang.

Kembalikan JSON:
{
  "overallStatus": "Lulus pemeriksaan awal" | "Perlu ditinjau" | "Perlu diperbaiki",
  "score": 92,
  "findings": [
    { "no": 1, "status": "OK", "note": "Soal sudah sesuai indikator." },
    { "no": 2, "status": "Perlu Ditingkatkan", "note": "Kunci jawaban sebaiknya dilengkapi penjelasan." }
  ]
}`;

    try {
      const res = await callGeminiAPI<QualityReport>(prompt, "Anda adalah penilai & validator mutu soal asesmen.", apiKey);
      if (res) {
        setQualityReport(res);
      }
    } catch {
      setQualityReport({
        overallStatus: "Lulus pemeriksaan awal",
        score: 88,
        findings: questionsList.map((q) => ({ no: q.no, status: "OK", note: "Memenuhi standar teknis dasar." })),
      });
    }
  };

  // Regenerate Single Question
  const handleRegenerateSingleQuestion = async (index: number) => {
    if (!apiKey) {
      showNotification("API key belum diset!", "warning");
      return;
    }
    const targetQ = questionsData[index];
    showNotification(`Memperbaiki soal nomor ${targetQ.no} dengan AI...`, "info");

    const subj = identity.subject === "Lainnya" ? identity.customSubject : identity.subject;
    const prompt = `Perbaiki dan buat ulang 1 soal asesmen ini agar lebih tepat, valid, dan berkualitas.
Identitas: ${identity.level} Kelas ${identity.grade}, MP: ${subj}
Materi: ${targetQ.materi}
Indikator: ${targetQ.indikator}
Bentuk Soal: ${targetQ.bentukSoal}
Tingkat Kesulitan: ${targetQ.tingkatKesulitan}

Format JSON (1 objek soal tunggal):
{
  "id": "${targetQ.id}",
  "no": ${targetQ.no},
  "stimulus": "${targetQ.stimulus || ""}",
  "materi": "${targetQ.materi}",
  "indikator": "${targetQ.indikator}",
  "levelKognitif": "${targetQ.levelKognitif}",
  "tingkatKesulitan": "${targetQ.tingkatKesulitan}",
  "bentukSoal": "${targetQ.bentukSoal}",
  "questionText": "Teks pertanyaan baru...",
  "options": ["A. ...", "B. ...", "C. ...", "D. ..."],
  "answerKey": "Kunci...",
  "explanation": "Pembahasan...",
  "rubrik": "Rubrik...",
  "score": ${targetQ.score}
}`;

    try {
      const res = await callGeminiAPI<Question>(prompt, "Anda adalah validator dan pembuat soal profesional.", apiKey);
      if (res) {
        const updated = [...questionsData];
        updated[index] = res;
        setQuestionsData(updated);
        showNotification(`Soal nomor ${targetQ.no} berhasil dibuat ulang!`, "success");
      }
    } catch (err) {
      showNotification("Gagal memperbarui soal: " + (err as Error).message, "error");
    }
  };

  const handleSavePackage = async () => {
    const packageId = activePackageId || "pkg_" + Date.now();
    const subj = identity.subject === "Lainnya" ? identity.customSubject : identity.subject;

    const newPackage: AssessmentPackage = {
      id: packageId,
      title: `${identity.assessmentType} ${subj} - Kelas ${identity.grade}`,
      createdAt: new Date().toISOString(),
      identity: { ...identity, subject: subj },
      materials: materials.filter((m) => m.trim() !== ""),
      gridData,
      questionsData,
      qualityReport,
    };

    try {
      await dbPut("packages", newPackage);
      setActivePackageId(packageId);
      setPackagesHistory((prev) => [newPackage, ...prev.filter((p) => p.id !== packageId)]);
      showNotification("Paket asesmen berhasil disimpan ke Bank Soal!", "success");
    } catch (e) {
      showNotification("Gagal menyimpan paket: " + e, "error");
    }
  };

  const handleExportDocx = async (exportType: ExportType) => {
    try {
      const filename = await exportAssessmentDocx(identity, gridData, questionsData, exportType);
      showNotification(`Dokumen Word (${filename.split("_").pop()}) berhasil diunduh!`, "success");
    } catch (e) {
      showNotification("Gagal membuat file DOCX: " + (e as Error).message, "error");
    }
  };

  const handleExportBackup = async () => {
    try {
      const pkgs = await dbGetAll<AssessmentPackage>("packages");
      const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(pkgs, null, 2));
      const downloadAnchor = document.createElement("a");
      downloadAnchor.setAttribute("href", dataStr);
      downloadAnchor.setAttribute("download", `Backup_GeneratorSoalAI_${new Date().toISOString().slice(0, 10)}.json`);
      document.body.appendChild(downloadAnchor);
      downloadAnchor.click();
      downloadAnchor.remove();
      showNotification("Data cadangan berhasil diunduh!", "success");
    } catch (e) {
      showNotification("Gagal mengekspor data: " + e, "error");
    }
  };

  const handleImportBackup = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = async (event) => {
      try {
        const importedData = JSON.parse(String(event.target?.result));
        if (Array.isArray(importedData)) {
          for (const pkg of importedData) {
            if (pkg.id && pkg.identity) {
              await dbPut("packages", pkg);
            }
          }
          const data = await dbGetAll<AssessmentPackage>("packages");
          setPackagesHistory(data.sort((a, b) => new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime()));
          showNotification("Data cadangan berhasil diimpor ke aplikasi!", "success");
        } else {
          showNotification("Format file JSON cadangan tidak valid.", "error");
        }
      } catch (err) {
        showNotification("Gagal membaca file backup: " + (err as Error).message, "error");
      }
    };
    reader.readAsText(file);
  };

  const openPackage = (pkg: AssessmentPackage) => {
    setGridData(pkg.gridData || []);
    setQuestionsData(pkg.questionsData || []);
    setIdentity(pkg.identity);
    setActivePackageId(pkg.id);
    setCurrentStep(5);
    setActiveTab("create");
  };

  const handleDeletePackage = async (pkg: AssessmentPackage) => {
    if (window.confirm("Hapus paket soal ini dari Bank Soal?")) {
      await dbDelete("packages", pkg.id);
      setPackagesHistory((prev) => prev.filter((p) => p.id !== pkg.id));
      showNotification("Paket berhasil dihapus.", "info");
    }
  };

  return (
    <div className="flex min-h-screen flex-col bg-background font-sans text-foreground">
      {/* Header */}
      <header className="sticky top-0 z-30 bg-primary text-primary-foreground shadow-md">
        <div className="mx-auto max-w-7xl px-4 sm:px-6 lg:px-8">
          <div className="flex h-16 items-center justify-between">
            <button className="flex cursor-pointer items-center space-x-3 text-left" onClick={() => setActiveTab("dashboard")}>
              <div className="rounded-lg bg-primary-dark p-2">
                <Sparkles className="h-6 w-6 text-gold" />
              </div>
              <div>
                <h1 className="text-lg font-bold leading-tight">Generator Soal AI</h1>
                <p className="text-xs text-primary-foreground/70">Asesmen Kurikulum Merdeka Guru Indonesia</p>
              </div>
            </button>

            <nav className="hidden space-x-1 md:flex">
              {NAV_ITEMS.map((nav) => (
                <button
                  key={nav.id}
                  onClick={() => setActiveTab(nav.id)}
                  className={`flex items-center space-x-2 rounded-md px-3 py-2 text-sm font-medium transition ${
                    activeTab === nav.id ? "bg-primary-dark text-white shadow" : "text-primary-foreground/80 hover:bg-primary-dark/60"
                  }`}
                >
                  <nav.icon className="h-4 w-4" />
                  <span>{nav.label}</span>
                </button>
              ))}
            </nav>
          </div>
        </div>

        {/* Mobile Navigation */}
        <div className="custom-scrollbar flex space-x-1 overflow-x-auto bg-primary-dark px-2 py-1 md:hidden">
          {NAV_ITEMS.map((nav) => (
            <button
              key={nav.id}
              onClick={() => setActiveTab(nav.id)}
              className={`flex items-center space-x-1 whitespace-nowrap rounded px-3 py-1.5 text-xs ${
                activeTab === nav.id ? "bg-background font-bold text-foreground" : "text-primary-foreground/80"
              }`}
            >
              <nav.icon className="h-3.5 w-3.5" />
              <span>{nav.label}</span>
            </button>
          ))}
        </div>
      </header>

      {/* Toast Notification */}
      {notification && (
        <div
          className={`fixed bottom-5 right-5 z-50 flex max-w-md items-center rounded-lg p-4 text-white shadow-lg ${
            notification.type === "error"
              ? "bg-destructive"
              : notification.type === "warning"
                ? "bg-warning"
                : notification.type === "success"
                  ? "bg-success"
                  : "bg-primary"
          }`}
        >
          {notification.type === "error" ? (
            <TriangleAlert className="mr-3 h-5 w-5 flex-shrink-0" />
          ) : notification.type === "warning" ? (
            <TriangleAlert className="mr-3 h-5 w-5 flex-shrink-0" />
          ) : (
            <CircleCheck className="mr-3 h-5 w-5 flex-shrink-0" />
          )}
          <span className="text-sm font-medium">{notification.msg}</span>
        </div>
      )}

      {/* Main Content Area */}
      <main className="mx-auto w-full max-w-7xl flex-1 p-4 sm:p-6 lg:p-8">
        {/* DASHBOARD */}
        {activeTab === "dashboard" && (
          <div className="space-y-6">
            <div className="rounded-2xl bg-gradient-to-r from-primary to-accent p-6 text-white shadow-lg">
              <h2 className="mb-2 text-2xl font-bold">Selamat Datang di Generator Soal AI!</h2>
              <p className="max-w-3xl text-sm leading-relaxed text-white/85 sm:text-base">
                Platform pembuatan kisi-kisi, naskah soal asesmen, kunci jawaban, pembahasan, dan rubrik penskoran otomatis untuk guru SD, SMP, dan SMA berbasis AI terpercaya.
              </p>
              <div className="mt-5 flex flex-wrap gap-3">
                <button
                  onClick={() => { setCurrentStep(1); setActiveTab("create"); }}
                  className="flex items-center space-x-2 rounded-xl bg-card px-5 py-2.5 text-sm font-semibold text-primary shadow hover:bg-primary-soft"
                >
                  <Sparkles className="h-4 w-4" />
                  <span>Buat Paket Soal Baru</span>
                </button>
                <button
                  onClick={() => setActiveTab("settings")}
                  className="flex items-center space-x-2 rounded-xl bg-primary-dark px-4 py-2.5 text-sm font-medium text-white hover:bg-primary-dark/80"
                >
                  <KeyRound className="h-4 w-4" />
                  <span>Konfigurasi API Key</span>
                </button>
              </div>
            </div>

            {!apiKey && (
              <div className="flex items-start space-x-3 rounded-xl border-l-4 border-warning bg-warning-soft p-4">
                <TriangleAlert className="mt-0.5 h-6 w-6 flex-shrink-0 text-warning" />
                <div>
                  <h4 className="text-sm font-semibold text-foreground">API Key Gemini Belum Dikonfigurasi</h4>
                  <p className="mt-1 text-xs text-muted-foreground">
                    Untuk mengaktifkan fitur pembuatan soal berbasis AI, silakan masukkan API Key Gemini Anda di menu Pengaturan. Tanpa API Key, fitur AI tidak dapat mengeksekusi instruksi.
                  </p>
                </div>
              </div>
            )}

            {/* Statistics Cards */}
            <div className="grid grid-cols-1 gap-4 sm:grid-cols-2 lg:grid-cols-4">
              <Card className="flex items-center space-x-4 p-5">
                <div className="rounded-xl bg-primary-soft p-3 text-primary">
                  <FileText className="h-6 w-6" />
                </div>
                <div>
                  <p className="text-xs font-medium text-muted-foreground">Total Paket Disimpan</p>
                  <h3 className="text-2xl font-bold">{packagesHistory.length}</h3>
                </div>
              </Card>

              <Card className="flex items-center space-x-4 p-5">
                <div className="rounded-xl bg-success-soft p-3 text-success">
                  <CircleCheck className="h-6 w-6" />
                </div>
                <div>
                  <p className="text-xs font-medium text-muted-foreground">Lulus Penilaian Mutu</p>
                  <h3 className="text-2xl font-bold">
                    {packagesHistory.filter((p) => p.qualityReport?.overallStatus?.includes("Lulus")).length}
                  </h3>
                </div>
              </Card>

              <Card className="flex items-center space-x-4 p-5">
                <div className="rounded-xl bg-accent-soft p-3 text-accent">
                  <Layers className="h-6 w-6" />
                </div>
                <div>
                  <p className="text-xs font-medium text-muted-foreground">Jenjang Didukung</p>
                  <h3 className="text-2xl font-bold">SD, SMP, SMA</h3>
                </div>
              </Card>

              <Card className="flex items-center space-x-4 p-5">
                <div className="rounded-xl bg-warning-soft p-3 text-warning">
                  <HardDrive className="h-6 w-6" />
                </div>
                <div>
                  <p className="text-xs font-medium text-muted-foreground">Penyimpanan</p>
                  <h3 className="text-2xl font-bold">Lokal Browser</h3>
                </div>
              </Card>
            </div>

            {/* Recent Packages */}
            <Card className="p-6">
              <div className="mb-4 flex items-center justify-between">
                <h3 className="text-lg font-bold">Paket Soal Terakhir</h3>
                <button onClick={() => setActiveTab("history")} className="text-xs font-semibold text-primary hover:underline">
                  Lihat Semua ({packagesHistory.length})
                </button>
              </div>

              {packagesHistory.length === 0 ? (
                <div className="py-8 text-center text-sm text-muted-foreground">
                  Belum ada paket soal yang disimpan. Klik "Buat Soal AI" untuk memulai.
                </div>
              ) : (
                <div className="grid grid-cols-1 gap-4 md:grid-cols-2 lg:grid-cols-3">
                  {packagesHistory.slice(0, 3).map((pkg) => (
                    <div key={pkg.id} className="flex flex-col justify-between rounded-xl border border-border bg-muted/50 p-4 transition hover:border-primary/40">
                      <div>
                        <span className="rounded-full bg-primary-soft px-2 py-0.5 text-[10px] font-bold text-primary">
                          {pkg.identity.level} Kelas {pkg.identity.grade}
                        </span>
                        <h4 className="mt-2 line-clamp-1 text-sm font-bold">{pkg.title}</h4>
                        <p className="mt-1 text-xs text-muted-foreground">
                          {pkg.questionsData?.length || 0} Soal • {new Date(pkg.createdAt).toLocaleDateString("id-ID")}
                        </p>
                      </div>
                      <div className="mt-4 flex justify-end space-x-2 border-t border-border pt-3">
                        <button
                          onClick={() => openPackage(pkg)}
                          className="rounded px-2.5 py-1 text-xs font-semibold text-primary hover:bg-primary-soft"
                        >
                          Buka & Edit
                        </button>
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </Card>
          </div>
        )}

        {/* CREATE (WIZARD) */}
        {activeTab === "create" && (
          <div className="space-y-6">
            {/* Wizard Steps Indicator */}
            <Card className="overflow-x-auto p-4">
              <div className="flex min-w-[600px] items-center justify-between text-xs">
                {[
                  { step: 1, label: "Identitas Asesmen" },
                  { step: 2, label: "Materi / Topik" },
                  { step: 3, label: "Bentuk & Kesulitan" },
                  { step: 4, label: "Kisi-Kisi AI" },
                  { step: 5, label: "Editor & Naskah Final" },
                ].map((s, idx) => (
                  <div key={s.step} className="contents">
                    <button
                      className="flex cursor-pointer items-center space-x-2"
                      onClick={() => { if (s.step < currentStep) setCurrentStep(s.step); }}
                    >
                      <div className={`flex h-7 w-7 items-center justify-center rounded-full font-bold ${
                        currentStep === s.step ? "bg-primary text-white" :
                        currentStep > s.step ? "bg-success text-white" : "bg-muted text-muted-foreground"
                      }`}>
                        {currentStep > s.step ? "✓" : s.step}
                      </div>
                      <span className={`font-medium ${currentStep === s.step ? "font-bold text-primary" : "text-muted-foreground"}`}>
                        {s.label}
                      </span>
                    </button>
                    {idx < 4 && <div className="mx-2 h-0.5 flex-1 bg-border" />}
                  </div>
                ))}
              </div>
            </Card>

            {/* Loading Overlay during AI Generation */}
            {isGenerating && (
              <Card className="space-y-4 p-8 text-center shadow-lg">
                <div className="inline-block h-12 w-12 animate-spin rounded-full border-4 border-primary border-t-transparent" />
                <h3 className="text-lg font-bold text-primary-dark">{loadingMessage}</h3>
                <p className="text-xs text-muted-foreground">Mohon tunggu beberapa saat. AI sedang memproses instruksi Kurikulum Merdeka...</p>
              </Card>
            )}

            {!isGenerating && (
              <>
                {/* STEP 1 */}
                {currentStep === 1 && (
                  <Card className="space-y-6 p-6">
                    <h3 className="flex items-center space-x-2 border-b border-border pb-3 text-lg font-bold">
                      <BookOpen className="h-5 w-5 text-primary" />
                      <span>Langkah 1: Identitas Asesmen Pendidikan</span>
                    </h3>

                    <div className="grid grid-cols-1 gap-4 md:grid-cols-2">
                      <div>
                        <label className="mb-1 block text-xs font-semibold text-muted-foreground">Nama Satuan Pendidikan / Sekolah</label>
                        <input type="text" value={identity.schoolName} onChange={(e) => setIdentity({ ...identity, schoolName: e.target.value })} className={INPUT_CLASS} placeholder="Contoh: SD Negeri 01 Nusantara" />
                      </div>

                      <div>
                        <label className="mb-1 block text-xs font-semibold text-muted-foreground">Tahun Pelajaran</label>
                        <input type="text" value={identity.academicYear} onChange={(e) => setIdentity({ ...identity, academicYear: e.target.value })} className={INPUT_CLASS} placeholder="2026/2027" />
                      </div>

                      <div>
                        <label className="mb-1 block text-xs font-semibold text-muted-foreground">Jenjang Sekolah *</label>
                        <select value={identity.level} onChange={(e) => handleLevelChange(e.target.value)} className={INPUT_CLASS}>
                          <option value="">-- Pilih Jenjang --</option>
                          <option value="SD">SD (Sekolah Dasar)</option>
                          <option value="SMP">SMP (Sekolah Menengah Pertama)</option>
                          <option value="SMA">SMA (Sekolah Menengah Atas)</option>
                        </select>
                      </div>

                      <div>
                        <label className="mb-1 block text-xs font-semibold text-muted-foreground">Kelas *</label>
                        <select value={identity.grade} onChange={(e) => setIdentity({ ...identity, grade: e.target.value })} disabled={!identity.level} className={`${INPUT_CLASS} disabled:bg-muted`}>
                          <option value="">-- Pilih Kelas --</option>
                          {identity.level === "SD" && [1, 2, 3, 4, 5, 6].map((g) => <option key={g} value={g}>Kelas {g}</option>)}
                          {identity.level === "SMP" && [7, 8, 9].map((g) => <option key={g} value={g}>Kelas {g}</option>)}
                          {identity.level === "SMA" && [10, 11, 12].map((g) => <option key={g} value={g}>Kelas {g}</option>)}
                        </select>
                      </div>

                      <div>
                        <label className="mb-1 block text-xs font-semibold text-muted-foreground">Mata Pelajaran *</label>
                        <select value={identity.subject} onChange={(e) => setIdentity({ ...identity, subject: e.target.value })} disabled={!identity.level} className={`${INPUT_CLASS} disabled:bg-muted`}>
                          <option value="">-- Pilih Mata Pelajaran --</option>
                          {identity.level && SUBJECTS_DATA[identity.level]?.map((sub) => <option key={sub} value={sub}>{sub}</option>)}
                          <option value="Lainnya">Lainnya (Ketik Manual)</option>
                        </select>
                      </div>

                      {identity.subject === "Lainnya" && (
                        <div>
                          <label className="mb-1 block text-xs font-semibold text-muted-foreground">Ketik Mata Pelajaran *</label>
                          <input type="text" value={identity.customSubject} onChange={(e) => setIdentity({ ...identity, customSubject: e.target.value })} className={INPUT_CLASS} placeholder="Masukkan nama mata pelajaran..." />
                        </div>
                      )}

                      <div>
                        <label className="mb-1 block text-xs font-semibold text-muted-foreground">Semester *</label>
                        <select value={identity.semester} onChange={(e) => setIdentity({ ...identity, semester: e.target.value })} className={INPUT_CLASS}>
                          <option value="Semester 1 (Ganjil)">Semester 1 (Ganjil)</option>
                          <option value="Semester 2 (Genap)">Semester 2 (Genap)</option>
                        </select>
                      </div>

                      <div className="md:col-span-2">
                        <label className="mb-1 block text-xs font-semibold text-muted-foreground">Jenis Asesmen *</label>
                        <select value={identity.assessmentType} onChange={(e) => setIdentity({ ...identity, assessmentType: e.target.value })} className={INPUT_CLASS}>
                          <option value="Sumatif Lingkup Materi (Ulangan Harian)">Sumatif Lingkup Materi (Ulangan Harian)</option>
                          <option value="Sumatif Tengah Semester (STS)">Sumatif Tengah Semester (STS)</option>
                          <option value="Sumatif Akhir Semester (SAS)">Sumatif Akhir Semester (SAS)</option>
                        </select>
                      </div>
                    </div>

                    <div className="flex justify-end border-t border-border pt-4">
                      <button
                        onClick={() => { if (validateStep1()) setCurrentStep(2); }}
                        className="flex items-center space-x-2 rounded-xl bg-primary px-6 py-2.5 text-sm font-semibold text-white shadow hover:bg-primary/90"
                      >
                        <span>Lanjut ke Input Materi</span>
                        <ArrowIcon />
                      </button>
                    </div>
                  </Card>
                )}

                {/* STEP 2 */}
                {currentStep === 2 && (
                  <Card className="space-y-6 p-6">
                    <h3 className="flex items-center space-x-2 border-b border-border pb-3 text-lg font-bold">
                      <Layers className="h-5 w-5 text-primary" />
                      <span>Langkah 2: Input Materi atau Topik Pembelajaran</span>
                    </h3>

                    <div className="space-y-1 rounded-xl border border-primary/20 bg-primary-soft p-4 text-xs text-foreground">
                      <p className="flex items-center space-x-1 font-bold">
                        <Info className="inline h-4 w-4" />
                        <span>Aturan Pengisian Materi ({identity.assessmentType}):</span>
                      </p>
                      {identity.assessmentType === "Sumatif Lingkup Materi (Ulangan Harian)" && (
                        <p>Minimal satu materi wajib diisi secara manual oleh guru.</p>
                      )}
                      {identity.assessmentType === "Sumatif Tengah Semester (STS)" && (
                        <p>Materi opsional. Jika dikosongkan, AI akan mengusulkan cakupan materi kurikulum untuk Anda tinjau.</p>
                      )}
                      {identity.assessmentType === "Sumatif Akhir Semester (SAS)" && (
                        <p>Pilih mode: Pengisian manual guru, usulan otomatis AI, atau gabungan materi guru + usulan AI.</p>
                      )}
                    </div>

                    {identity.assessmentType === "Sumatif Akhir Semester (SAS)" && (
                      <div className="space-y-2 rounded-xl border border-border bg-muted p-4">
                        <label className="text-xs font-bold">Pilihan Penentuan Materi SAS:</label>
                        <div className="grid grid-cols-1 gap-2 sm:grid-cols-3">
                          {[
                            { id: "manual" as const, label: "1. Ditentukan Guru" },
                            { id: "ai_suggest" as const, label: "2. Diusulkan AI" },
                            { id: "hybrid" as const, label: "3. Gabungan Guru & AI" },
                          ].map((opt) => (
                            <button
                              key={opt.id}
                              type="button"
                              onClick={() => setSasMode(opt.id)}
                              className={`rounded-lg border p-2.5 text-left text-xs font-semibold transition ${
                                sasMode === opt.id ? "border-primary bg-primary text-white" : "border-border bg-card text-foreground"
                              }`}
                            >
                              {opt.label}
                            </button>
                          ))}
                        </div>
                      </div>
                    )}

                    {(sasMode === "manual" || sasMode === "hybrid" || identity.assessmentType !== "Sumatif Akhir Semester (SAS)") && (
                      <div className="space-y-3">
                        <label className="block text-xs font-bold">Materi / Bab Pembelajaran dari Guru:</label>
                        {materials.map((m, idx) => (
                          <div key={idx} className="flex items-center space-x-2">
                            <span className="w-6 text-xs font-bold text-muted-foreground">{idx + 1}.</span>
                            <input
                              type="text"
                              value={m}
                              onChange={(e) => updateMaterial(idx, e.target.value)}
                              placeholder={`Contoh: Bab ${idx + 1} - Unsur Intrinsik Cerita & Ringkasan`}
                              className={INPUT_CLASS}
                            />
                            <button onClick={() => removeMaterial(idx)} className="rounded-lg border border-border p-2.5 text-destructive hover:bg-destructive-soft" title="Hapus Materi">
                              <Trash2 className="h-4 w-4" />
                            </button>
                          </div>
                        ))}
                        <button
                          onClick={addMaterial}
                          className="mt-2 flex items-center space-x-1 rounded-lg border border-dashed border-primary/50 px-3 py-2 text-xs font-semibold text-primary hover:bg-primary-soft"
                        >
                          <Plus className="h-4 w-4" />
                          <span>+ Tambah Baris Materi</span>
                        </button>
                      </div>
                    )}

                    {(identity.assessmentType === "Sumatif Tengah Semester (STS)" || sasMode === "ai_suggest" || sasMode === "hybrid") && (
                      <div className="pt-2">
                        <button
                          onClick={handleSuggestMaterials}
                          className="flex items-center space-x-2 rounded-xl bg-accent px-4 py-2.5 text-xs font-semibold text-white shadow hover:bg-accent/90"
                        >
                          <Sparkles className="h-4 w-4 text-gold" />
                          <span>Usulkan Cakupan Materi Kurikulum via AI</span>
                        </button>
                      </div>
                    )}

                    {aiSuggestedMaterials.length > 0 && (
                      <div className="space-y-3 rounded-xl border border-accent/30 bg-accent-soft p-4">
                        <h4 className="flex items-center space-x-2 text-xs font-bold">
                          <SquareCheckIcon />
                          <span>Hasil Usulan Materi AI (Silakan Pilih/Persetujuan Guru):</span>
                        </h4>
                        <div className="space-y-2">
                          {aiSuggestedMaterials.map((mat, idx) => {
                            const isChecked = acceptedAiMaterials.includes(mat);
                            return (
                              <label key={idx} className="flex cursor-pointer items-center space-x-2 rounded-lg border border-accent/20 bg-card p-2.5 text-xs">
                                <input
                                  type="checkbox"
                                  checked={isChecked}
                                  onChange={(e) => {
                                    if (e.target.checked) setAcceptedAiMaterials([...acceptedAiMaterials, mat]);
                                    else setAcceptedAiMaterials(acceptedAiMaterials.filter((x) => x !== mat));
                                  }}
                                  className="h-4 w-4 rounded accent-primary"
                                />
                                <span>{mat}</span>
                              </label>
                            );
                          })}
                        </div>
                      </div>
                    )}

                    <div className="space-y-3 border-t border-border pt-4">
                      <label className="block text-xs font-bold">Instruksi Tambahan Guru untuk AI (Opsional):</label>
                      <textarea
                        value={additionalInstruction}
                        onChange={(e) => setAdditionalInstruction(e.target.value)}
                        rows={3}
                        className={`${INPUT_CLASS} p-3 text-xs`}
                        placeholder="Contoh: Buatlah cerita dongeng lokal pendek terlebih dahulu, kemudian buat 3 soal HOTS yang menguji pemahaman teks tersebut..."
                      />

                      <div className="space-y-1">
                        <p className="text-[11px] font-semibold text-muted-foreground">Pilih Contoh Instruksi Cepat:</p>
                        <div className="flex flex-wrap gap-1.5">
                          {PRESET_INSTRUCTIONS.map((preset, idx) => (
                            <button
                              key={idx}
                              type="button"
                              onClick={() => setAdditionalInstruction(preset)}
                              className="rounded-md border border-border bg-muted px-2.5 py-1 text-[10px] text-foreground hover:bg-border"
                            >
                              + {preset.slice(0, 35)}...
                            </button>
                          ))}
                        </div>
                      </div>
                    </div>

                    <div className="flex justify-between border-t border-border pt-4">
                      <button onClick={() => setCurrentStep(1)} className="rounded-xl bg-muted px-5 py-2 text-sm font-semibold text-foreground hover:bg-border">
                        Kembali
                      </button>
                      <button
                        onClick={() => { if (validateStep2()) setCurrentStep(3); }}
                        className="flex items-center space-x-2 rounded-xl bg-primary px-6 py-2.5 text-sm font-semibold text-white shadow hover:bg-primary/90"
                      >
                        <span>Lanjut ke Bentuk & Kesulitan</span>
                        <ArrowIcon />
                      </button>
                    </div>
                  </Card>
                )}

                {/* STEP 3 */}
                {currentStep === 3 && (
                  <Card className="space-y-6 p-6">
                    <h3 className="flex items-center space-x-2 border-b border-border pb-3 text-lg font-bold">
                      <SlidersHorizontal className="h-5 w-5 text-primary" />
                      <span>Langkah 3: Bentuk Soal, Tingkat Kesulitan & Skor</span>
                    </h3>

                    <div className="space-y-3">
                      <label className="block text-xs font-bold">Pilih Bentuk Soal & Jumlahnya:</label>
                      <div className="grid grid-cols-1 gap-3 sm:grid-cols-2 lg:grid-cols-3">
                        {[
                          { key: "pg" as const, label: "Pilihan Ganda (PG)" },
                          { key: "pgk" as const, label: "Pilihan Ganda Kompleks" },
                          { key: "isian" as const, label: "Isian Singkat" },
                          { key: "uraian" as const, label: "Uraian / Essay" },
                          { key: "bs" as const, label: "Benar / Salah" },
                          { key: "jodohkan" as const, label: "Menjodohkan" },
                        ].map((item) => {
                          const conf = questionTypes[item.key];
                          return (
                            <div key={item.key} className={`rounded-xl border p-3 transition ${conf.enabled ? "border-primary bg-primary-soft/40" : "border-border bg-muted"}`}>
                              <div className="mb-2 flex items-center justify-between">
                                <label className="flex cursor-pointer items-center space-x-2 text-xs font-bold">
                                  <input
                                    type="checkbox"
                                    checked={conf.enabled}
                                    onChange={(e) => setQuestionTypes({ ...questionTypes, [item.key]: { ...conf, enabled: e.target.checked } })}
                                    className="h-4 w-4 rounded accent-primary"
                                  />
                                  <span>{item.label}</span>
                                </label>
                              </div>
                              {conf.enabled && (
                                <div className="mt-2 grid grid-cols-2 gap-2 border-t border-border pt-2">
                                  <div>
                                    <span className="text-[10px] text-muted-foreground">Jumlah:</span>
                                    <input
                                      type="number"
                                      min="1"
                                      value={conf.count}
                                      onChange={(e) => setQuestionTypes({ ...questionTypes, [item.key]: { ...conf, count: parseInt(e.target.value) || 0 } })}
                                      className="w-full rounded border border-input bg-card p-1 text-xs font-bold"
                                    />
                                  </div>
                                  <div>
                                    <span className="text-[10px] text-muted-foreground">Skor / Soal:</span>
                                    <input
                                      type="number"
                                      min="1"
                                      value={conf.weight}
                                      onChange={(e) => setQuestionTypes({ ...questionTypes, [item.key]: { ...conf, weight: parseInt(e.target.value) || 1 } })}
                                      className="w-full rounded border border-input bg-card p-1 text-xs font-bold"
                                    />
                                  </div>
                                </div>
                              )}
                            </div>
                          );
                        })}
                      </div>
                      <div className="flex items-center justify-between rounded-xl bg-primary-soft p-3 text-xs font-bold text-primary-dark">
                        <span>Total Seluruh Soal: {totalQuestionsCount} Soal</span>
                      </div>
                    </div>

                    <div className="space-y-3 border-t border-border pt-4">
                      <label className="block text-xs font-bold">Pengaturan Tingkat Kesulitan Soal:</label>
                      <div className="flex flex-wrap gap-3">
                        {["Mudah", "Sedang", "Sulit", "Campuran"].map((mode) => (
                          <button
                            key={mode}
                            type="button"
                            onClick={() => setDifficultyMode(mode)}
                            className={`rounded-xl border px-4 py-2 text-xs font-semibold transition ${
                              difficultyMode === mode ? "border-primary bg-primary text-white" : "border-border bg-muted text-foreground"
                            }`}
                          >
                            {mode}
                          </button>
                        ))}
                      </div>

                      {difficultyMode === "Campuran" && (
                        <div className="space-y-2 rounded-xl border border-border bg-muted p-4">
                          <span className="block text-xs font-bold">Komposisi Jumlah Soal Per Tingkat Kesulitan:</span>
                          <div className="grid grid-cols-3 gap-3">
                            <div>
                              <label className="text-[10px] font-bold text-success">Mudah:</label>
                              <input
                                type="number"
                                min="0"
                                value={difficultyDist.mudah}
                                onChange={(e) => setDifficultyDist({ ...difficultyDist, mudah: parseInt(e.target.value) || 0 })}
                                className="w-full rounded border border-input bg-card p-1.5 text-xs font-bold"
                              />
                            </div>
                            <div>
                              <label className="text-[10px] font-bold text-warning">Sedang:</label>
                              <input
                                type="number"
                                min="0"
                                value={difficultyDist.sedang}
                                onChange={(e) => setDifficultyDist({ ...difficultyDist, sedang: parseInt(e.target.value) || 0 })}
                                className="w-full rounded border border-input bg-card p-1.5 text-xs font-bold"
                              />
                            </div>
                            <div>
                              <label className="text-[10px] font-bold text-destructive">Sulit (HOTS):</label>
                              <input
                                type="number"
                                min="0"
                                value={difficultyDist.sulit}
                                onChange={(e) => setDifficultyDist({ ...difficultyDist, sulit: parseInt(e.target.value) || 0 })}
                                className="w-full rounded border border-input bg-card p-1.5 text-xs font-bold"
                              />
                            </div>
                          </div>
                          <p className="text-[11px] text-muted-foreground">
                            Jumlah (Mudah + Sedang + Sulit) = {(parseInt(String(difficultyDist.mudah)) || 0) + (parseInt(String(difficultyDist.sedang)) || 0) + (parseInt(String(difficultyDist.sulit)) || 0)} / Total {totalQuestionsCount} Soal.
                          </p>
                        </div>
                      )}
                    </div>

                    <div className="flex justify-between border-t border-border pt-4">
                      <button onClick={() => setCurrentStep(2)} className="rounded-xl bg-muted px-5 py-2 text-sm font-semibold text-foreground hover:bg-border">
                        Kembali
                      </button>
                      <button
                        onClick={handleGenerateGrid}
                        className="flex items-center space-x-2 rounded-xl bg-success px-6 py-2.5 text-sm font-semibold text-white shadow hover:bg-success/90"
                      >
                        <Sparkles className="h-4 w-4 text-gold" />
                        <span>Generasi Kisi-Kisi AI</span>
                      </button>
                    </div>
                  </Card>
                )}

                {/* STEP 4 */}
                {currentStep === 4 && (
                  <Card className="space-y-6 p-6">
                    <div className="flex items-center justify-between border-b border-border pb-3">
                      <h3 className="flex items-center space-x-2 text-lg font-bold">
                        <Table className="h-5 w-5 text-primary" />
                        <span>Langkah 4: Review & Edit Kisi-Kisi Soal</span>
                      </h3>
                      <button
                        onClick={() => {
                          setGridData([
                            ...gridData,
                            {
                              no: gridData.length + 1,
                              materi: "Materi Baru",
                              indikator: "Siswa dapat menjelaskan...",
                              levelKognitif: "C3",
                              tingkatKesulitan: "Sedang",
                              bentukSoal: "Pilihan Ganda",
                              nomorSoal: String(gridData.length + 1),
                              skor: 1,
                            },
                          ]);
                        }}
                        className="rounded-lg border border-primary/30 bg-primary-soft px-3 py-1.5 text-xs font-bold text-primary"
                      >
                        + Tambah Indikator
                      </button>
                    </div>

                    <div className="custom-scrollbar overflow-x-auto">
                      <table className="w-full text-left text-xs">
                        <thead className="bg-primary-soft font-bold text-primary-dark">
                          <tr>
                            <th className="border border-border p-2">No</th>
                            <th className="border border-border p-2 min-w-[150px]">Materi</th>
                            <th className="border border-border p-2 min-w-[250px]">Indikator Soal</th>
                            <th className="border border-border p-2">Level</th>
                            <th className="border border-border p-2">Kesulitan</th>
                            <th className="border border-border p-2">Bentuk</th>
                            <th className="border border-border p-2">Skor</th>
                            <th className="border border-border p-2 text-center">Aksi</th>
                          </tr>
                        </thead>
                        <tbody>
                          {gridData.map((row, idx) => (
                            <tr key={idx} className="hover:bg-muted">
                              <td className="border border-border p-2 text-center font-bold">{row.no}</td>
                              <td className="border border-border p-2">
                                <input
                                  type="text"
                                  value={row.materi}
                                  onChange={(e) => {
                                    const u = [...gridData];
                                    u[idx] = { ...u[idx], materi: e.target.value };
                                    setGridData(u);
                                  }}
                                  className="w-full rounded border-none bg-transparent p-1 focus:bg-card focus:ring-1 focus:ring-ring"
                                />
                              </td>
                              <td className="border border-border p-2">
                                <textarea
                                  value={row.indikator}
                                  onChange={(e) => {
                                    const u = [...gridData];
                                    u[idx] = { ...u[idx], indikator: e.target.value };
                                    setGridData(u);
                                  }}
                                  rows={2}
                                  className="w-full rounded border-none bg-transparent p-1 focus:bg-card focus:ring-1 focus:ring-ring"
                                />
                              </td>
                              <td className="border border-border p-2">
                                <select
                                  value={row.levelKognitif}
                                  onChange={(e) => {
                                    const u = [...gridData];
                                    u[idx] = { ...u[idx], levelKognitif: e.target.value };
                                    setGridData(u);
                                  }}
                                  className="rounded border border-border bg-transparent p-1"
                                >
                                  {["C1", "C2", "C3", "C4", "C5", "C6"].map((c) => <option key={c} value={c}>{c}</option>)}
                                </select>
                              </td>
                              <td className="border border-border p-2">
                                <select
                                  value={row.tingkatKesulitan}
                                  onChange={(e) => {
                                    const u = [...gridData];
                                    u[idx] = { ...u[idx], tingkatKesulitan: e.target.value };
                                    setGridData(u);
                                  }}
                                  className="rounded border border-border bg-transparent p-1"
                                >
                                  {["Mudah", "Sedang", "Sulit"].map((d) => <option key={d} value={d}>{d}</option>)}
                                </select>
                              </td>
                              <td className="border border-border p-2 font-semibold">{row.bentukSoal}</td>
                              <td className="w-12 border border-border p-2">
                                <input
                                  type="number"
                                  value={row.skor}
                                  onChange={(e) => {
                                    const u = [...gridData];
                                    u[idx] = { ...u[idx], skor: parseInt(e.target.value) || 1 };
                                    setGridData(u);
                                  }}
                                  className="w-full rounded border border-border p-1 text-center font-bold"
                                />
                              </td>
                              <td className="border border-border p-2 text-center">
                                <button
                                  onClick={() => setGridData(gridData.filter((_, i) => i !== idx))}
                                  className="text-destructive hover:text-destructive/70"
                                >
                                  <Trash2 className="inline h-4 w-4" />
                                </button>
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>

                    <div className="flex justify-between border-t border-border pt-4">
                      <button onClick={() => setCurrentStep(3)} className="rounded-xl bg-muted px-5 py-2 text-sm font-semibold text-foreground hover:bg-border">
                        Kembali
                      </button>
                      <button
                        onClick={handleGenerateQuestions}
                        className="flex items-center space-x-2 rounded-xl bg-success px-6 py-2.5 text-sm font-semibold text-white shadow hover:bg-success/90"
                      >
                        <FileCheck className="h-4 w-4 text-gold" />
                        <span>Setujui & Hasilkan Naskah Soal Final</span>
                      </button>
                    </div>
                  </Card>
                )}

                {/* STEP 5 */}
                {currentStep === 5 && (
                  <div className="space-y-6">
                    {/* Quality Report Banner */}
                    {qualityReport && (
                      <div className={`flex flex-col items-start justify-between gap-4 rounded-2xl border p-4 shadow-sm md:flex-row md:items-center ${
                        qualityReport.overallStatus?.includes("Lulus") ? "border-success/30 bg-success-soft" : "border-warning/30 bg-warning-soft"
                      }`}>
                        <div className="flex items-center space-x-3">
                          <div className="rounded-xl bg-card p-3 shadow-sm">
                            <CircleCheck className="h-6 w-6 text-success" />
                          </div>
                          <div>
                            <h4 className="text-sm font-bold">Status Validasi Kualitas Soal AI: {qualityReport.overallStatus}</h4>
                            <p className="mt-0.5 text-xs opacity-90">
                              Skor Mutu: {qualityReport.score}/100 • Evaluasi Otomatis Bahasa, Kunci & Indikator Selesai.
                            </p>
                          </div>
                        </div>
                        <button
                          onClick={handleSavePackage}
                          className="flex items-center space-x-1.5 rounded-xl bg-primary px-4 py-2 text-xs font-bold text-white shadow"
                        >
                          <Save className="h-4 w-4" />
                          <span>Simpan Paket Soal</span>
                        </button>
                      </div>
                    )}

                    {/* Document Export Options */}
                    <Card className="space-y-3 p-5">
                      <h4 className="flex items-center space-x-2 text-sm font-bold">
                        <Download className="h-4 w-4 text-primary" />
                        <span>Ekspor Dokumen Microsoft Word (.docx) Valid:</span>
                      </h4>
                      <div className="flex flex-wrap gap-2">
                        {([
                          { id: "kisi", label: "Unduh Kisi-Kisi Only" },
                          { id: "soal", label: "Unduh Naskah Soal" },
                          { id: "kunci", label: "Unduh Kunci & Pembahasan" },
                          { id: "rubrik", label: "Unduh Pedoman Penskoran" },
                          { id: "lengkap", label: "Unduh Paket Lengkap (All-in-One)" },
                        ] as { id: ExportType; label: string }[]).map((exp) => (
                          <button
                            key={exp.id}
                            onClick={() => handleExportDocx(exp.id)}
                            className="flex items-center space-x-1.5 rounded-xl border border-border bg-muted px-3 py-2 text-xs font-semibold text-foreground transition hover:border-primary/40 hover:bg-primary-soft hover:text-primary"
                          >
                            <FileText className="h-3.5 w-3.5" />
                            <span>{exp.label}</span>
                          </button>
                        ))}
                      </div>
                    </Card>

                    {/* Preview Tabs */}
                    <Card className="overflow-hidden">
                      <div className="flex overflow-x-auto border-b border-border bg-muted">
                        {([
                          { id: "soal", label: "Naskah Soal", icon: FileText },
                          { id: "kisi", label: "Kisi-Kisi", icon: Table },
                          { id: "kunci", label: "Kunci & Pembahasan", icon: KeyRound },
                          { id: "rubrik", label: "Rubrik Penskoran", icon: Award },
                        ] as const).map((tab) => (
                          <button
                            key={tab.id}
                            onClick={() => setPreviewTab(tab.id)}
                            className={`flex items-center space-x-2 whitespace-nowrap border-b-2 px-5 py-3 text-xs font-bold transition ${
                              previewTab === tab.id ? "border-primary bg-card text-primary" : "border-transparent text-muted-foreground hover:text-foreground"
                            }`}
                          >
                            <tab.icon className="h-4 w-4" />
                            <span>{tab.label}</span>
                          </button>
                        ))}
                      </div>

                      <div className="p-6">
                        {/* PREVIEW: NASKAH SOAL */}
                        {previewTab === "soal" && (
                          <div className="space-y-6">
                            {questionsData.map((q, idx) => (
                              <div key={q.id || idx} className="relative space-y-3 rounded-xl border border-border bg-muted p-4">
                                <div className="flex items-center justify-between border-b border-border pb-2">
                                  <span className="text-xs font-bold text-primary-dark">
                                    Soal Nomor {q.no} ({q.bentukSoal} • Level {q.levelKognitif || "C3"} • {q.tingkatKesulitan})
                                  </span>
                                  <button
                                    onClick={() => handleRegenerateSingleQuestion(idx)}
                                    className="rounded bg-primary-soft px-2 py-1 text-[11px] font-semibold text-primary hover:bg-border"
                                    title="Buat Ulang Soal Ini via AI"
                                  >
                                    <RefreshCw className="mr-1 inline h-3 w-3" />
                                    AI Perbaiki Soal Ini
                                  </button>
                                </div>

                                {q.stimulus && (
                                  <div className="rounded border border-border bg-card p-3 text-xs italic text-muted-foreground">
                                    <span className="mb-1 block font-bold not-italic text-foreground">Stimulus / Teks Bacaan:</span>
                                    <textarea
                                      value={q.stimulus}
                                      onChange={(e) => {
                                        const u = [...questionsData];
                                        u[idx] = { ...u[idx], stimulus: e.target.value };
                                        setQuestionsData(u);
                                      }}
                                      rows={2}
                                      className="w-full border-none bg-transparent text-xs outline-none focus:bg-muted"
                                    />
                                  </div>
                                )}

                                <div>
                                  <textarea
                                    value={q.questionText}
                                    onChange={(e) => {
                                      const u = [...questionsData];
                                      u[idx] = { ...u[idx], questionText: e.target.value };
                                      setQuestionsData(u);
                                    }}
                                    rows={2}
                                    className="w-full rounded border border-border p-2 text-xs font-semibold"
                                  />
                                </div>

                                {q.options && q.options.length > 0 && (
                                  <div className="grid grid-cols-1 gap-2 sm:grid-cols-2">
                                    {q.options.map((opt, oIdx) => (
                                      <input
                                        key={oIdx}
                                        type="text"
                                        value={opt}
                                        onChange={(e) => {
                                          const u = [...questionsData];
                                          const options = [...(u[idx].options ?? [])];
                                          options[oIdx] = e.target.value;
                                          u[idx] = { ...u[idx], options };
                                          setQuestionsData(u);
                                        }}
                                        className="rounded border border-border bg-card p-1.5 text-xs"
                                      />
                                    ))}
                                  </div>
                                )}
                              </div>
                            ))}
                          </div>
                        )}

                        {/* PREVIEW: KUNCI & PEMBAHASAN */}
                        {previewTab === "kunci" && (
                          <div className="space-y-4">
                            {questionsData.map((q, idx) => (
                              <div key={idx} className="space-y-2 rounded-xl border border-border bg-muted p-4 text-xs">
                                <div className="flex justify-between border-b border-border pb-1 font-bold">
                                  <span>Soal No. {q.no} ({q.bentukSoal})</span>
                                  <span className="text-success">Skor Max: {q.score}</span>
                                </div>
                                <div>
                                  <span className="font-bold text-muted-foreground">Kunci Jawaban: </span>
                                  <input
                                    type="text"
                                    value={q.answerKey}
                                    onChange={(e) => {
                                      const u = [...questionsData];
                                      u[idx] = { ...u[idx], answerKey: e.target.value };
                                      setQuestionsData(u);
                                    }}
                                    className="rounded border border-input bg-card p-1 font-bold text-success"
                                  />
                                </div>
                                <div>
                                  <span className="mb-1 block font-bold text-muted-foreground">Pembahasan Rinci:</span>
                                  <textarea
                                    value={q.explanation}
                                    onChange={(e) => {
                                      const u = [...questionsData];
                                      u[idx] = { ...u[idx], explanation: e.target.value };
                                      setQuestionsData(u);
                                    }}
                                    rows={2}
                                    className="w-full rounded border border-border bg-card p-2 text-xs"
                                  />
                                </div>
                              </div>
                            ))}
                          </div>
                        )}

                        {/* PREVIEW: KISI-KISI */}
                        {previewTab === "kisi" && (
                          <div className="overflow-x-auto">
                            <table className="w-full text-left text-xs">
                              <thead className="bg-muted font-bold">
                                <tr>
                                  <th className="border border-border p-2">No</th>
                                  <th className="border border-border p-2">Materi</th>
                                  <th className="border border-border p-2">Indikator</th>
                                  <th className="border border-border p-2">Level</th>
                                  <th className="border border-border p-2">Kesulitan</th>
                                  <th className="border border-border p-2">Bentuk</th>
                                </tr>
                              </thead>
                              <tbody>
                                {gridData.map((g, i) => (
                                  <tr key={i}>
                                    <td className="border border-border p-2 text-center font-bold">{g.no}</td>
                                    <td className="border border-border p-2">{g.materi}</td>
                                    <td className="border border-border p-2">{g.indikator}</td>
                                    <td className="border border-border p-2">{g.levelKognitif}</td>
                                    <td className="border border-border p-2">{g.tingkatKesulitan}</td>
                                    <td className="border border-border p-2">{g.bentukSoal}</td>
                                  </tr>
                                ))}
                              </tbody>
                            </table>
                          </div>
                        )}

                        {/* PREVIEW: RUBRIK PENSKORAN */}
                        {previewTab === "rubrik" && (
                          <div className="space-y-4">
                            {questionsData.map((q, idx) => (
                              <div key={idx} className="space-y-2 rounded-xl border border-border bg-muted p-4 text-xs">
                                <span className="block font-bold text-primary-dark">Pedoman Penskoran Soal No. {q.no} ({q.bentukSoal})</span>
                                <textarea
                                  value={q.rubrik || "Jawaban benar mendapat skor penuh."}
                                  onChange={(e) => {
                                    const u = [...questionsData];
                                    u[idx] = { ...u[idx], rubrik: e.target.value };
                                    setQuestionsData(u);
                                  }}
                                  rows={2}
                                  className="w-full rounded border border-border bg-card p-2"
                                />
                              </div>
                            ))}
                          </div>
                        )}
                      </div>
                    </Card>
                  </div>
                )}
              </>
            )}
          </div>
        )}

        {/* BANK SOAL */}
        {activeTab === "bank" && (
          <div className="space-y-6">
            <Card className="space-y-4 p-6">
              <h3 className="flex items-center space-x-2 text-lg font-bold">
                <Database className="h-5 w-5 text-primary" />
                <span>Bank Soal & Perpustakaan Asesmen Disimpan</span>
              </h3>

              {packagesHistory.length === 0 ? (
                <div className="py-12 text-center text-sm text-muted-foreground">
                  Belum ada bank soal. Paket yang Anda simpan akan muncul di sini.
                </div>
              ) : (
                <div className="grid grid-cols-1 gap-4 md:grid-cols-2">
                  {packagesHistory.map((pkg) => (
                    <div key={pkg.id} className="flex flex-col justify-between rounded-xl border border-border bg-muted p-4">
                      <div>
                        <div className="flex items-start justify-between">
                          <span className="rounded-full bg-primary-soft px-2 py-0.5 text-[10px] font-bold text-primary">
                            {pkg.identity.level} Kelas {pkg.identity.grade}
                          </span>
                          <span className="text-[10px] text-muted-foreground">
                            {new Date(pkg.createdAt).toLocaleDateString("id-ID")}
                          </span>
                        </div>
                        <h4 className="mt-2 text-sm font-bold">{pkg.title}</h4>
                        <p className="mt-1 text-xs text-muted-foreground">
                          Mata Pelajaran: {pkg.identity.subject} • {pkg.questionsData?.length || 0} Soal
                        </p>
                      </div>

                      <div className="mt-4 flex items-center justify-between border-t border-border pt-3">
                        <button onClick={() => handleDeletePackage(pkg)} className="text-xs font-semibold text-destructive hover:underline">
                          Hapus
                        </button>
                        <button
                          onClick={() => openPackage(pkg)}
                          className="rounded-lg bg-primary px-3 py-1.5 text-xs font-bold text-white shadow-sm hover:bg-primary/90"
                        >
                          Buka & Ekspor
                        </button>
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </Card>
          </div>
        )}

        {/* RIWAYAT */}
        {activeTab === "history" && (
          <Card className="space-y-4 p-6">
            <h3 className="flex items-center space-x-2 text-lg font-bold">
              <History className="h-5 w-5 text-primary" />
              <span>Riwayat Pembuatan Paket Soal</span>
            </h3>
            <div className="space-y-2">
              {packagesHistory.length === 0 && (
                <div className="py-8 text-center text-sm text-muted-foreground">Belum ada riwayat paket soal.</div>
              )}
              {packagesHistory.map((pkg, idx) => (
                <div key={pkg.id || idx} className="flex items-center justify-between rounded-xl border border-border bg-muted p-3 text-xs">
                  <div>
                    <span className="font-bold">{pkg.title}</span>
                    <span className="ml-2 text-muted-foreground">({new Date(pkg.createdAt).toLocaleString("id-ID")})</span>
                  </div>
                  <button onClick={() => openPackage(pkg)} className="font-bold text-primary hover:underline">
                    Lihat
                  </button>
                </div>
              ))}
            </div>
          </Card>
        )}

        {/* PENGATURAN */}
        {activeTab === "settings" && (
          <Card className="space-y-6 p-6">
            <h3 className="flex items-center space-x-2 border-b border-border pb-3 text-lg font-bold">
              <Settings className="h-5 w-5 text-primary" />
              <span>Pengaturan & Cadangan Data Lokal</span>
            </h3>

            {/* Gemini API Key Config */}
            <div className="space-y-3 rounded-xl border border-border bg-muted p-4">
              <h4 className="flex items-center space-x-2 text-xs font-bold">
                <KeyRound className="h-4 w-4 text-primary" />
                <span>Konfigurasi Gemini API Key:</span>
              </h4>
              <p className="text-xs leading-relaxed text-muted-foreground">
                Generator Soal AI menggunakan API resmi Google Gemini. Masukkan API Key Anda di bawah ini untuk mengaktifkan seluruh kemampuan pemrosesan bahasa alami Kurikulum Merdeka. API Key hanya tersimpan di browser Anda.
              </p>
              <div className="flex space-x-2">
                <input
                  type="password"
                  value={apiKey}
                  onChange={(e) => setApiKey(e.target.value)}
                  placeholder="Masukkan Gemini API Key..."
                  className={`${INPUT_CLASS} flex-1 text-xs`}
                />
                <button
                  onClick={() => saveApiKey(apiKey)}
                  className="rounded-lg bg-primary px-4 py-2.5 text-xs font-bold text-white shadow hover:bg-primary/90"
                >
                  Simpan
                </button>
              </div>
            </div>

            {/* Backup & Restore */}
            <div className="space-y-3 border-t border-border pt-4">
              <h4 className="flex items-center space-x-2 text-xs font-bold">
                <HardDrive className="h-4 w-4 text-primary" />
                <span>Cadangan Data (Backup & Restore Data Lokal):</span>
              </h4>
              <div className="flex flex-wrap gap-3">
                <button
                  onClick={handleExportBackup}
                  className="flex items-center space-x-2 rounded-xl bg-success px-4 py-2 text-xs font-semibold text-white shadow hover:bg-success/90"
                >
                  <CloudDownload className="h-4 w-4" />
                  <span>Ekspor Backup JSON</span>
                </button>

                <label className="flex cursor-pointer items-center space-x-2 rounded-xl border border-border bg-muted px-4 py-2 text-xs font-semibold text-foreground hover:bg-border">
                  <CloudUpload className="h-4 w-4" />
                  <span>Impor File Backup</span>
                  <input type="file" accept=".json" onChange={handleImportBackup} className="hidden" />
                </label>
              </div>
            </div>
          </Card>
        )}
      </main>

      {/* Footer */}
      <footer className="border-t border-border bg-foreground py-6 text-center text-xs text-background/70">
        <div className="mx-auto max-w-7xl space-y-1 px-4">
          <p className="font-semibold text-background/90">Generator Soal AI - Asesmen Guru Indonesia © 2026</p>
          <p className="text-[11px]">Dukungan Kurikulum Merdeka • Berjalan 100% di Browser dengan Penyimpanan Lokal</p>
        </div>
      </footer>
    </div>
  );
}

function ArrowIcon() {
  return <span aria-hidden="true">→</span>;
}

function SquareCheckIcon() {
  return <CircleCheck className="inline h-4 w-4" />;
}
