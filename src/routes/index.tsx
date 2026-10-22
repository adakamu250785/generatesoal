import { createFileRoute } from "@tanstack/react-router";

import { GeneratorApp } from "@/components/GeneratorApp";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "Generator Soal AI — Asesmen Guru Indonesia" },
      {
        name: "description",
        content:
          "Platform pembuatan kisi-kisi, naskah soal asesmen, kunci jawaban, pembahasan, dan rubrik penskoran otomatis untuk guru SD, SMP, dan SMA berbasis AI.",
      },
      { property: "og:title", content: "Generator Soal AI — Asesmen Guru Indonesia" },
      {
        property: "og:description",
        content:
          "Buat kisi-kisi, naskah soal, kunci jawaban, pembahasan, dan rubrik penskoran untuk asesmen Kurikulum Merdeka dengan bantuan AI.",
      },
      { property: "og:type", content: "website" },
      { name: "twitter:card", content: "summary_large_image" },
    ],
  }),
  component: Index,
});

function Index() {
  return <GeneratorApp />;
}
