<!-- LOVABLE:BEGIN -->
> [!IMPORTANT]
> This project is connected to [Lovable](https://lovable.dev). Avoid rewriting
> published git history — force pushing, or rebasing/amending/squashing commits
> that are already pushed — as it rewrites history on Lovable's side and the
> user will likely lose their project history.
>
> Commits you push to the connected branch sync back to Lovable and show up in
> the editor, so keep the branch in a working state.
<!-- LOVABLE:END -->

# Architecture decisions

- The app is fully client-side: no backend, no auth, no database. All data
  (packages, history, API key) lives in the browser's IndexedDB
  (`src/lib/idb.ts`). Why: the user explicitly chose local-only storage.
- AI runs in the browser calling Google Gemini directly with the teacher's own
  API key (stored locally via menu Pengaturan). Why: the user chose "API key
  milik guru" over a hosted key.
- Word (.docx) export lives in `src/lib/docx-export.ts` and must keep the
  Indonesian school document format (A4, Times New Roman, bordered identity
  table, shaded kisi-kisi table, numbered questions, answer-key summary table,
  rubric, page numbers). Why: document quality is the user's top improvement
  request.
