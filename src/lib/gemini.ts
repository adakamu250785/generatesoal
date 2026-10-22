// Client-side call to the official Google Gemini API using the teacher's own API key.
// The key is stored locally in IndexedDB (menu Pengaturan) and never leaves the browser
// except to Google's endpoint.

export async function callGeminiAPI<T = Record<string, unknown>>(
  prompt: string,
  systemInstruction = "",
  apiKey = "",
): Promise<T> {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-3-flash-preview:generateContent?key=${apiKey}`;
  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    systemInstruction: systemInstruction ? { parts: [{ text: systemInstruction }] } : undefined,
    generationConfig: {
      responseMimeType: "application/json",
    },
  };

  let retries = 0;
  const maxRetries = 3;
  while (retries < maxRetries) {
    try {
      const response = await fetch(url, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
      if (!response.ok) {
        const errText = await response.text();
        throw new Error(`API Request Error [${response.status}]: ${errText}`);
      }
      const result = await response.json();
      const jsonText = result?.candidates?.[0]?.content?.parts?.[0]?.text;
      if (!jsonText) throw new Error("Format respons API tidak sesuai.");
      return JSON.parse(jsonText) as T;
    } catch (err) {
      retries++;
      if (retries >= maxRetries) throw err;
      await new Promise((r) => setTimeout(r, 1000 * Math.pow(2, retries)));
    }
  }
  throw new Error("Panggilan API gagal.");
}
