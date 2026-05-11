import express from "express";

const app = express();
app.use(express.json({ limit: "10mb" }));

const GEMINI_API_KEY = process.env.GEMINI_API_KEY;
const GEMINI_MODEL = process.env.GEMINI_MODEL || "gemini-2.5-flash";
const SHARED_SECRET = process.env.SHARED_SECRET || "";

function pad2(n) {
  return String(n).padStart(2, "0");
}

function isValidDate(year, month, day) {
  const dt = new Date(year, month - 1, day);
  return dt.getFullYear() === year && dt.getMonth() === month - 1 && dt.getDate() === day;
}

function normalizeDate(raw, captureYear) {
  if (!raw) return "";
  const text = String(raw).trim();
  if (!text) return "";
  const y = Number(captureYear) || new Date().getFullYear();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  let m = text.match(/^(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})$/);
  if (m) {
    const yy = Number(m[1]);
    const mm = Number(m[2]);
    const dd = Number(m[3]);
    if (!isValidDate(yy, mm, dd)) return "";
    return `${yy}-${pad2(mm)}-${pad2(dd)}`;
  }

  m = text.match(/^(\d{1,2})[\/\-](\d{1,2})$/);
  if (m) {
    const mm = Number(m[1]);
    const dd = Number(m[2]);
    if (!isValidDate(y, mm, dd)) return "";
    if (new Date(y, mm - 1, dd) > today && isValidDate(y - 1, mm, dd)) {
      return `${y - 1}-${pad2(mm)}-${pad2(dd)}`;
    }
    return `${y}-${pad2(mm)}-${pad2(dd)}`;
  }

  m = text.match(/^(\d{1,2})月(\d{1,2})日$/);
  if (m) {
    const mm = Number(m[1]);
    const dd = Number(m[2]);
    if (!isValidDate(y, mm, dd)) return "";
    if (new Date(y, mm - 1, dd) > today && isValidDate(y - 1, mm, dd)) {
      return `${y - 1}-${pad2(mm)}-${pad2(dd)}`;
    }
    return `${y}-${pad2(mm)}-${pad2(dd)}`;
  }

  return "";
}

async function callGeminiWithRetry(url, payload, maxRetries = 3) {
  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    const r = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    });
    const json = await r.json();

    if (json.error) {
      const code = json.error.code || 0;
      if ((code === 429 || code === 503) && attempt < maxRetries) {
        const waitMs = Math.min(10000 + attempt * 10000, 30000);
        console.warn(`Gemini rate limit (${code}), retry ${attempt + 1}/${maxRetries} after ${waitMs}ms`);
        await new Promise((resolve) => setTimeout(resolve, waitMs));
        continue;
      }
      throw new Error(json.error.message || "gemini error");
    }

    return json;
  }
  throw new Error("Gemini: max retries exceeded");
}

app.get("/healthz", (_req, res) => {
  res.json({ ok: true });
});

app.post("/extract", async (req, res) => {
  try {
    if (SHARED_SECRET) {
      const token = req.header("x-shared-secret");
      if (token !== SHARED_SECRET) {
        return res.status(401).json({ success: false, error: "unauthorized" });
      }
    }

    if (!GEMINI_API_KEY) {
      return res.status(500).json({ success: false, error: "GEMINI_API_KEY is not set" });
    }

    const { imageBase64, mimeType = "image/jpeg", captureYear } = req.body || {};
    if (!imageBase64) {
      return res.status(400).json({ success: false, error: "imageBase64 is required" });
    }

    const prompt = [
      "領収書または交通履歴を解析して、JSONのみ返してください。",
      "出力形式: {\"entries\":[{\"date\":\"YYYY-MM-DD\",\"amount\":\"数値のみ\",\"category\":\"費目\"}]}",
      "ルール:",
      "- date は YYYY-MM-DD",
      "- amount は正の整数",
      "- 02/06, 2/6, 2月6日 は 月/日 として解釈する",
      "- 年がない場合は captureYear を使う",
      "- 読み取れない項目は空文字",
      "- JSON以外の文字は返さない"
    ].join("\n");

    const url = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent?key=${GEMINI_API_KEY}`;
    const payload = {
      contents: [{
        role: "user",
        parts: [
          { text: `${prompt}\ncaptureYear: ${captureYear || ""}` },
          {
            inlineData: {
              mimeType,
              data: String(imageBase64).replace(/^data:[^;]+;base64,/, "")
            }
          }
        ]
      }]
    };

    const json = await callGeminiWithRetry(url, payload);

    const parts = json?.candidates?.[0]?.content?.parts || [];
    const text = parts.map((p) => p.text || "").join("\n");
    const match = text.match(/\{[\s\S]*\}/);
    if (!match) {
      return res.status(500).json({ success: false, error: "json parse failed" });
    }

    let parsed = JSON.parse(match[0]);
    if (!Array.isArray(parsed.entries)) parsed = { entries: [parsed] };

    const entries = parsed.entries.map((e) => {
      const amount = String(parseInt(String(e.amount || "").replace(/[^\d]/g, ""), 10) || "");
      return {
        date: normalizeDate(e.date, captureYear),
        amount,
        category: e.category || ""
      };
    });

    res.json({ success: true, entries });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

const port = process.env.PORT || 8080;
app.listen(port, () => {
  console.log(`listening on ${port}`);
});
