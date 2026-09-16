export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    if (url.pathname === "/api/health") return Response.json({ ok: true });

    if (url.pathname === "/api/ai-check" && request.method === "POST") {
      try {
        if (!env.GEMINI_API_KEY) return Response.json({ success: false, error: "GEMINI_API_KEY is not configured" }, { status: 500 });

        const body = await request.json();
        const selected = Array.isArray(body.selected) ? body.selected : [];
        const texts = Array.isArray(body.texts) ? body.texts : [];
        const answers = Array.isArray(body.answers) ? body.answers : [];

        if (selected.length === 0 && texts.length === 0) return Response.json({ success: false, error: "Excelの読み取り結果がありません" }, { status: 400 });

        // 氏名・住所・電話番号など、直接個人を特定できる情報はAIへ送信しない。
        // 要介護度・認知症高齢者日常生活自立度は、原因究明・再発防止を考えるための参考情報として送信する。
        const excludedFields = new Set([
          "氏名",
          "住所",
          "連絡先（電話番号）",
          "医療機関名",
          "診断名"
        ]);

        const sensitiveValues = texts
          .filter((item) => item.field === "氏名")
          .map((item) => String(item.value || "").trim())
          .filter(Boolean);

        const sanitizeValue = (value) => {
          let sanitized = String(value || "");
          for (const sensitiveValue of sensitiveValues) sanitized = sanitized.split(sensitiveValue).join("[個人名]");
          sanitized = sanitized.replace(/(?:0\d{1,4}[-ー]?\d{1,4}[-ー]?\d{3,4})/g, "[電話番号]");
          return sanitized;
        };

        const safeSelected = selected
          .filter((item) => !excludedFields.has(item.field))
          .map((item) => ({ field: item.field, value: sanitizeValue(item.value) }));
        const safeTexts = texts
          .filter((item) => !excludedFields.has(item.field))
          .map((item) => ({ field: item.field, value: sanitizeValue(item.value) }));

        const reportText = [
          "【Excelから読み取った情報】",
          ...safeSelected.map((item) => `- ${item.field}: ${item.value}`),
          ...safeTexts.map((item) => `- ${item.field}: ${item.value}`),
          answers.length ? "" : "",
          answers.length ? "【これまでの確認質問への回答】" : "",
          ...answers.map((item) => `- ${item.question}: ${item.answer}`)
        ].filter((line) => line !== "").join("\n");

        const prompt = `あなたは介護施設の事故報告書を確認する補助AIです。

あなたの役割は「問いかけ・整理・見落としチェック」です。
原因分析や再発防止策をAIだけで決定してはいけません。職員や管理者が判断するための確認材料を整理してください。

【重要な扱い】
・氏名、住所、電話番号など、個人を直接特定できる情報は送信対象から除外しています。
・「要介護度」「認知症高齢者日常生活自立度」は、本人の状態を把握し、原因究明や再発防止を考えるための参考情報として使用してください。
・これらの情報から、書かれていない事実を推測・断定してはいけません。
・要介護度や認知症高齢者日常生活自立度だけを理由に、事故原因や対策を決めつけてはいけません。

【確認する観点】
1. 事故が起きる前の普段の状態はどうだったか
2. 最近、本人の状態に変化はなかったか
3. センサーマットなど、事故を防ぐための対策はしていたか
4. 機器・環境などのハード面での対策はできていたか
5. 職員の対応や見守りの状況はどうだったか
6. 原因分析に必要な情報が不足していないか
7. 再発防止策について、ハード面の対策、本人への確認・話し合い、必要に応じた家族への協力依頼などを検討するための情報が不足していないか
8. 明らかな矛盾や入力ミスがないか

本人との話し合いが適切でないケース（例：本人との意思疎通が難しいケース）では、その確認を無理に求めないでください。

【質問のルール】
・質問は本当に不足している情報がある場合だけ作成してください。
・最大10問です。10問に満たなくても構いません。
・すでにExcelに書かれている内容を、同じ意味で聞き直してはいけません。
・原則として選択式にしてください。自由記述は、選択肢では確認できない場合だけ使用してください。
・原因や再発防止策そのものを提案する質問ではなく、判断に必要な事実を確認する質問にしてください。

【出力形式】
質問が必要な場合は、必ず次のJSONだけを返してください。
{
  "questions": [
    {
      "question": "質問文",
      "options": ["選択肢1", "選択肢2", "選択肢3"],
      "allow_text": false
    }
  ]
}

質問が不要な場合は、次のJSONだけを返してください。
{
  "questions": []
}

JSON以外の文章は絶対に追加しないでください。

事故報告書の情報:
${reportText}`;

        const response = await fetch("https://generativelanguage.googleapis.com/v1beta/models/gemini-3.6-flash:generateContent", {
          method: "POST",
          headers: { "Content-Type": "application/json", "x-goog-api-key": env.GEMINI_API_KEY },
          body: JSON.stringify({
            contents: [{ parts: [{ text: prompt }] }],
            generationConfig: { temperature: 0.2, maxOutputTokens: 3000 }
          })
        });

        const data = await response.json();
        if (!response.ok) return Response.json({ success: false, error: data?.error?.message || "Gemini API request failed" }, { status: response.status });

        const text = data?.candidates?.[0]?.content?.parts?.map((part) => part.text || "").join("").trim();
        if (!text) return Response.json({ success: false, error: "Geminiから回答を取得できませんでした" }, { status: 502 });

        let parsed;
        try {
          parsed = JSON.parse(text.replace(/^```json\s*/i, "").replace(/\s*```$/i, "").trim());
        } catch {
          return Response.json({ success: true, questions: [], draft: text });
        }

        return Response.json({ success: true, questions: Array.isArray(parsed.questions) ? parsed.questions.slice(0, 10) : [], draft: parsed.draft || "" });
      } catch (error) {
        return Response.json({ success: false, error: error.message }, { status: 500 });
      }
    }

    if (url.pathname === "/api/reports" && request.method === "POST") {
      try {
        const report = await request.json();
        const required = ["submission_date", "report_number", "person_name", "occurrence_location", "severity", "cause_analysis", "recurrence_prevention"];
        for (const field of required) {
          if (!report[field]) return Response.json({ success: false, error: `${field} is required` }, { status: 400 });
        }
        const result = await env.DB.prepare(`INSERT INTO accident_reports (submission_date, report_number, person_name, occurrence_location, severity, cause_analysis, recurrence_prevention) VALUES (?, ?, ?, ?, ?, ?, ?)`).bind(report.submission_date, report.report_number, report.person_name, report.occurrence_location, report.severity, report.cause_analysis, report.recurrence_prevention).run();
        return Response.json({ success: true, id: result.meta.last_row_id });
      } catch (error) {
        return Response.json({ success: false, error: error.message }, { status: 500 });
      }
    }

    if (url.pathname === "/api/reports" && request.method === "GET") {
      const { results } = await env.DB.prepare(`SELECT * FROM accident_reports ORDER BY id DESC`).all();
      return Response.json({ success: true, reports: results });
    }

    if (env.ASSETS) return env.ASSETS.fetch(request);
    return new Response("Not Found", { status: 404 });
  }
};
