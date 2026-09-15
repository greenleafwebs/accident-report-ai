export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    if (url.pathname === "/api/health") {
      return Response.json({ ok: true });
    }

    if (url.pathname === "/api/ai-check" && request.method === "POST") {
      try {
        if (!env.GEMINI_API_KEY) {
          return Response.json(
            { success: false, error: "GEMINI_API_KEY is not configured" },
            { status: 500 }
          );
        }

        const body = await request.json();
        const selected = Array.isArray(body.selected) ? body.selected : [];
        const texts = Array.isArray(body.texts) ? body.texts : [];

        if (selected.length === 0 && texts.length === 0) {
          return Response.json(
            { success: false, error: "Excelの読み取り結果がありません" },
            { status: 400 }
          );
        }

        const reportText = [
          "【選択された項目】",
          ...selected.map((item) => `- ${item.field}: ${item.value}`),
          "",
          "【自由記入欄】",
          ...texts.map((item) => `- ${item.field}: ${item.value}`)
        ].join("\n");

        const prompt = `あなたは介護施設の事故報告書を確認する補助AIです。

あなたの役割は「問いかけ・整理・見落としチェック」です。
原因分析や再発防止策を勝手に作成・決定してはいけません。
職員や管理者が判断するための確認材料だけを示してください。

次の事故報告書について、以下の観点で確認してください。
1. 重要な情報が不足していないか
2. 記載内容に矛盾や分かりにくい表現がないか
3. 原因分析や再発防止策について、追加で確認したほうがよいことがないか
4. 明らかな誤字・入力ミスがないか

回答は日本語で、簡潔にしてください。
問題がなければ「特に大きな見落としはありません」としてください。
確認事項がある場合は、1項目につき1つの短い箇条書きにしてください。
各項目は「・」で始めてください。
回答を途中で省略せず、すべての確認事項を最後まで出してください。
原因や再発防止策そのものをAIが提案するのではなく、「○○について確認してください」のような問いかけにしてください。

事故報告書の読み取り結果:
${reportText}`;

        const response = await fetch(
          "https://generativelanguage.googleapis.com/v1beta/models/gemini-3.6-flash:generateContent",
          {
            method: "POST",
            headers: {
              "Content-Type": "application/json",
              "x-goog-api-key": env.GEMINI_API_KEY
            },
            body: JSON.stringify({
              contents: [
                {
                  parts: [{ text: prompt }]
                }
              ],
              generationConfig: {
                temperature: 0.2,
                maxOutputTokens: 1600
              }
            })
          }
        );

        const data = await response.json();

        if (!response.ok) {
          const message = data?.error?.message || "Gemini API request failed";
          return Response.json(
            { success: false, error: message },
            { status: response.status }
          );
        }

        const text = data?.candidates?.[0]?.content?.parts
          ?.map((part) => part.text || "")
          .join("")
          .trim();

        if (!text) {
          return Response.json(
            { success: false, error: "Geminiから回答を取得できませんでした" },
            { status: 502 }
          );
        }

        return Response.json({ success: true, text });
      } catch (error) {
        return Response.json(
          { success: false, error: error.message },
          { status: 500 }
        );
      }
    }

    if (url.pathname === "/api/reports" && request.method === "POST") {
      try {
        const report = await request.json();

        const required = [
          "submission_date",
          "report_number",
          "person_name",
          "occurrence_location",
          "severity",
          "cause_analysis",
          "recurrence_prevention"
        ];

        for (const field of required) {
          if (!report[field]) {
            return Response.json(
              { success: false, error: `${field} is required` },
              { status: 400 }
            );
          }
        }

        const result = await env.DB.prepare(
          `INSERT INTO accident_reports
           (submission_date, report_number, person_name, occurrence_location, severity, cause_analysis, recurrence_prevention)
           VALUES (?, ?, ?, ?, ?, ?, ?)`
        )
          .bind(
            report.submission_date,
            report.report_number,
            report.person_name,
            report.occurrence_location,
            report.severity,
            report.cause_analysis,
            report.recurrence_prevention
          )
          .run();

        return Response.json({ success: true, id: result.meta.last_row_id });
      } catch (error) {
        return Response.json(
          { success: false, error: error.message },
          { status: 500 }
        );
      }
    }

    if (url.pathname === "/api/reports" && request.method === "GET") {
      const { results } = await env.DB.prepare(
        `SELECT * FROM accident_reports ORDER BY id DESC`
      ).all();
      return Response.json({ success: true, reports: results });
    }

    if (env.ASSETS) {
      return env.ASSETS.fetch(request);
    }

    return new Response("Not Found", { status: 404 });
  }
};
