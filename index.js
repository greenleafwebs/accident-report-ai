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

        // AIには直接の個人識別情報や健康情報を送らない。
        // ブラウザ側だけでなくWorker側でも除外して、送信前に二重チェックする。
        const excludedFields = new Set([
          "氏名",
          "住所",
          "連絡先（電話番号）",
          "医療機関名",
          "診断名",
          "利用者の状況",
          "本人、家族、関係先等への追加対応予定"
        ]);

        const sensitiveValues = texts
          .filter((item) => item.field === "氏名")
          .map((item) => String(item.value || "").trim())
          .filter(Boolean);

        const sanitizeValue = (value) => {
          let sanitized = String(value || "");
          for (const sensitiveValue of sensitiveValues) {
            sanitized = sanitized.split(sensitiveValue).join("[個人名]");
          }
          // 電話番号らしき文字列もAIへ送らない。
          sanitized = sanitized.replace(/(?:0\d{1,4}[-ー]?\d{1,4}[-ー]?\d{3,4})/g, "[電話番号]");
          return sanitized;
        };

        const safeSelected = selected
          .filter((item) => !excludedFields.has(item.field))
          .map((item) => ({ ...item, value: sanitizeValue(item.value) }));
        const safeTexts = texts
          .filter((item) => !excludedFields.has(item.field))
          .map((item) => ({ ...item, value: sanitizeValue(item.value) }));

        const reportText = [
          "【AIチェック対象：選択された項目】",
          ...safeSelected.map((item) => `- ${item.field}: ${item.value}`),
          "",
          "【AIチェック対象：自由記入欄】",
          ...safeTexts.map((item) => `- ${item.field}: ${item.value}`)
        ].join("\n");

        const prompt = `あなたは介護施設の事故報告書を確認する補助AIです。

あなたの役割は「問いかけ・整理・見落としチェック」です。
原因分析や再発防止策を勝手に作成・決定してはいけません。
職員や管理者が判断するための確認材料だけを示してください。

個人情報保護のため、氏名、住所、電話番号、医療機関名、診断名、利用者の健康状態などの個人を特定できる情報はAIチェック対象から除外しています。

次の事故報告書について、以下の観点で確認してください。
1. 重要な情報が不足していないか
2. 記載内容に矛盾や分かりにくい表現がないか
3. 原因分析や再発防止策について、追加で確認したほうがよいことがないか
4. 明らかな誤字・入力ミスがないか

回答は日本語で、簡潔にしてください。
確認事項は本当に重要なものだけに絞り、似た内容・関連する内容は1つにまとめてください。
確認事項は最大10項目までに絞ってください。
1項目は短く、原則として1文で完結させてください。
1つの確認事項を、複数の箇条書きに分解しないでください。
「誰が」「いつ」「どのように」など、複数の確認ポイントがあっても、関連する内容は1つの箇条書きにまとめてください。
例えば、「再発防止策の『見守りを徹底する』について、誰が、どのタイミングで、どのように行うか確認してください。」のように、1つの文章にまとめてください。
問題がなければ「特に大きな見落としはありません」としてください。
確認事項がある場合は、1項目につき1つの箇条書きにしてください。
各項目は必ず「・」で始めてください。
回答は必ず最後の項目まで完成させてください。文の途中で終わらせてはいけません。
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
                maxOutputTokens: 3000
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
