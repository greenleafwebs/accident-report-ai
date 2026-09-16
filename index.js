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
        const stage = body.stage || "incident_detail";
        const round = Number(body.round || 0);

        if (selected.length === 0 && texts.length === 0) return Response.json({ success: false, error: "Excelの読み取り結果がありません" }, { status: 400 });

        const excludedFields = new Set([
          "氏名",
          "住所",
          "性別",
          "続柄",
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
          answers.length ? "【今回の確認質問への回答】" : "",
          ...answers.map((item) => `- ${item.question}: ${item.answer}`)
        ].filter(Boolean).join("\n");

        const isIncidentStage = stage === "incident_detail";
        const prompt = isIncidentStage
          ? `あなたは介護施設の事故報告書を整理する補助AIです。

今回は「発生時状況、事故内容の詳細」だけを扱います。
目的は、第三者が読んだときに「いつ・どこで・誰に・何が起きたのか、その時の利用者・職員・ハードや環境の状況、発見時の状況」が分かる文章に整理することです。

【確認する6つの観点】
1. 利用者の状況：事故直前の状態、普段との違いなど
2. 職員の状況：事故時に職員がどこで何をしていたか
3. ハード・環境の状況：ベッド、車いす、床、手すり、センサーなどの状態
4. 事故発生直前の行動：利用者が何をしようとしていたか
5. 事故そのもの：いつ、どこで、どのように起きたか
6. 発見時の状況：誰が、どのような状態で発見したか

【重要】
・今回は事実関係の整理だけです。原因分析や再発防止策を決めたり提案したりしないでください。
・「なぜ事故が起きたのか」ではなく、「何が起きたのか」を確認してください。
・書かれていない事実を推測・創作してはいけません。
・「分からない」「確認できない」「記録なし」も有効な回答です。
・要介護度、認知症高齢者日常生活自立度は参考情報ですが、それだけを理由に状態や原因を推測してはいけません。
・質問は不足している事実がある場合だけ作成してください。
・質問は原則として一度に必要なものをまとめてください。最大6問です。
・選択肢を作り、「その他」を必ず選択肢に含め、自由入力できるようにしてください。
・今回の回答後に再度質問する場合でも、追加質問は1回だけです。roundが1の場合は、重大な不足がない限り質問せず、文章を作成してください。

【文章化のルール】
・Excelの記載内容と今回の回答だけを材料にしてください。
・時系列が分かるように整理してください。
・原因を断定する表現は避けてください。
・読みやすい自然な文章にしてください。

【出力形式】
roundが0で、情報不足がある場合は次のJSONだけを返してください。
{
  "questions": [
    {
      "question": "質問文",
      "options": ["選択肢1", "選択肢2", "その他"],
      "allow_text": true
    }
  ]
}

情報が十分、またはroundが1の場合は、次のJSONだけを返してください。
{
  "questions": [],
  "draft": "発生時状況、事故内容の詳細の文章"
}

JSON以外の文章は絶対に追加しないでください。

事故報告書の情報:
${reportText}`
          : `今回は「${stage}」の整理を行います。現在は開発中のため、まず発生時状況の整理を行う段階です。質問はせず、入力情報をそのまま整理してください。JSONのみで返してください。\n\n${reportText}`;

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

        return Response.json({
          success: true,
          questions: Array.isArray(parsed.questions) ? parsed.questions.slice(0, 6) : [],
          draft: parsed.draft || ""
        });
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
