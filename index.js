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
        const confirmed = body.confirmed && typeof body.confirmed === "object" ? body.confirmed : {};
        const stage = body.stage || "incident_detail";
        const round = Number(body.round || 0);

        if (selected.length === 0 && texts.length === 0) return Response.json({ success: false, error: "Excelの読み取り結果がありません" }, { status: 400 });

        const excludedFields = new Set(["氏名", "住所", "性別", "続柄", "連絡先（電話番号）", "医療機関名", "診断名", "本人、家族、関係先等への追加対応予定"]);
        const sensitiveValues = texts.filter((item) => item.field === "氏名").map((item) => String(item.value || "").trim()).filter(Boolean);
        const sanitizeValue = (value) => {
          let sanitized = String(value || "");
          for (const sensitiveValue of sensitiveValues) sanitized = sanitized.split(sensitiveValue).join("[個人名]");
          return sanitized.replace(/(?:0\d{1,4}[-ー]?\d{1,4}[-ー]?\d{3,4})/g, "[電話番号]");
        };

        const safeSelected = selected.filter((item) => !excludedFields.has(item.field)).map((item) => ({ field: item.field, value: sanitizeValue(item.value) }));
        const safeTexts = texts.filter((item) => !excludedFields.has(item.field)).map((item) => ({ field: item.field, value: sanitizeValue(item.value) }));
        const excelText = [
          ...safeSelected.map((item) => `- ${item.field}: ${item.value}`),
          ...safeTexts.map((item) => `- ${item.field}: ${item.value}`)
        ].join("\n");
        const previousText = Object.entries(confirmed).filter(([, value]) => value).map(([key, value]) => `【確定済み：${key}】\n${sanitizeValue(value)}`).join("\n\n");
        const answerText = answers.length ? `【今回の確認質問への回答】\n${answers.map((item) => `- ${item.question}: ${item.answer}`).join("\n")}` : "";
        const reportText = ["【Excelから読み取った情報】", excelText, previousText, answerText].filter(Boolean).join("\n\n");

        const stageInfo = {
          incident_detail: {
            title: "発生時状況、事故内容の詳細",
            purpose: "いつ・どこで・どのような状況で事故が起き、発見時にどのような状態だったかを事実として整理する",
            points: ["利用者の事故直前の状況", "事故時の職員の状況", "ハード・環境の状況", "事故発生直前の行動", "事故そのもの", "発見時の状況"],
            ask: "原因を推測せず、④事故の原因分析に必要となる重要な事実が不足していないか確認する。不足している場合だけ質問する"
          },
          response: {
            title: "事故発生後の対応",
            purpose: "事故発見後に職員が行った対応を、できるだけ時系列で整理する",
            points: ["発見直後の対応", "安全確保・状態確認", "バイタル等の確認", "受診・救急要請等の対応", "医療機関・関係機関への連絡", "その後の見守りや経過観察"],
            ask: "実際に行った対応だけを確認し、行っていない対応を推測しない。④事故の原因分析に必要となる発見直後の状況・状態確認・判断・対応の経過が不足していないかも確認する。不足している場合だけ質問する。家族への報告や上司への報告は、この項目では記入・質問の対象にしない"
          },
          user_status: {
            title: "利用者の状況",
            purpose: "事故後の利用者の身体・精神面や受診結果など、報告書に必要な状態を整理する",
            points: ["事故直後の身体状態", "痛み・出血・外傷等の確認", "意識・受け答え・普段との違い", "バイタル等の確認結果", "受診・検査・診断の結果", "事故後の経過や現在の状態"],
            ask: "記録や回答にない状態を推測せず、確認できた事実だけを整理する。④事故の原因分析に必要となる事故前後の利用者の状態や普段との違いが不足していないかも確認し、不足している場合だけ質問する"
          },
          cause: {
            title: "事故の原因分析",
            purpose: "本人要因・職員要因・環境要因の情報を整理し、原因分析のたたき台を作る",
            points: ["本人要因", "職員要因", "環境要因", "それぞれを裏付ける事実", "複数要因の関係"],
            ask: "①発生時状況、②発生時の対応、③利用者の状況を確認したうえで、原因分析に必要な事実が不足していないか点検する。不足している場合は質問して補う。そのうえで本人要因・職員要因・環境要因などを、確認できた事実と分析・可能性を分けて整理する。AIだけで最終判断しない"
          },
          prevention: {
            title: "再発防止策",
            purpose: "原因分析と確定内容を踏まえ、実行可能な再発防止策の案を整理する",
            points: ["手順の変更", "環境の変更", "職員への周知・教育", "その他の対応", "実施担当や運用方法", "評価時期", "評価結果"],
            ask: "既に行った対策と今後行う対策を区別し、実際に施設で実施できる案として整理する。最終的な対策は職員・管理者が判断する"
          }
        };
        const info = stageInfo[stage] || stageInfo.incident_detail;
        const questionAllowed = stage === "incident_detail" || stage === "response" || stage === "user_status" || stage === "cause";

        const prompt = `あなたは介護施設の事故報告書を整理する補助AIです。\n\n今回は「${info.title}」を扱います。\n目的：${info.purpose}\n\n【確認・整理する観点】\n${info.points.map((point, index) => `${index + 1}. ${point}`).join("\n")}\n\n【重要】\n・${info.ask}\n・Excelの記載内容と、今回の回答、すでに確定した内容だけを材料にしてください。\n・要介護度、認知症高齢者日常生活自立度は参考情報ですが、それだけを理由に状態や原因を推測してはいけません。\n・書かれていない事実を創作しないでください。\n・「不明」「確認できない」「記録なし」は有効な情報です。\n・第三者が読んでも分かる、客観的で簡潔な事故報告書の文章にしてください。\n・情報を単純な短文の羅列にせず、時系列や前後関係が自然につながる文章にしてください。\n・「その際」「その後」などの接続表現は、実際の関係が確認できる場合だけ使用してください。\n・同じ表現の繰り返しを避けてください。\n\n【後の原因分析を見据えた情報収集】\n・①発生時状況、②発生時の対応、③利用者の状況では、単に文章を整えるだけでなく、④事故の原因分析に必要となる重要な事実が不足していないか確認してください。\n・不足がある場合だけ質問してください。質問は原因を推測するためではなく、原因分析の材料となる事実を確認するために行ってください。\n・①～③で原因を断定したり、原因や責任を決めつけたりしてはいけません。\n・④事故の原因分析では、①～③の内容を必ず確認し、原因分析に必要な事実が不足していれば質問してください。情報が揃ったら、確認できた事実と分析・可能性を分けて整理してください。\n\n【質問について】\n${questionAllowed ? `roundが0の場合、今回の段階で本当に不足している情報を点検してください。①～③では④の原因分析に必要な重要事実が不足している場合だけ質問してください。④では①～③を確認したうえで、原因分析を行うために必要な事実が不足している場合だけ質問してください。情報が十分なら質問は0個にしてください。不足がある場合だけ、本当に必要な質問を必要な数だけ作成してください。質問数に固定の下限・目標はありません。情報不足が多い場合は最大10問まで作成して構いませんが、10問にするために不要な質問を追加してはいけません。質問は重要度の高いものから並べてください。roundが1の場合は、重大な不足がない限り追加質問をせず、文章を作成してください。「不明」「確認できない」「記録なし」と回答できる内容について、無理に回答を求める質問にしないでください。選択肢を作り、allow_textがtrueの場合は選択肢に「その他」を入れず、自由入力欄だけを用意してください。` : `質問は行わず、与えられた情報から整理案を作成してください。`}\n\n【原因・再発防止について】\n${stage === "cause" ? "原因分析はAIによる整理案です。①～③の確定内容を材料にし、本人・職員・環境の要因を、確認できた事実と分析・可能性に分けて表現してください。原因を断定せず、最終判断は職員・管理者が行います。" : ""}\n${stage === "prevention" ? "再発防止策はAIによる案です。確定した原因分析と確定内容に沿って、具体的で実行可能な案にしてください。施設側の最終判断が必要な事項は断定しないでください。" : ""}\n\n【出力形式】\n${questionAllowed ? `roundが0で情報不足がある場合はJSONだけを返してください：\n{\"questions\":[{\"question\":\"質問文\",\"options\":[\"選択肢1\",\"選択肢2\"],\"allow_text\":true}]}\n情報が十分、またはroundが1の場合はJSONだけを返してください：\n{\"questions\":[],\"draft\":\"${info.title}の自然な文章\"}` : `JSONだけを返してください：\n{\"questions\":[],\"draft\":\"${info.title}の整理案\"}`}\n\n事故報告書の情報：\n${reportText}`;

        const response = await fetch("https://generativelanguage.googleapis.com/v1beta/models/gemini-3.6-flash:generateContent", {
          method: "POST",
          headers: { "Content-Type": "application/json", "x-goog-api-key": env.GEMINI_API_KEY },
          body: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }], generationConfig: { temperature: 0.2, maxOutputTokens: 3000 } })
        });
        const data = await response.json();
        if (!response.ok) return Response.json({ success: false, error: data?.error?.message || "Gemini API request failed" }, { status: response.status });
        const text = data?.candidates?.[0]?.content?.parts?.map((part) => part.text || "").join("").trim();
        if (!text) return Response.json({ success: false, error: "Geminiから回答を取得できませんでした" }, { status: 502 });
        let parsed;
        try { parsed = JSON.parse(text.replace(/^```json\s*/i, "").replace(/\s*```$/i, "").trim()); }
        catch { return Response.json({ success: true, questions: [], draft: text }); }
        return Response.json({ success: true, questions: Array.isArray(parsed.questions) ? parsed.questions.slice(0, 10) : [], draft: parsed.draft || "" });
      } catch (error) {
        return Response.json({ success: false, error: error.message }, { status: 500 });
      }
    }

    if (url.pathname === "/api/reports" && request.method === "POST") {
      try {
        const report = await request.json();
        const required = ["submission_date", "report_number", "person_name", "occurrence_location", "severity", "cause_analysis", "recurrence_prevention"];
        for (const field of required) if (!report[field]) return Response.json({ success: false, error: `${field} is required` }, { status: 400 });
        const result = await env.DB.prepare(`INSERT INTO accident_reports (submission_date, report_number, person_name, occurrence_location, severity, cause_analysis, recurrence_prevention) VALUES (?, ?, ?, ?, ?, ?, ?)`).bind(report.submission_date, report.report_number, report.person_name, report.occurrence_location, report.severity, report.cause_analysis, report.recurrence_prevention).run();
        return Response.json({ success: true, id: result.meta.last_row_id });
      } catch (error) { return Response.json({ success: false, error: error.message }, { status: 500 }); }
    }

    if (url.pathname === "/api/reports" && request.method === "GET") {
      const { results } = await env.DB.prepare(`SELECT * FROM accident_reports ORDER BY id DESC`).all();
      return Response.json({ success: true, reports: results });
    }

    if (env.ASSETS) return env.ASSETS.fetch(request);
    return new Response("Not Found", { status: 404 });
  }
};