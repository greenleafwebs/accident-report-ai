export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    if (url.pathname === "/api/health") return Response.json({ ok: true });

    if (url.pathname === "/api/ai-check" && request.method === "POST") {
      try {
        if (!env.GEMINI_API_KEY) return Response.json({success:false,error:"GEMINI_API_KEY is not configured"},{status:500});
        const body=await request.json();
        const selected=Array.isArray(body.selected)?body.selected:[],texts=Array.isArray(body.texts)?body.texts:[],answers=Array.isArray(body.answers)?body.answers:[],confirmed=body.confirmed&&typeof body.confirmed==="object"?body.confirmed:{},round=Number(body.round||0);
        if(selected.length===0&&texts.length===0)return Response.json({success:false,error:"Excelの読み取り結果がありません"},{status:400});
        const excluded=new Set(["氏名","住所","性別","続柄","連絡先（電話番号）","医療機関名","診断名","本人、家族、関係先等への追加対応予定"]);
        const names=texts.filter(x=>x.field==="氏名").map(x=>String(x.value||"").trim()).filter(Boolean);
        const sanitize=v=>{let s=String(v||"");for(const n of names)s=s.split(n).join("[個人名]");return s.replace(/(?:0\d{1,4}[-ー]?\d{1,4}[-ー]?\d{3,4})/g,"[電話番号]");};
        const safeSelected=selected.filter(x=>!excluded.has(x.field)).map(x=>({field:x.field,value:sanitize(x.value)}));
        const safeTexts=texts.filter(x=>!excluded.has(x.field)).map(x=>({field:x.field,value:sanitize(x.value)}));
        const excelText=[...safeSelected.map(x=>"- "+x.field+": "+x.value),...safeTexts.map(x=>"- "+x.field+": "+x.value)].join("\n");
        const previousText=Object.entries(confirmed).filter(([,v])=>v).map(([k,v])=>"【確定済み："+k+"】\n"+sanitize(v)).join("\n\n");
        const answerText=answers.length?"【今回の確認質問への回答】\n"+answers.map(x=>"- 【"+(x.stage_label||x.stage||"該当項目")+"】 "+x.question+": "+x.answer).join("\n"):"";
        const reportText=["【Excelから読み取った情報】",excelText,previousText,answerText].filter(Boolean).join("\n\n");
        const stages=[
          {key:"incident_detail",label:"① 発生時状況、事故内容の詳細",purpose:"事故直前から発見時までの事実を整理",points:["事故直前の利用者の状況","事故時の職員の状況","環境・設備","事故直前の行動","事故そのもの","発見時の状況"]},
          {key:"response",label:"② 発生時の対応",purpose:"事故発見後の実際の対応を時系列で整理",points:["発見直後","安全確保・状態確認","バイタル等","受診・救急要請","医療機関・関係機関","見守り・経過観察"]},
          {key:"user_status",label:"③ 利用者の状況",purpose:"事故後に確認できた利用者の状態を整理",points:["身体状態","痛み・出血・外傷","意識・受け答え","バイタル等","受診・検査・診断","事故後の経過"]},
          {key:"cause",label:"④ 事故の原因分析",purpose:"本人・職員・環境要因を事実と分析に分けて整理",points:["本人要因","職員要因","環境要因","根拠となる事実","複数要因の関係"]},
          {key:"prevention",label:"⑤ 再発防止策",purpose:"原因分析を踏まえ実行可能な対策案を整理",points:["手順変更","環境変更","周知・教育","その他対応","担当・運用","評価時期・方法"]}
        ];
        const stageText=stages.map(s=>"【"+s.label+"】\n目的："+s.purpose+"\n確認観点：\n"+s.points.map((p,i)=>(i+1)+". "+p).join("\n")).join("\n\n");
        const rule=round===0?"①〜⑤を一度に点検し、事故報告として内容が具体的に伝わるように質問する。質問は最低5問、最大20問とする。Excelに既にある情報は重複して聞かない。①発生時の状況は特に具体化し、事故直前の行動、利用者の状態、職員の位置・対応、環境、事故の瞬間、発見時の状況など、第三者が場面をイメージできる情報を優先する。②③では対応と事故後の状態を具体化し、④の原因分析に必要な事実も確認する。④では①〜③を踏まえる。⑤では原因分析と確定内容を踏まえる。単純な一択だけでなく、複数の項目が同時に当てはまる場合はtype=multiの複数選択式にする。選択肢は介護現場で答えやすい具体的な内容にする。自由記述は必要な補足だけallow_text=trueにする。「不明」「確認できない」「記録なし」で回答できる内容を無理に回答させない。":"今回の回答を反映し、追加質問は原則0個として①〜⑤の文章を作成する。";
        const prompt="あなたは介護施設の事故報告書を整理する補助AIです。①〜⑤を一連の事故報告として扱います。\n\n"+stageText+"\n\n【重要ルール】\n・Excelの記載内容、今回の回答、確定内容だけを材料にする。\n・書かれていない事実を創作しない。\n・要介護度や認知症高齢者日常生活自立度だけで推測しない。\n・客観的で簡潔な文章にする。\n・②では家族への報告や上司への報告を記入・質問対象にしない。\n・④は事実と分析・可能性を分け、原因や責任を断定しない。\n・⑤の最終判断は職員・管理者が行う。\n・個人を直接特定できる情報を出力しない。\n\n【質問】\n"+rule+"\n\n【出力】\nround=0で不足がある場合はJSONだけで {\"questions\":[{\"stage\":\"incident_detail\",\"stage_label\":\"① 発生時状況、事故内容の詳細\",\"question\":\"質問文\",\"options\":[\"選択肢1\",\"選択肢2\"],\"allow_text\":true}]}\n情報が十分、またはround=1の場合はJSONだけで {\"questions\":[],\"drafts\":{\"① 発生時状況、事故内容の詳細\":\"文章\",\"② 発生時の対応\":\"文章\",\"③ 利用者の状況\":\"文章\",\"④ 事故の原因分析\":\"文章\",\"⑤ 再発防止策\":\"文章\"}}\n\n事故報告書の情報：\n"+reportText;
        const url="https://generativelanguage.googleapis.com/v1beta/models/gemini-3.6-flash:generateContent";
        const payload=JSON.stringify({contents:[{parts:[{text:prompt}]}],generationConfig:{temperature:0.2,maxOutputTokens:5000}});
        let response,data;const delays=[2000,4000,8000];
        for(let attempt=0;attempt<=delays.length;attempt++){response=await fetch(url,{method:"POST",headers:{"Content-Type":"application/json","x-goog-api-key":env.GEMINI_API_KEY},body:payload});data=await response.json();if(response.ok||response.status!==503||attempt===delays.length)break;await new Promise(r=>setTimeout(r,delays[attempt]));}
        if(!response.ok)return Response.json({success:false,error:data?.error?.message||"Gemini API request failed"},{status:response.status});
        const text=data?.candidates?.[0]?.content?.parts?.map(p=>p.text||"").join("").trim();if(!text)return Response.json({success:false,error:"Geminiから回答を取得できませんでした"},{status:502});
        let parsed;try{parsed=JSON.parse(text.replace(/^[`]{3}json\s*/i,"").replace(/\s*[`]{3}$/,"").trim());}catch{return Response.json({success:false,error:"GeminiのJSON形式を解析できませんでした"},{status:502});}
        return Response.json({success:true,questions:Array.isArray(parsed.questions)?parsed.questions.slice(0,20):[],drafts:parsed.drafts&&typeof parsed.drafts==="object"?parsed.drafts:{}});
      } catch(error) { return Response.json({success:false,error:error.message},{status:500}); }
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