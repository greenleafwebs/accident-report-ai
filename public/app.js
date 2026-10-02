const form = document.getElementById("reportForm");
const result = document.getElementById("result");
const excelFile = document.getElementById("excelFile");
const excelFileLabel = document.getElementById("excelFileLabel");
const fileName = document.getElementById("selectedFileName");
const readExcel = document.getElementById("readExcel");
const excelResult = document.getElementById("excelResult");
const preview = document.getElementById("preview");
const aiArea = document.getElementById("aiArea");
const checkedItems = document.getElementById("checkedItems");
const textItems = document.getElementById("textItems");
const aiCheck = document.getElementById("aiCheck");
const skipStageButton = document.getElementById("skipStage");
const aiResult = document.getElementById("aiResult");
const questionArea = document.getElementById("questionArea");

const selectionFields = ["第○報","事故状況の程度","要介護度","認知症高齢者日常生活自立度","発生場所","事故の種別","受診方法","診断内容","連絡した関係機関 (連絡した場合のみ)"];
const textFields = ["発生日時","発生時状況、事故内容の詳細","発生時の対応","医療機関名","診断名","検査、処置等の概要","利用者の状況","本人、家族、関係先等への追加対応予定","原因分析","再発防止策","その他"];
const stages = [
  { key: "incident_detail", label: "① 発生時状況、事故内容の詳細" },
  { key: "response", label: "② 発生時の対応" },
  { key: "user_status", label: "③ 利用者の状況" },
  { key: "cause", label: "④ 事故の原因分析" },
  { key: "prevention", label: "⑤ 再発防止策" }
];

let latestSelected = [];
let latestTexts = [];
let currentStageIndex = 0;
let currentRound = 0;
let confirmedSections = {};
let completedStageIndexes = new Set();
let reportStorageKey = "";
let aiAbortController = null;
let isAiProcessing = false;
let currentStageBlock = null;
let originalWorkbook = null;
let originalFileBuffer = null;
let originalFileName = "";

if (excelFileLabel && excelFile) {
  excelFileLabel.addEventListener("click", (event) => {
    event.preventDefault();
    excelFile.click();
  });
}

excelFile.addEventListener("change", () => {
  const file = excelFile.files[0];
  fileName.textContent = file ? file.name : "ファイルが選択されていません。";
});

readExcel.addEventListener("click", async () => {
  const file = excelFile.files[0];
  if (!file) { excelResult.textContent = "Excelファイルを選択してください。"; return; }
  excelResult.textContent = "読み込み中…";
  checkedItems.innerHTML = "";
  textItems.innerHTML = "";
  questionArea.innerHTML = "";
  aiResult.textContent = "";
  aiCheck.disabled = true;
  preview.hidden = true;
  aiArea.hidden = true;
  preview.open = true;
  currentStageIndex = 0;
  currentRound = 0;
  completedStageIndexes = new Set();
  reportStorageKey = `accident-report-ai:${file.name}:${file.size}:${file.lastModified}`;
  const saved = loadSavedProgress(reportStorageKey);
  confirmedSections = saved.confirmedSections || {};
  try {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array", cellStyles: true });
    originalWorkbook = workbook;
    originalFileBuffer = buffer;
    originalFileName = file.name;
    const selected = [];
    const texts = [];
    for (const sheetName of workbook.SheetNames) {
      const sheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
      for (let rowIndex = 0; rowIndex < rows.length; rowIndex++) {
        const row = rows[rowIndex];
        for (const fieldName of selectionFields) {
          for (const fieldColumn of findFieldColumns(row, fieldName)) {
            getCheckedValues(row, fieldColumn).forEach((value) => selected.push({ field: fieldName, value, sheet: sheetName, row: rowIndex + 1 }));
          }
        }
        for (const fieldName of textFields) {
          for (const fieldColumn of findFieldColumns(row, fieldName)) {
            const value = isMultiRowTextField(fieldName) ? getMultiRowTextValue(rows, rowIndex, fieldColumn, fieldName) : getTextValue(row, fieldColumn, fieldName);
            if (value) texts.push({ field: fieldName, value, sheet: sheetName, row: rowIndex + 1 });
          }
        }
      }
    }
    latestSelected = selected;
    latestTexts = texts;
    renderReadOnlyFields(checkedItems, selected);
    renderReadOnlyFields(textItems, texts);
    preview.hidden = false;
    preview.open = true;
    aiArea.hidden = false;
    aiCheck.hidden = false;
    aiCheck.disabled = selected.length === 0 && texts.length === 0;
    updateStageStartControls(0);
    excelResult.textContent = saved.confirmedSections
      ? `読み込み完了：${workbook.SheetNames.length}シート（確認済み内容あり）`
      : `読み込み完了：${workbook.SheetNames.length}シート`;
  } catch (error) {
    excelResult.textContent = `読み込みエラー：${error.message}`;
  }
});

function findFieldColumns(row, fieldName) {
  const columns = [];
  row.forEach((cell, index) => {
    const value = String(cell ?? "").trim();
    const normalized = value.replace(/\s+/g, "");
    if (fieldName === "その他") {
      if (/^9その他(?:特記すべき事項)?$/.test(normalized)) columns.push(index);
      return;
    }
    if (value === fieldName || value.includes(fieldName)) columns.push(index);
  });
  return columns;
}

function getCheckedValues(row, fieldColumn) {
  const values = [];
  for (let index = fieldColumn + 1; index < row.length; index++) {
    const value = String(row[index] ?? "").trim();
    if (!value) continue;
    if (value.includes("■")) {
      const right = String(row[index + 1] ?? "").trim();
      if (right && !right.includes("□") && !right.includes("■") && !right.includes("☐") && !right.includes("☑")) values.push(normalizeSelectedValue(right));
    }
  }
  return [...new Set(values)];
}

function normalizeSelectedValue(value) { return String(value ?? "").replace(/[\r\n]+/g, " ").replace(/\s+/g, " ").trim(); }
function getTextValue(row, fieldColumn, fieldName) {
  const values = [];
  for (let index = fieldColumn + 1; index < row.length; index++) {
    const value = String(row[index] ?? "").trim();
    if (!value || isCheckbox(value) || isPlaceholderText(value)) continue;
    if (isKnownFieldLabel(value, fieldName)) break;
    values.push(value);
  }
  return cleanTextValues(values, fieldName);
}
function isMultiRowTextField(fieldName) { return fieldName === "原因分析" || fieldName === "再発防止策"; }
function getMultiRowTextValue(rows, startRow, fieldColumn, fieldName) {
  for (let rowIndex = startRow + 1; rowIndex < Math.min(rows.length, startRow + 3); rowIndex++) {
    const row = rows[rowIndex];
    for (let index = fieldColumn + 1; index < row.length; index++) {
      const value = String(row[index] ?? "").trim();
      if (!value || isCheckbox(value) || isPlaceholderText(value)) continue;
      if (isKnownFieldLabel(value, fieldName)) return "";
      return cleanTextValues([value], fieldName);
    }
  }
  return "";
}
function isPlaceholderText(value) { return value.includes("できるだけ具体的に記載すること"); }
function cleanTextValues(values, fieldName) { return [...new Set(values.map((value) => value.replace(/\s+/g, " ").trim()).filter(Boolean).filter((value) => value !== fieldName).filter((value) => !isPlaceholderText(value)))].join(" ").trim(); }
function isCheckbox(value) { return /^[□■☐☑]+$/.test(value); }
function isKnownFieldLabel(value, currentFieldName) { return [...selectionFields, ...textFields].some((fieldName) => fieldName !== currentFieldName && value === fieldName); }
function normalizeFieldName(fieldName) { return fieldName.replace(/[：:]+$/g, "").trim(); }

function renderReadOnlyFields(container, items) {
  if (items.length === 0) { container.textContent = "読み取りできる内容は見つかりませんでした。"; return; }
  const grouped = new Map();
  items.forEach((item) => {
    if (!grouped.has(item.field)) grouped.set(item.field, []);
    grouped.get(item.field).push(item.value);
  });
  grouped.forEach((values, fieldName) => {
    const wrapper = document.createElement("div");
    wrapper.className = "read-item";
    const label = document.createElement("span");
    label.className = "read-item-label";
    label.textContent = `${normalizeFieldName(fieldName)}：`;
    const value = document.createElement("span");
    value.className = "read-value";
    value.textContent = [...new Set(values)].join("、");
    wrapper.append(label, value);
    container.appendChild(wrapper);
  });
}

function loadSavedProgress(key){if(!key)return{};try{const saved=JSON.parse(localStorage.getItem(key)||"{}");return saved&&typeof saved==="object"?saved:{};}catch{return{};}}
function updateStageStartControls(){aiCheck.hidden=false;aiCheck.disabled=false;aiCheck.textContent="🤖 5項目をまとめて整理する";}
aiCheck.addEventListener("click",()=>{if(isAiProcessing)return;startAllStages();});
function startAllStages(){questionArea.innerHTML="";aiResult.textContent="AIが①〜⑤の不足情報を確認中…";isAiProcessing=true;aiAbortController=new AbortController();aiCheck.disabled=true;requestAiAll(0,[]);}
function requestAiAll(round,answers){
  aiResult.textContent=round===0?"AIが①〜⑤の不足情報を確認中…":"AIが①〜⑤の文章を作成中…";
  fetch("/api/ai-check",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({selected:latestSelected,texts:latestTexts,answers:answers,confirmed:confirmedSections,stage:"all",round:round}),signal:aiAbortController?.signal}).then(async r=>{const b=await r.json();if(!r.ok||!b.success)throw new Error(b.error||"AI処理に失敗しました");return b;}).then(b=>{isAiProcessing=false;aiAbortController=null;aiCheck.disabled=false;aiCheck.textContent="🤖 5項目をまとめて整理する";if(b.questions?.length){renderAllQuestions(b.questions);aiResult.textContent="不足している内容だけ回答してください。";}else{renderAllDrafts(b.drafts||{});aiResult.textContent="①〜⑤の文章を作成しました。内容を確認してください。";}}).catch(err=>{if(err.name==="AbortError")return;isAiProcessing=false;aiAbortController=null;aiCheck.disabled=false;aiResult.textContent="AI処理エラー："+err.message;const b=document.createElement("button");b.textContent="もう一度確認する";b.onclick=startAllStages;aiResult.appendChild(document.createElement("br"));aiResult.appendChild(b);});
}
function renderAllQuestions(questions){
  questionArea.innerHTML="";
  const h=document.createElement("h3");
  h.textContent="確認質問（"+questions.length+"問）";
  questionArea.appendChild(h);
  questions.forEach((q,i)=>{
    const c=document.createElement("div");
    c.className="question-card";
    if(q.stage_label){
      const st=document.createElement("p");
      st.textContent="【"+q.stage_label+"】";
      c.appendChild(st);
    }
    const p=document.createElement("p");
    p.textContent=(i+1)+". "+q.question;
    c.appendChild(p);
    const type=q.type==="multi"?"checkbox":"radio";
    const name="allq-"+i;
    (Array.isArray(q.options)?q.options:[]).filter(o=>String(o).trim()!=="その他").forEach(o=>{
      const l=document.createElement("label"),x=document.createElement("input");
      x.type=type;x.name=name;x.value=o;
      l.append(x," "+o);
      c.appendChild(l);
    });
    if(q.allow_text){
      const l=document.createElement("label"),x=document.createElement("input"),t=document.createElement("input");
      x.type=type;x.name=name;x.value="その他";
      t.type="text";t.className="question-other";t.placeholder="内容を入力";t.dataset.other="true";
      l.append(x," その他：",t);
      c.appendChild(l);
    }
    questionArea.appendChild(c);
  });
  const submit=document.createElement("button");
  submit.textContent="回答して5項目の文章を作成";
  submit.onclick=()=>{
    const answers=[];
    questionArea.querySelectorAll(".question-card").forEach((c,i)=>{
      const inputs=[...c.querySelectorAll("input[name=\"allq-"+i+"\"]:checked")];
      const other=c.querySelector("[data-other=\"true\"]");
      let values=inputs.map(x=>x.value);
      if(values.includes("その他")&&other?.value.trim()) values=values.map(v=>v==="その他"?"その他："+other.value.trim():v);
      answers.push({
        stage:questions[i].stage,
        stage_label:questions[i].stage_label,
        question:questions[i].question,
        answer:values.join("、")
      });
    });
    if(answers.some(x=>!x.answer)){
      aiResult.textContent="未回答の質問があります。必要な質問に回答してください。";
      return;
    }
    aiCheck.disabled=true;
    isAiProcessing=true;
    aiAbortController=new AbortController();
    requestAiAll(1,answers);
  };
  questionArea.appendChild(submit);
}
function renderAllDrafts(drafts){
  questionArea.innerHTML="";const labels=["① 発生時状況、事故内容の詳細","② 発生時の対応","③ 利用者の状況","④ 事故の原因分析","⑤ 再発防止策"];
  labels.forEach(label=>{const block=document.createElement("section"),title=document.createElement("h3"),box=document.createElement("div"),actions=document.createElement("div");block.className="draft-result-block stage-result";title.textContent=label;box.className="draft-box";box.textContent=String(drafts[label]||"");actions.className="draft-actions";
    const ok=document.createElement("button"),edit=document.createElement("button");ok.textContent="これでOK";edit.textContent="手動で訂正";
    const confirm=v=>{confirmedSections[label]=v;saveProgress();ok.disabled=true;edit.disabled=true;showDownloadButton();};ok.onclick=()=>confirm(box.textContent.trim());
    edit.onclick=()=>{const ta=document.createElement("textarea"),save=document.createElement("button");ta.className="draft-edit";ta.value=box.textContent;ta.rows=8;save.textContent="手動修正した内容でOK";save.onclick=()=>{const v=ta.value.trim();if(!v){aiResult.textContent="文章を入力してください。";return;}box.textContent=v;ta.replaceWith(box);save.remove();confirm(v);};box.replaceWith(ta);actions.appendChild(save);edit.disabled=true;};
    actions.append(ok,edit);block.append(title,box,actions);questionArea.appendChild(block);});showDownloadButton();
}
function showDownloadButton(){const labels=["① 発生時状況、事故内容の詳細","② 発生時の対応","③ 利用者の状況","④ 事故の原因分析","⑤ 再発防止策"];const all=labels.every(x=>Object.prototype.hasOwnProperty.call(confirmedSections,x));aiResult.innerHTML="";const p=document.createElement("p");p.textContent=all?"①〜⑤の確認が完了しました。":"各項目の内容を確認して「これでOK」または「手動で訂正」を行ってください。";aiResult.appendChild(p);if(all){const b=document.createElement("button");b.textContent="📥 完成したExcelをダウンロード";b.onclick=downloadCompletedExcel;aiResult.appendChild(b);}}
function saveProgress(){if(!reportStorageKey)return;try{localStorage.setItem(reportStorageKey,JSON.stringify({confirmedSections}));}catch{}}
async function downloadCompletedExcel(){
  const labels=["① 発生時状況、事故内容の詳細","② 発生時の対応","③ 利用者の状況","④ 事故の原因分析","⑤ 再発防止策"];
  if(!originalFileBuffer){aiResult.textContent="元のExcelファイルが見つかりません。もう一度Excelを読み込んでください。";return;}
  if(!labels.every(label=>Object.prototype.hasOwnProperty.call(confirmedSections,label))){aiResult.textContent="①〜⑤すべての内容を「これでOK」または「手動で訂正」してからダウンロードしてください。";return;}
  const targets={"① 発生時状況、事故内容の詳細":"D29","② 発生時の対応":"D29","③ 利用者の状況":"D38","④ 事故の原因分析":"E44","⑤ 再発防止策":"E46"};
  try {
    const zip = await JSZip.loadAsync(originalFileBuffer);
    const sheetPath = "xl/worksheets/sheet1.xml";
    const sheetXml = await zip.file(sheetPath).async("string");
    const escapeXml = (value) => String(value ?? "")
      .replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;")
      .replace(/"/g,"&quot;").replace(/'/g,"&apos;");
    let updatedXml = sheetXml;
    Object.entries(targets).forEach(([label,address]) => {
      const value = escapeXml(String(confirmedSections[label] ?? "").trim());
      const pattern = new RegExp('<c\\b([^>]*\\br="' + address + '"[^>]*)>([\\s\\S]*?)</c>|<c\\b([^>]*\\br="' + address + '"[^>]*)\\/>');
      if(!pattern.test(updatedXml)) throw new Error(address+"セルが見つかりません。");
      updatedXml = updatedXml.replace(pattern, (match, attrs1, inner, attrs2) => {
        const attrs = attrs1 || attrs2;
        const cleanAttrs = attrs.replace(/\\s+t="[^"]*"/g,"");
        return '<c' + cleanAttrs + ' t="inlineStr"><is><t xml:space="preserve">' + value + '</t></is></c>';
      });
    });
    zip.file(sheetPath, updatedXml);
    const output = await zip.generateAsync({type:"blob",compression:"DEFLATE"});
    const url = URL.createObjectURL(output);
    const link = document.createElement("a");
    const baseName=originalFileName.replace(/\\.[^.]+$/,"")||"事故報告書";
    link.href=url;
    link.download=baseName+"_完成版.xlsx";
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
    aiResult.textContent="完成版Excelをダウンロードしました。元ファイルの書式を維持したまま、①〜⑤を書き込んでいます。";
  } catch(error) {
    aiResult.textContent="Excel書き戻しエラー："+error.message;
  }
}

if (form) {
  form.addEventListener("submit", async (event) => {
    event.preventDefault();
    result.textContent = "保存中…";
    const data = Object.fromEntries(new FormData(form).entries());
    try {
      const response = await fetch("/api/reports", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(data) });
      const body = await response.json();
      if (!response.ok || !body.success) throw new Error(body.error || "保存に失敗しました");
      result.textContent = `保存しました（ID: ${body.id}）`;
      form.reset();
    } catch (error) {
      result.textContent = `保存エラー：${error.message}`;
    }
  });
}
