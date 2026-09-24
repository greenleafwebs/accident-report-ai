const form = document.getElementById("reportForm");
const result = document.getElementById("result");
const excelFile = document.getElementById("excelFile");
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
const draftResult = document.getElementById("draftResult");

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
  draftResult.innerHTML = "";
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
    const workbook = XLSX.read(buffer, { type: "array" });
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

aiCheck.addEventListener("click", () => {
  if (isAiProcessing) { cancelStage(currentStageIndex); return; }
  startStage(currentStageIndex);
});
skipStageButton.addEventListener("click", () => skipStage(currentStageIndex));

function loadSavedProgress(key) {
  if (!key) return {};
  try {
    const saved = JSON.parse(localStorage.getItem(key) || "{}");
    return saved && typeof saved === "object" ? saved : {};
  } catch {
    return {};
  }
}

function saveProgress() {
  if (!reportStorageKey) return;
  try {
    localStorage.setItem(reportStorageKey, JSON.stringify({ confirmedSections }));
  } catch {
    // 保存できない環境でも通常処理は継続する
  }
}

function updateStageStartControls(index) {
  currentStageIndex = index;
  const stage = stages[index];
  aiCheck.hidden = false;
  aiCheck.disabled = false;
  aiCheck.textContent = `${stage.label}を整理する`;

  const alreadyConfirmed = Boolean(confirmedSections[stage.label]);
  skipStageButton.hidden = false;
  skipStageButton.disabled = !alreadyConfirmed;
  skipStageButton.textContent = alreadyConfirmed
    ? `${stage.label}をスキップ（確認済み）`
    : `${stage.label}をスキップ（未確認）`;
}

function startStage(index) {
  currentStageIndex = index;
  currentRound = 0;
  questionArea.innerHTML = "";
  aiResult.textContent = "";
  const stage = stages[index];
  isAiProcessing = true;
  aiAbortController = new AbortController();
  aiCheck.disabled = false;
  aiCheck.hidden = false;
  aiCheck.textContent = `${stage.label}をキャンセルして次へ`;
  skipStageButton.hidden = true;
  requestAi(stage.key, 0, []);
}

function skipStage(index) {
  const stage = stages[index];
  if (completedStageIndexes.has(index)) return;
  questionArea.innerHTML = "";
  aiResult.textContent = "";
  if (confirmedSections[stage.label]) {
    renderConfirmedStage(stage, confirmedSections[stage.label]);
    completeStage(index, `${stage.label}をスキップしました。確認済みの内容を引き継ぎます。`);
  } else {
    completeStage(index, `${stage.label}をスキップしました。今回はこの項目をAIで整理せず、次へ進みます。`);
  }
}

function cancelStage(index) {
  if (!isAiProcessing || completedStageIndexes.has(index)) return;
  const stage = stages[index];
  isAiProcessing = false;
  if (aiAbortController) aiAbortController.abort();
  aiAbortController = null;
  questionArea.innerHTML = "";
  aiResult.textContent = "";
  aiCheck.hidden = true;
  aiCheck.disabled = true;
  skipStageButton.hidden = true;
  completeStage(index, `${stage.label}をキャンセルしました。今回はこの項目をAIで整理せず、次へ進みます。`);
}

function renderConfirmedStagefunction renderConfirmedStage(stage, text) {
  const stageBlock = document.createElement("section");
  stageBlock.className = "stage-result";
  const title = document.createElement("h3");
  title.textContent = stage.label;
  const status = document.createElement("p");
  status.textContent = "確認済みの内容を引き継ぎました。";
  const box = document.createElement("div");
  box.className = "draft-box";
  box.textContent = text;
  stageBlock.append(title, status, box);
  draftResult.appendChild(stageBlock);
}

function completeStage(index, messageText = "") {
  if (completedStageIndexes.has(index)) return;
  completedStageIndexes.add(index);
  isAiProcessing = false;
  aiAbortController = null;
  const next = index + 1;
  const title = stages[index].label;
  if (next < stages.length) {
    aiResult.innerHTML = "";
    const message = document.createElement("p");
    message.textContent = messageText || `${title}を確認済みにしました。次は${stages[next].label}です。`;
    const nextActions = document.createElement("div");    const nextActions = document.createElement("div");
    nextActions.className = "draft-actions";
    const nextButton = document.createElement("button");
    nextButton.type = "button";
    nextButton.textContent = `${stages[next].label}を整理する`;
    nextButton.addEventListener("click", () => startStage(next));
    const skipButton = document.createElement("button");
    skipButton.type = "button";
    skipButton.textContent = confirmedSections[stages[next].label]
      ? `${stages[next].label}をスキップ（確認済み）`
      : `${stages[next].label}をスキップ（未確認）`;
    skipButton.disabled = !confirmedSections[stages[next].label];
    skipButton.addEventListener("click", () => skipStage(next));
    nextActions.append(nextButton, skipButton);
    aiResult.append(message, nextActions);
  } else {
    aiResult.textContent = "①〜⑤の整理が完了しました。内容を確認してから最終的な報告書として使用してください。";
  }
}

async function requestAi(stage, round, answers) {
  aiResult.textContent = "AIが確認・作成中…";
  aiResult.scrollIntoView({ behavior: "smooth", block: "center" });
  questionArea.innerHTML = "";
  try {
    const response = await fetch("/api/ai-check", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ selected: latestSelected, texts: latestTexts, answers, confirmed: confirmedSections, stage, round }),
      signal: aiAbortController?.signal
    });
    const body = await response.json();
    if (!response.ok || !body.success) throw new Error(body.error || "AI処理に失敗しました");
    if (body.questions?.length) {
      renderQuestions(body.questions, stage);
      aiResult.textContent = "";
    } else {
      renderDraft(stage, body.draft || body.text || "AIから文章を取得できませんでした。");
      aiResult.textContent = "内容を確認してください。";
    }
  } catch (error) {
    if (error.name === "AbortError") return;
    isAiProcessing = false;
    aiAbortController = null;
    aiResult.textContent = `AI処理エラー：${error.message}`;
    aiResult.appendChild(document.createElement("br"));
    const retryButton = document.createElement("button");
    retryButton.type = "button";
    retryButton.textContent = `${stages.find((item) => item.key === stage)?.label || stage}を整理する`;
    retryButton.addEventListener("click", () => startStage(currentStageIndex));
    aiResult.appendChild(retryButton);
    const cancelButton = document.createElement("button");
    cancelButton.type = "button";
    cancelButton.textContent = `${stages.find((item) => item.key === stage)?.label || stage}をキャンセルして次へ`;
    cancelButton.addEventListener("click", () => {
      completeStage(currentStageIndex, `${stages[currentStageIndex].label}をキャンセルしました。今回はこの項目をAIで整理せず、次へ進みます。`);
    });
    aiResult.appendChild(document.createElement("br"));
    aiResult.appendChild(cancelButton);
  }
}

function renderQuestions(questions, stage) {
  const title = document.createElement("h3");
  title.textContent = `確認質問（${questions.length}問）`;
  questionArea.appendChild(title);
  questions.forEach((question, index) => {
    const card = document.createElement("div");
    card.className = "question-card";
    const prompt = document.createElement("p");
    prompt.textContent = `${index + 1}. ${question.question}`;
    card.appendChild(prompt);
    const options = Array.isArray(question.options) ? question.options : [];
    options.filter((option) => String(option).trim() !== "その他").forEach((option) => {
      const label = document.createElement("label");
      const radio = document.createElement("input");
      radio.type = "radio";
      radio.name = `question-${index}`;
      radio.value = option;
      label.append(radio, ` ${option}`);
      card.appendChild(label);
    });
    if (question.allow_text) {
      const label = document.createElement("label");
      const radio = document.createElement("input");
      radio.type = "radio";
      radio.name = `question-${index}`;
      radio.value = "その他";
      const input = document.createElement("input");
      input.type = "text";
      input.className = "question-other";
      input.placeholder = "内容を入力";
      input.dataset.other = "true";
      label.append(radio, " その他：", input);
      card.appendChild(label);
    }
    questionArea.appendChild(card);
  });
  const submit = document.createElement("button");
  submit.type = "button";
  submit.textContent = "回答して文章を作成";
  submit.addEventListener("click", async () => {
    const answers = [];
    document.querySelectorAll(".question-card").forEach((card, index) => {
      const selected = card.querySelector(`input[name="question-${index}"]:checked`);
      const other = card.querySelector("[data-other='true']");
      let value = selected?.value || "";
      if (value === "その他" && other?.value.trim()) value = `その他：${other.value.trim()}`;
      answers.push({ question: questions[index].question, answer: value });
    });
    if (answers.some((item) => !item.answer)) { aiResult.textContent = "未回答の質問があります。必要な質問に回答してください。"; return; }
    currentRound = 1;
    aiAbortController = new AbortController();
    isAiProcessing = true;
    await requestAi(stage, 1, answers);
  });
  questionArea.appendChild(submit);
}

function renderDraft(stage, text) {
  const stageBlock = document.createElement("section");
  stageBlock.className = "stage-result";
  const title = document.createElement("h3");
  title.textContent = stages.find((item) => item.key === stage)?.label || stage;
  const box = document.createElement("div");
  box.className = "draft-box";
  box.textContent = text;
  stageBlock.append(title, box);

  const actions = document.createElement("div");
  actions.className = "draft-actions";
  const ok = document.createElement("button");
  ok.type = "button";
  ok.textContent = "これでOK";
  const edit = document.createElement("button");
  edit.type = "button";
  edit.textContent = "手動で訂正";

  const confirm = (value) => {
    confirmedSections[title.textContent] = value;
    saveProgress();
    ok.disabled = true;
    edit.disabled = true;
    aiCheck.hidden = true;
    skipStageButton.hidden = true;
    completeStage(currentStageIndex);
  };

  ok.addEventListener("click", () => confirm(text));
  edit.addEventListener("click", () => {
    const textarea = document.createElement("textarea");
    textarea.className = "draft-edit";
    textarea.value = box.textContent;
    textarea.rows = Math.max(6, Math.min(14, text.split("\n").length + 2));
    const save = document.createElement("button");
    save.type = "button";
    save.textContent = "手動修正した内容でOK";
    save.addEventListener("click", () => {
      const value = textarea.value.trim();
      if (!value) { aiResult.textContent = "文章を入力してください。"; return; }
      box.textContent = value;
      textarea.replaceWith(box);
      save.remove();
      confirm(value);
    });
    box.replaceWith(textarea);
    actions.appendChild(save);
    edit.disabled = true;
  });
  actions.append(ok, edit);
  stageBlock.appendChild(actions);
  draftResult.appendChild(stageBlock);
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
