const form = document.getElementById("reportForm");
const result = document.getElementById("result");
const excelFile = document.getElementById("excelFile");
const fileName = document.getElementById("selectedFileName");
const readExcel = document.getElementById("readExcel");
const excelResult = document.getElementById("excelResult");
const preview = document.getElementById("preview");
const checkedItems = document.getElementById("checkedItems");
const textItems = document.getElementById("textItems");
const aiCheck = document.getElementById("aiCheck");
const aiResult = document.getElementById("aiResult");
const questionArea = document.getElementById("questionArea");
const draftResult = document.getElementById("draftResult");

const selectionFields = ["第○報","事故状況の程度","要介護度","認知症高齢者日常生活自立度","発生場所","事故の種別","受診方法","診断内容","連絡した関係機関 (連絡した場合のみ)"];
const textFields = ["発生日時","発生時状況、事故内容の詳細","発生時の対応","医療機関名","診断名","検査、処置等の概要","利用者の状況","本人、家族、関係先等への追加対応予定","原因分析","再発防止策","その他"];

let latestSelected = [];
let latestTexts = [];

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
    renderSelectedFields(checkedItems, selected);
    renderTextFields(textItems, texts);
    preview.hidden = false;
    aiCheck.disabled = selected.length === 0 && texts.length === 0;
    excelResult.textContent = `読み込み完了：${workbook.SheetNames.length}シート`;
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
      if (right && !right.includes("□") && !right.includes("■") && !right.includes("☐") && !right.includes("☑")) {
        values.push(normalizeSelectedValue(right));
      }
    }
  }
  return [...new Set(values)];
}

function normalizeSelectedValue(value) {
  return String(value ?? "").replace(/[\r\n]+/g, " ").replace(/\s+/g, " ").trim();
}

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
function cleanTextValues(values, fieldName) {
  const cleaned = values.map((value) => value.replace(/\s+/g, " ").trim()).filter(Boolean).filter((value) => value !== fieldName).filter((value) => !isPlaceholderText(value));
  return [...new Set(cleaned)].join(" ").trim();
}
function isCheckbox(value) { return /^[□■☐☑]+$/.test(value); }
function isKnownFieldLabel(value, currentFieldName) {
  const allFields = [...selectionFields, ...textFields];
  return allFields.some((fieldName) => fieldName !== currentFieldName && value === fieldName);
}
function normalizeFieldName(fieldName) { return fieldName.replace(/[：:]+$/g, "").trim(); }

function renderSelectedFields(container, items) { renderReadOnlyFields(container, items); }
function renderTextFields(container, items) { renderReadOnlyFields(container, items); }

function renderReadOnlyFields(container, items) {
  if (items.length === 0) {
    container.textContent = "読み取りできる内容は見つかりませんでした。";
    return;
  }
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
    const readValue = document.createElement("span");
    readValue.className = "read-value";
    readValue.textContent = [...new Set(values)].join("、");
    wrapper.append(label, readValue);
    container.appendChild(wrapper);
  });
}

aiCheck.addEventListener("click", async () => {
  await requestAi({ selected: latestSelected, texts: latestTexts, answers: [] });
});

async function requestAi(payload) {
  aiResult.textContent = "AIが確認・作成中…";
  questionArea.innerHTML = "";
  draftResult.innerHTML = "";
  aiCheck.disabled = true;
  try {
    const response = await fetch("/api/ai-check", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload) });
    const body = await response.json();
    if (!response.ok || !body.success) throw new Error(body.error || "AI処理に失敗しました");
    if (body.questions?.length) {
      renderQuestions(body.questions);
      aiResult.textContent = "不足している情報があります。選択式を中心に回答してください。";
    } else {
      renderDraft(body.draft || body.text || "AIから回答を取得できませんでした。");
      aiResult.textContent = "確認が完了しました。必要に応じて内容を修正してから最終確認してください。";
    }
  } catch (error) {
    aiResult.textContent = `AI処理エラー：${error.message}`;
  } finally {
    aiCheck.disabled = false;
  }
}

function renderQuestions(questions) {
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
    options.forEach((option) => {
      const label = document.createElement("label");
      const radio = document.createElement("input");
      radio.type = "radio";
      radio.name = `question-${index}`;
      radio.value = option;
      label.append(radio, ` ${option}`);
      card.appendChild(label);
    });
    if (question.allow_text) {
      const other = document.createElement("input");
      other.type = "text";
      other.className = "question-other";
      other.placeholder = "選択肢にない場合のみ入力";
      other.dataset.other = "true";
      card.appendChild(other);
    }
    questionArea.appendChild(card);
  });
  const submit = document.createElement("button");
  submit.type = "button";
  submit.textContent = "回答して報告書を作成";
  submit.addEventListener("click", async () => {
    const answers = [];
    document.querySelectorAll(".question-card").forEach((card, index) => {
      const selected = card.querySelector(`input[name="question-${index}"]:checked`);
      const other = card.querySelector("[data-other='true']");
      const value = other?.value.trim() || selected?.value || "";
      answers.push({ question: questions[index].question, answer: value });
    });
    if (answers.some((item) => !item.answer)) {
      aiResult.textContent = "未回答の質問があります。必要な質問に回答してください。";
      return;
    }
    await requestAi({ selected: latestSelected, texts: latestTexts, answers });
  });
  questionArea.appendChild(submit);
}

function renderDraft(text) {
  const title = document.createElement("h3");
  title.textContent = "AIが整理した報告書案";
  const box = document.createElement("div");
  box.className = "draft-box";
  box.textContent = text;
  draftResult.append(title, box);
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
      result.textContent = `保存しました。ID: ${body.id}`;
      form.reset();
    } catch (error) {
      result.textContent = `エラー：${error.message}`;
    }
  });
}
