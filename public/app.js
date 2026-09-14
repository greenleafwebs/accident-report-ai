const form = document.getElementById("reportForm");
const result = document.getElementById("result");
const excelFile = document.getElementById("excelFile");
const readExcel = document.getElementById("readExcel");
const excelResult = document.getElementById("excelResult");
const preview = document.getElementById("preview");
const checkedItems = document.getElementById("checkedItems");
const textItems = document.getElementById("textItems");

// Excel帳票で読み取りたい「選択式の項目名」。
const selectionFields = [
  "第○報",
  "事故状況の程度",
  "性別：",
  "住所",
  "要介護度",
  "発生場所",
  "事故の種別",
  "受診方法",
  "診断内容",
  "続柄",
  "連絡した関係機関 (連絡した場合のみ)"
];

// Excel帳票で読み取りたい「自由記入・入力式の項目名」。
const textFields = [
  "氏名",
  "発生日時",
  "発生時状況、事故内容の詳細",
  "発生時の対応",
  "医療機関名",
  "連絡先（電話番号）",
  "診断名",
  "検査、処置等の概要",
  "利用者の状況",
  "本人、家族、関係先等への追加対応予定",
  "原因分析",
  "再発防止策",
  "その他特記すべき事項"
];

readExcel.addEventListener("click", async () => {
  const file = excelFile.files[0];

  if (!file) {
    excelResult.textContent = "Excelファイルを選択してください。";
    return;
  }

  excelResult.textContent = "読み込み中…";
  checkedItems.innerHTML = "";
  textItems.innerHTML = "";
  preview.hidden = true;

  try {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array" });
    const selected = [];
    const texts = [];

    for (const sheetName of workbook.SheetNames) {
      const sheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        defval: ""
      });

      for (let rowIndex = 0; rowIndex < rows.length; rowIndex++) {
        const row = rows[rowIndex];

        // 選択式：項目名を起点に、その行の■だけを取得
        for (const fieldName of selectionFields) {
          const fieldColumns = findFieldColumns(row, fieldName);

          for (const fieldColumn of fieldColumns) {
            const values = getCheckedValues(row, fieldColumn);

            values.forEach((value) => {
              selected.push({
                field: fieldName,
                value,
                sheet: sheetName,
                row: rowIndex + 1
              });
            });
          }
        }

        // 自由記入：通常は同じ行の右側、原因分析・再発防止策は複数行の入力欄を取得
        for (const fieldName of textFields) {
          const fieldColumns = findFieldColumns(row, fieldName);

          for (const fieldColumn of fieldColumns) {
            const value = isMultiRowTextField(fieldName)
              ? getMultiRowTextValue(rows, rowIndex, fieldColumn, fieldName)
              : getTextValue(row, fieldColumn, fieldName);

            if (value) {
              texts.push({
                field: fieldName,
                value,
                sheet: sheetName,
                row: rowIndex + 1
              });
            }
          }
        }
      }
    }

    renderSelectedFields(checkedItems, selected);
    renderTextFields(textItems, texts);

    preview.hidden = false;
    excelResult.textContent = `読み込み完了：${workbook.SheetNames.length}シート`;
  } catch (error) {
    excelResult.textContent = `読み込みエラー：${error.message}`;
  }
});

function findFieldColumns(row, fieldName) {
  const columns = [];

  row.forEach((cell, index) => {
    const value = String(cell ?? "").trim();
    if (value === fieldName || value.includes(fieldName)) {
      columns.push(index);
    }
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
      if (
        right &&
        !right.includes("□") &&
        !right.includes("■") &&
        !right.includes("☐") &&
        !right.includes("☑")
      ) {
        values.push(right);
      }
    }
  }

  return [...new Set(values)];
}

function getTextValue(row, fieldColumn, fieldName) {
  const values = [];

  for (let index = fieldColumn + 1; index < row.length; index++) {
    const value = String(row[index] ?? "").trim();
    if (!value) continue;
    if (isCheckbox(value)) continue;
    if (isPlaceholderText(value)) continue;

    // 別の主要項目名に到達したら、そこで終了する。
    if (isKnownFieldLabel(value, fieldName)) break;

    values.push(value);
  }

  return cleanTextValues(values, fieldName);
}

// 「原因分析」「再発防止策」は、項目名の次の行に入力欄があるため、
// 項目名の行だけでなく、その後の数行も確認する。
function isMultiRowTextField(fieldName) {
  return fieldName === "原因分析" || fieldName === "再発防止策";
}

function getMultiRowTextValue(rows, startRow, fieldColumn, fieldName) {
  const values = [];
  const maxRows = Math.min(rows.length, startRow + 3);

  for (let rowIndex = startRow; rowIndex < maxRows; rowIndex++) {
    const row = rows[rowIndex];

    for (let index = fieldColumn + 1; index < row.length; index++) {
      const value = String(row[index] ?? "").trim();
      if (!value) continue;
      if (isCheckbox(value)) continue;
      if (isPlaceholderText(value)) continue;

      // 次の主要項目が始まったら、この入力欄は終了。
      if (rowIndex > startRow && isKnownFieldLabel(value, fieldName)) {
        return cleanTextValues(values, fieldName);
      }

      values.push(value);
    }
  }

  return cleanTextValues(values, fieldName);
}

function isPlaceholderText(value) {
  return value.includes("できるだけ具体的に記載すること");
}

function cleanTextValues(values, fieldName) {
  const cleaned = values
    .map((value) => value.replace(/\s+/g, " ").trim())
    .filter(Boolean)
    .filter((value) => value !== fieldName)
    .filter((value) => !isPlaceholderText(value));

  return [...new Set(cleaned)].join(" ").trim();
}

function isCheckbox(value) {
  return /^[□■☐☑]+$/.test(value);
}

function isKnownFieldLabel(value, currentFieldName) {
  const allFields = [...selectionFields, ...textFields];
  return allFields.some((fieldName) =>
    fieldName !== currentFieldName && value === fieldName
  );
}

function normalizeFieldName(fieldName) {
  return fieldName.replace(/[：:]+$/g, "").trim();
}

function renderSelectedFields(container, items) {
  if (items.length === 0) {
    container.textContent = "選択された項目は見つかりませんでした。";
    return;
  }

  const grouped = new Map();

  items.forEach((item) => {
    if (!grouped.has(item.field)) {
      grouped.set(item.field, []);
    }
    grouped.get(item.field).push(item);
  });

  const list = document.createElement("ul");

  grouped.forEach((fieldItems, fieldName) => {
    const li = document.createElement("li");
    const values = [...new Set(fieldItems.map((item) => item.value))];
    li.textContent = `${normalizeFieldName(fieldName)}：${values.join("、")}`;
    list.appendChild(li);
  });

  container.appendChild(list);
}

function renderTextFields(container, items) {
  if (items.length === 0) {
    container.textContent = "自由記入欄は見つかりませんでした。";
    return;
  }

  const grouped = new Map();

  items.forEach((item) => {
    if (!grouped.has(item.field)) {
      grouped.set(item.field, []);
    }
    grouped.get(item.field).push(item.value);
  });

  const list = document.createElement("ul");

  grouped.forEach((values, fieldName) => {
    const li = document.createElement("li");
    li.textContent = `${normalizeFieldName(fieldName)}：${[...new Set(values)].join(" / ")}`;
    list.appendChild(li);
  });

  container.appendChild(list);
}

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  result.textContent = "保存中…";

  const data = Object.fromEntries(new FormData(form).entries());

  try {
    const response = await fetch("/api/reports", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(data)
    });

    const body = await response.json();

    if (!response.ok || !body.success) {
      throw new Error(body.error || "保存に失敗しました");
    }

    result.textContent = `保存しました。ID: ${body.id}`;
    form.reset();
  } catch (error) {
    result.textContent = `エラー：${error.message}`;
  }
});