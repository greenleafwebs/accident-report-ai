const form = document.getElementById("reportForm");
const result = document.getElementById("result");
const excelFile = document.getElementById("excelFile");
const readExcel = document.getElementById("readExcel");
const excelResult = document.getElementById("excelResult");
const preview = document.getElementById("preview");
const checkedItems = document.getElementById("checkedItems");
const textItems = document.getElementById("textItems");

// Excel帳票で読み取りたい「項目名」。
// 項目名を起点に同じ行の■だけを調べるため、未選択の☐は拾わない。
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

        // 項目名を起点に、その行の選択結果を取得
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

        // 自由記入欄などの値は今まで通り確認用に保持
        row.forEach((cell, colIndex) => {
          const value = String(cell ?? "").trim();
          if (!value) return;

          if (!/[□■☐☑]/.test(value) && value.length >= 2) {
            texts.push({
              sheet: sheetName,
              row: rowIndex + 1,
              column: colIndex + 1,
              value
            });
          }
        });
      }
    }

    renderSelectedFields(checkedItems, selected);
    renderList(textItems, texts, "入力内容なし");

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

  // 項目名より右側だけを調べる。
  // 次の項目名らしきセルまでを対象にすることで、別項目の■を混ぜにくくする。
  for (let index = fieldColumn + 1; index < row.length; index++) {
    const value = String(row[index] ?? "").trim();
    if (!value) continue;

    // ■があれば「その右隣」を選択値として取得
    if (value.includes("■")) {
      const right = String(row[index + 1] ?? "").trim();
      if (right && !right.includes("□") && !right.includes("■") && !right.includes("☐") && !right.includes("☑")) {
        values.push(right);
      }
    }
  }

  return [...new Set(values)];
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
    li.textContent = `${fieldName}：${values.join("、")}`;
    list.appendChild(li);
  });

  container.appendChild(list);
}

function renderList(container, items, emptyText) {
  if (items.length === 0) {
    container.textContent = emptyText;
    return;
  }

  const list = document.createElement("ul");

  items.forEach((item) => {
    const li = document.createElement("li");
    li.textContent = `${item.value} （${item.sheet} / ${item.row}行 ${item.column}列）`;
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
