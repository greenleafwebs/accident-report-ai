const form = document.getElementById("reportForm");
const result = document.getElementById("result");
const excelFile = document.getElementById("excelFile");
const readExcel = document.getElementById("readExcel");
const excelResult = document.getElementById("excelResult");
const preview = document.getElementById("preview");
const checkedItems = document.getElementById("checkedItems");
const textItems = document.getElementById("textItems");

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

    const checked = [];
    const texts = [];

    for (const sheetName of workbook.SheetNames) {
      const sheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        defval: ""
      });

      rows.forEach((row, rowIndex) => {
        row.forEach((cell, colIndex) => {
          const value = String(cell ?? "").trim();

          if (!value) return;

          // 選択欄：■の右隣セルを取得
          if (value.includes("■")) {
            const right = String(row[colIndex + 1] ?? "").trim();
            if (right) {
              checked.push({
                sheet: sheetName,
                row: rowIndex + 1,
                column: colIndex + 1,
                value: right
              });
            }
          }

          // 自由記入欄などの値も確認できるように保持
          if (!/[□■]/.test(value) && value.length >= 2) {
            texts.push({
              sheet: sheetName,
              row: rowIndex + 1,
              column: colIndex + 1,
              value
            });
          }
        });
      });
    }

    renderList(checkedItems, checked, "選択なし");
    renderList(textItems, texts, "入力内容なし");

    preview.hidden = false;
    excelResult.textContent = `読み込み完了：${workbook.SheetNames.length}シート`;
  } catch (error) {
    excelResult.textContent = `読み込みエラー：${error.message}`;
  }
});

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
