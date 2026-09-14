const form = document.getElementById("reportForm");
const result = document.getElementById("result");

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
