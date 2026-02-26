const fileInput = document.getElementById("ppt-file");
const form = document.getElementById("upload-form");
const statusEl = document.getElementById("status");
const convertBtn = document.getElementById("convert-btn");
const fileLabelText = document.getElementById("file-label-text");

function setStatus(message, type = "info") {
  statusEl.textContent = message || "";
  statusEl.classList.remove("status--info", "status--success", "status--error");
  if (type === "success") statusEl.classList.add("status--success");
  else if (type === "error") statusEl.classList.add("status--error");
  else statusEl.classList.add("status--info");
}

fileInput.addEventListener("change", () => {
  const file = fileInput.files[0];
  fileLabelText.textContent = file ? file.name : "Choose PPTX file…";
  setStatus("");
});

form.addEventListener("submit", async (e) => {
  e.preventDefault();

  const file = fileInput.files[0];
  if (!file) {
    setStatus("Please select a PPTX file first.", "error");
    return;
  }

  if (!/\.(pptx|ppt)$/i.test(file.name)) {
    setStatus("Only .pptx / .ppt files are supported.", "error");
    return;
  }

  convertBtn.disabled = true;
  setStatus("Uploading and converting…", "info");

  try {
    const formData = new FormData();
    formData.append("file", file);

    const resp = await fetch("/api/convert", {
      method: "POST",
      body: formData,
    });

    if (!resp.ok) {
      const data = await resp.json().catch(() => ({}));
      const msg =
        data && data.detail
          ? data.detail
          : `Conversion failed (HTTP ${resp.status}).`;
      throw new Error(msg);
    }

    const blob = await resp.blob();
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;

    const dispo = resp.headers.get("Content-Disposition") || "";
    const match = dispo.match(/filename="(.+?)"/);
    const filename =
      (match && match[1]) || file.name.replace(/\.(pptx|ppt)$/i, "") + "_test_sheet.xlsx";

    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    window.URL.revokeObjectURL(url);

    setStatus("Conversion complete. Excel file downloaded.", "success");
  } catch (err) {
    setStatus(err.message || "Conversion failed. Please try again.", "error");
  } finally {
    convertBtn.disabled = false;
  }
});

