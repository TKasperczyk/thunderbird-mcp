const aiButton = document.getElementById("aiTranslate");
const qualityButton = document.getElementById("qualityTranslate");
const localButton = document.getElementById("localTranslate");
const originalButton = document.getElementById("showOriginal");
const status = document.getElementById("status");

function setBusy(busy) {
  aiButton.disabled = busy;
  qualityButton.disabled = busy;
  localButton.disabled = busy;
  originalButton.disabled = busy;
}

async function getActiveTab() {
  const [tab] = await browser.tabs.query({ active: true, currentWindow: true });
  if (!tab?.id) throw new Error("Açık ileti sekmesi bulunamadı.");
  return tab;
}

async function translate(provider) {
  setBusy(true);
  status.className = "";
  status.textContent = provider === "openai-fast"
    ? "Hızlı AI ile çevriliyor…"
    : provider === "openai-quality"
      ? "Kaliteli AI ile çevriliyor…"
      : "Cihazda çevriliyor…";
  try {
    const tab = await getActiveTab();
    const result = await browser.mcpServer.translateDisplayedMessageInline(
      tab.id,
      "tr",
      "auto",
      provider
    );
    if (result?.error) throw new Error(result.error);
    status.textContent = result?.alreadyTargetLanguage
      ? "İleti zaten Türkçe görünüyor."
      : result?.translated === false
        ? "Orijinal ileti gösteriliyor."
        : provider.startsWith("openai")
          ? "AI çevirisi gösteriliyor."
          : "Yerel çeviri gösteriliyor.";
  } catch (error) {
    status.className = "error";
    status.textContent = error?.message || String(error);
  } finally {
    setBusy(false);
  }
}

async function restoreOriginal() {
  setBusy(true);
  status.className = "";
  status.textContent = "Orijinal ileti açılıyor…";
  try {
    const tab = await getActiveTab();
    const result = await browser.mcpServer.translateDisplayedMessageInline(
      tab.id,
      "tr",
      "auto",
      "restore"
    );
    if (result?.error) throw new Error(result.error);
    status.textContent = result?.alreadyOriginal
      ? "İleti zaten orijinal halinde."
      : "Orijinal ileti gösteriliyor.";
  } catch (error) {
    status.className = "error";
    status.textContent = error?.message || String(error);
  } finally {
    setBusy(false);
  }
}

aiButton.addEventListener("click", () => translate("openai-fast"));
qualityButton.addEventListener("click", () => translate("openai-quality"));
localButton.addEventListener("click", () => translate("local"));
originalButton.addEventListener("click", restoreOriginal);
