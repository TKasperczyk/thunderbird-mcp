const aiButton = document.getElementById("aiTranslate");
const localButton = document.getElementById("localTranslate");
const originalButton = document.getElementById("showOriginal");
const status = document.getElementById("status");

function setBusy(busy) {
  aiButton.disabled = busy;
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
  status.textContent = provider === "openai"
    ? "gpt-4o-mini ile çevriliyor…"
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
        : provider === "openai"
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

aiButton.addEventListener("click", () => translate("openai"));
localButton.addEventListener("click", () => translate("local"));
originalButton.addEventListener("click", restoreOriginal);
