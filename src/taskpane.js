/* Task pane UI: reply generation + settings form. */
(function () {
  "use strict";
  const $ = (id) => document.getElementById(id);
  const FIELDS = ["provider", "baseUrl", "apiKey", "model", "maxTokens", "tone", "customTone", "length",
    "language", "swissSpelling", "replyAll", "aboutMe", "signature", "extraInstructions"];
  let settings, ctx;

  function setStatus(el, msg, cls) { el.textContent = msg || ""; el.className = "status" + (cls ? " " + cls : ""); }

  function fillForm(s) {
    FIELDS.forEach((f) => {
      const el = $("s-" + f);
      if (el.type === "checkbox") el.checked = !!s[f]; else el.value = s[f] == null ? "" : s[f];
    });
    toggleFields();
  }

  function readForm() {
    const s = Object.assign({}, settings);
    FIELDS.forEach((f) => {
      const el = $("s-" + f);
      s[f] = el.type === "checkbox" ? el.checked : el.value.trim();
    });
    s.maxTokens = Number(s.maxTokens) || KI.DEFAULTS.maxTokens;
    return s;
  }

  function toggleFields() {
    const p = $("s-provider").value;
    $("s-baseurl-wrap").hidden = p === "anthropic";
    $("s-baseUrl").placeholder = p === "openai" ? "https://api.openai.com/v1" : "https://mein-proxy.example.com";
    $("s-apiKey").placeholder = p === "openai" ? "sk-… (bei lokalen Modellen leer lassen)" : "sk-ant-…";
    $("s-customTone-wrap").hidden = $("s-tone").value !== "custom";
  }

  function switchTab(name) {
    document.querySelectorAll(".tab").forEach((t) => t.classList.toggle("active", t.dataset.tab === name));
    document.querySelectorAll(".panel").forEach((p) => p.classList.toggle("active", p.id === "tab-" + name));
  }

  async function onGenerate() {
    const btn = $("btn-generate");
    const st = $("status");
    const out = $("result");
    if (!settings.apiKey && settings.provider !== "openai") {
      switchTab("settings");
      setStatus($("s-status"), "Bitte zuerst einen API-Key erfassen und speichern.", "err");
      return;
    }
    btn.disabled = true; $("btn-retry").disabled = true;
    setStatus(st, "Lese E-Mail …");
    try {
      ctx = await KI.getMailContext();
      const s = Object.assign({}, settings);
      if ($("r-tone").value) s.tone = $("r-tone").value;
      const prompt = KI.buildPrompt(s, ctx, $("r-hint").value.trim());
      out.hidden = false; out.value = ""; $("result-actions").hidden = true;
      setStatus(st, "Claude schreibt …".replace("Claude", s.provider === "openai" ? "KI" : "Claude"));
      const text = await KI.generate(s, prompt, (_d, full) => { out.value = full; out.scrollTop = out.scrollHeight; });
      out.value = text;
      $("result-actions").hidden = false;
      $("btn-use").textContent = ctx.isCompose ? "In E-Mail einfügen" : "In Antwort übernehmen";
      setStatus(st, "Fertig – prüfen, bei Bedarf anpassen, dann übernehmen.", "ok");
    } catch (e) {
      setStatus(st, (e && e.message) || String(e), "err");
    } finally {
      btn.disabled = false; $("btn-retry").disabled = false;
    }
  }

  async function onUse() {
    const text = $("result").value.trim();
    if (!text) return;
    try {
      if (ctx && ctx.isCompose) {
        await KI.insertIntoDraft(text);
        setStatus($("status"), "In die E-Mail eingefügt.", "ok");
      } else {
        KI.openReply(text, $("r-replyall").checked);
        setStatus($("status"), "Antwortfenster geöffnet.", "ok");
      }
    } catch (e) { setStatus($("status"), (e && e.message) || String(e), "err"); }
  }

  async function onSave() {
    settings = readForm();
    try { await KI.saveSettings(settings); setStatus($("s-status"), "Gespeichert.", "ok"); $("r-replyall").checked = settings.replyAll; }
    catch (e) { setStatus($("s-status"), "Speichern fehlgeschlagen: " + ((e && e.message) || e), "err"); }
  }

  async function onModels() {
    const s = readForm();
    setStatus($("s-status"), "Lade Modelle …");
    try {
      const models = await KI.listModels(s);
      $("model-list").innerHTML = models.map((m) => `<option value="${m.id}">${m.name}</option>`).join("");
      setStatus($("s-status"), `${models.length} Modelle gefunden. Feld «Modell» anklicken zur Auswahl.\n` +
        models.slice(0, 8).map((m) => "• " + m.id).join("\n"), "ok");
    } catch (e) { setStatus($("s-status"), (e && e.message) || String(e), "err"); }
  }

  async function onTest() {
    const s = readForm();
    s.maxTokens = 20;
    setStatus($("s-status"), "Teste Verbindung …");
    try {
      const t = await KI.generate(s, { system: "Antworte nur mit: OK", user: "Test" });
      setStatus($("s-status"), `Verbindung OK (${s.model}): «${t}»`, "ok");
    } catch (e) { setStatus($("s-status"), (e && e.message) || String(e), "err"); }
  }

  Office.onReady(() => {
    settings = KI.loadSettings();
    fillForm(settings);
    $("r-replyall").checked = settings.replyAll;
    const item = Office.context.mailbox.item;
    const isCompose = item && typeof item.subject !== "string";
    $("r-replyall-wrap").hidden = isCompose;
    if (isCompose) $("btn-generate").textContent = "Antwortvorschlag mit KI erstellen und einfügen";

    document.querySelectorAll(".tab").forEach((t) => t.addEventListener("click", () => switchTab(t.dataset.tab)));
    $("s-provider").addEventListener("change", toggleFields);
    $("s-tone").addEventListener("change", toggleFields);
    $("btn-generate").addEventListener("click", onGenerate);
    $("btn-retry").addEventListener("click", onGenerate);
    $("btn-use").addEventListener("click", onUse);
    $("btn-copy").addEventListener("click", () => { navigator.clipboard && navigator.clipboard.writeText($("result").value); });
    $("btn-save").addEventListener("click", onSave);
    $("btn-models").addEventListener("click", onModels);
    $("btn-test").addEventListener("click", onTest);

    if (!settings.apiKey && settings.provider !== "openai") {
      switchTab("settings");
      setStatus($("s-status"), "Willkommen! Bitte API-Key erfassen und speichern.");
    }
  });
})();
