const MODULES = {
  bmf: "Brand Messaging Framework",
  pain_points: "Pain Point Questions",
  faqs: "Home Page FAQs",
  info: "AI Info Page",
};

const brandInput = document.querySelector("#brand");
const nicheInput = document.querySelector("#niche");
const urlInput = document.querySelector("#url");
const runAllButton = document.querySelector("#run-all");
const formError = document.querySelector("#form-error");
const processStatus = document.querySelector("#process-status");
const outputStatus = document.querySelector("#output-status");
const outputList = document.querySelector("#output-list");
const progressFill = document.querySelector("#progress-fill");
const progressLabel = document.querySelector("#progress-label");
const progressValue = document.querySelector("#progress-value");
const statusNote = document.querySelector("#status-note");
const engineDot = document.querySelector("#engine-dot");

const state = { results: new Map(), running: false };

function showError(message) {
  formError.textContent = message;
  formError.classList.toggle("hidden", !message);
}

function percent(value) {
  const safe = Math.max(0, Math.min(100, Math.round(value)));
  progressFill.style.width = `${safe}%`;
  progressValue.textContent = `${safe}%`;
}

function payloadFor(module) {
  return { module, brand: brandInput.value.trim(), niche: nicheInput.value.trim(), url: urlInput.value.trim() };
}

function validate(module = "") {
  if (!brandInput.value.trim()) return "Add a brand name first.";
  if (!nicheInput.value.trim()) return "Add a target niche first.";
  if (module === "info" && !urlInput.value.trim()) return "AI Info Page needs an official website URL.";
  return "";
}

function download(content, name, type) {
  const blob = new Blob([content], { type });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = name;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(link.href);
}

function slug(value) {
  return value.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/(^-|-$)/g, "") || "brand-audit";
}

function renderOutputs() {
  if (!state.results.size) {
    outputList.innerHTML = '<div class="output-placeholder"><span class="output-orbit">✦</span><strong>Reports will appear here</strong><p>Run one module or the full audit to unlock downloads.</p></div>';
    outputStatus.textContent = "PENDING";
    return;
  }
  outputStatus.textContent = `${state.results.size} READY`;
  outputList.innerHTML = [...state.results.entries()].map(([module, result]) => `
    <div class="output-item">
      <strong>${result.label}</strong>
      <div class="output-actions">
        <button class="primary-button download-button" data-download="md" data-module="${module}" type="button">↓ Markdown</button>
        <button class="primary-button download-button secondary" data-download="json" data-module="${module}" type="button">↓ JSON</button>
      </div>
    </div>
  `).join("");
}

function setBusy(busy) {
  state.running = busy;
  runAllButton.disabled = busy;
  document.querySelectorAll(".process-item").forEach((button) => { button.disabled = busy; });
}

async function runModule(module) {
  const validation = validate(module);
  if (validation) { showError(validation); return null; }
  showError("");
  const item = document.querySelector(`[data-module="${module}"]`);
  item?.classList.add("running");
  item?.querySelector(".process-state") && (item.querySelector(".process-state").textContent = "RUNNING");
  progressLabel.textContent = `RUNNING ${MODULES[module].toUpperCase()}`;
  try {
    const response = await fetch("/api/run", { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify(payloadFor(module)) });
    const data = await response.json();
    if (!response.ok || !data.ok) throw new Error(data.error || "The Worker could not complete this module.");
    state.results.set(module, data);
    item?.classList.remove("running");
    item?.classList.add("done");
    item?.querySelector(".process-state") && (item.querySelector(".process-state").textContent = "DONE");
    renderOutputs();
    return data;
  } catch (error) {
    item?.classList.remove("running");
    item?.querySelector(".process-state") && (item.querySelector(".process-state").textContent = "ERROR");
    throw error;
  }
}

async function runAll() {
  if (state.running) return;
  const validation = validate();
  if (validation) { showError(validation); return; }
  setBusy(true);
  state.results.clear();
  renderOutputs();
  processStatus.textContent = "RUNNING";
  const modules = Object.keys(MODULES);
  try {
    for (let index = 0; index < modules.length; index += 1) {
      percent((index / modules.length) * 100);
      try { await runModule(modules[index]); }
      catch (error) { showError(error.message); }
      percent(((index + 1) / modules.length) * 100);
    }
    progressLabel.textContent = `${state.results.size} OF ${modules.length} MODULES COMPLETE`;
    processStatus.textContent = "COMPLETE";
  } finally {
    setBusy(false);
  }
}

runAllButton.addEventListener("click", runAll);
document.querySelectorAll(".process-item").forEach((button) => {
  button.addEventListener("click", async () => {
    if (state.running) return;
    const module = button.dataset.module;
    setBusy(true);
    processStatus.textContent = "RUNNING";
    try {
      await runModule(module);
      progressLabel.textContent = `${MODULES[module].toUpperCase()} COMPLETE`;
      percent(100);
      processStatus.textContent = "COMPLETE";
    } catch (error) {
      showError(error.message);
      progressLabel.textContent = "RUN FAILED";
      processStatus.textContent = "ERROR";
    } finally { setBusy(false); }
  });
});

outputList.addEventListener("click", (event) => {
  const button = event.target.closest("[data-download]");
  if (!button) return;
  const result = state.results.get(button.dataset.module);
  if (!result) return;
  const base = `${slug(brandInput.value)}-${button.dataset.module}`;
  if (button.dataset.download === "json") download(JSON.stringify(result, null, 2), `${base}.json`, "application/json");
  else download(`# ${result.label}\n\n${result.content.trim()}\n`, `${base}.md`, "text/markdown");
});

async function checkHealth() {
  try {
    const response = await fetch("/api/health");
    const data = await response.json();
    if (!data.configured?.openai) throw new Error("OPENAI_API_KEY is not configured");
    statusNote.textContent = "Worker connected · key configured";
  } catch (error) {
    engineDot.classList.add("offline");
    statusNote.textContent = `Worker setup needed · ${error.message}`;
  }
}

renderOutputs();
checkHealth();
