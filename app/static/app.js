"use strict";
const $ = (id) => document.getElementById(id);
let files = [];
let busy = false;
let ready = false;
const codes = new Map();

function message(text) {
  $("load-error").textContent = text;
  $("load-error").hidden = !text;
}

function renderFiles() {
  $("file-list").replaceChildren();
  files.forEach((file, index) => {
    const row = document.createElement("div");
    row.className = "file-row";
    const top = document.createElement("div");
    top.className = "file-top";
    const name = document.createElement("span");
    name.className = "filename";
    name.textContent = file.name;
    const remove = document.createElement("button");
    remove.type = "button";
    remove.className = "remove";
    remove.textContent = "×";
    remove.disabled = busy;
    remove.setAttribute("aria-label", `Quitar ${file.name}`);
    remove.addEventListener("click", () => { codes.delete(file); files.splice(index, 1); renderFiles(); });
    top.append(name, remove);
    const label = document.createElement("label");
    label.className = "code-label";
    label.textContent = "Código correcto · déjalo vacío para conservar el actual";
    const input = document.createElement("input");
    input.className = "code-input";
    input.id = `code-${index}`;
    input.type = "text";
    input.maxLength = 13;
    input.pattern = "[Qq][Cc][Pp]-[0-9]{4}-[0-9]{4}";
    input.placeholder = "QCP-2026-0001";
    input.value = codes.get(file) || "";
    input.disabled = busy;
    input.addEventListener("input", () => codes.set(file, input.value));
    label.htmlFor = input.id;
    row.append(top, label, input);
    $("file-list").append(row);
  });
  $("submit").disabled = busy || !ready || files.length === 0;
}

function addFiles(incoming) {
  if (busy || !ready) return;
  const errors = [];
  for (const file of incoming) {
    if (!/\.(xlsx|pdf)$/i.test(file.name)) { errors.push(`${file.name}: selecciona un Excel .xlsx o un PDF.`); continue; }
    if (file.size > 15 * 1024 * 1024) { errors.push(`${file.name}: supera los 15 MB.`); continue; }
    if (files.length >= 20) { errors.push("El lote admite hasta 20 archivos."); break; }
    if (files.some((other) => other.name === file.name && other.size === file.size && other.lastModified === file.lastModified)) continue;
    files.push(file);
  }
  message(errors.join(" "));
  renderFiles();
}

$("document-files").addEventListener("change", (event) => { addFiles(event.target.files); event.target.value = ""; });
const dropzone = $("dropzone");
for (const eventName of ["dragenter", "dragover"]) {
  dropzone.addEventListener(eventName, (event) => { event.preventDefault(); if (!busy) dropzone.classList.add("dragover"); });
}
for (const eventName of ["dragleave", "drop"]) {
  dropzone.addEventListener(eventName, (event) => { event.preventDefault(); dropzone.classList.remove("dragover"); });
}
dropzone.addEventListener("drop", (event) => addFiles(event.dataTransfer.files));

function addResult(file, result, error) {
  const card = document.createElement("article");
  card.className = `result-card${error ? " failed" : ""}`;
  const label = document.createElement("div");
  label.className = `result-label${error ? " failed" : ""}`;
  label.textContent = error ? "NO SE PUDO PROCESAR" : "CERTIFICADO PREPARADO";
  const title = document.createElement("strong");
  title.textContent = file.name;
  card.append(label, title);
  if (error) {
    const detail = document.createElement("p");
    detail.textContent = error;
    card.append(detail);
  } else {
    const detail = document.createElement("p");
    detail.textContent = `${result.finalCode || "Sin código único detectado"} · ${result.signatureIncluded ? "Con imagen de firma" : "Sin firma"}`;
    card.append(detail);
    for (const warning of result.warnings) {
      const text = document.createElement("p");
      text.className = "warning";
      text.textContent = warning;
      card.append(text);
    }
    const actions = document.createElement("div");
    actions.className = "actions";
    const download = document.createElement("a");
    download.href = result.downloadUrl;
    download.textContent = "Descargar PDF ↓";
    const verify = document.createElement("a");
    // Relative URL keeps local development usable if localhost and 127.0.0.1 differ.
    verify.href = `/certificados/${result.id}`;
    verify.target = "_blank";
    verify.rel = "noopener";
    verify.textContent = "Ver consulta ↗";
    actions.append(download, verify);
    card.append(actions);
  }
  $("result-list").append(card);
}

$("certificate-form").addEventListener("submit", async (event) => {
  event.preventDefault();
  if (busy || !ready || !files.length) return;
  const signature = $("signature").files[0];
  if (signature && (signature.size > 2 * 1024 * 1024 || !/\.(png|jpe?g)$/i.test(signature.name))) {
    return message("La firma debe ser PNG o JPG de hasta 2 MB.");
  }
  message("");
  busy = true;
  renderFiles();
  $("document-files").disabled = true;
  $("signature").disabled = true;
  $("submit").textContent = "Preparando…";
  $("result-list").replaceChildren();
  $("results").hidden = false;
  const failed = [];
  const batch = [...files];
  let completed = 0;
  try {
    for (const [index, file] of batch.entries()) {
      $("progress").textContent = `Procesando ${index + 1} de ${batch.length}: ${file.name}`;
      const form = new FormData();
      form.append(/\.xlsx$/i.test(file.name) ? "excel" : "pdf", file);
      form.append("codigoCorrecto", (codes.get(file) || "").trim());
      if (signature) form.append("firma", signature);
      try {
        const response = await fetch("/api/certificates", { method: "POST", body: form });
        const result = await response.json();
        if (!response.ok) throw new Error(typeof result.detail === "string" ? result.detail : "La carga no es válida.");
        addResult(file, result, null);
        codes.delete(file);
        completed++;
      } catch (error) {
        failed.push(file);
        addResult(file, null, error instanceof TypeError ? "No se pudo conectar con el portal. Revisa tu conexión e inténtalo de nuevo." : error.message);
      }
    }
  } finally {
    files = failed;
    busy = false;
    $("document-files").disabled = false;
    $("signature").disabled = false;
    $("submit").textContent = failed.length ? "Reintentar archivos pendientes →" : "Preparar certificados →";
    $("progress").textContent = `${completed} de ${batch.length} certificados preparados.${failed.length ? " Los archivos pendientes siguen seleccionados." : " Ya puedes descargarlos."}`;
    renderFiles();
  }
});

async function initialize() {
  try {
    const response = await fetch("/api/config");
    if (!response.ok) throw new Error("El portal no está disponible.");
    const config = await response.json();
    $("local-notice").hidden = !config.localMode;
    ready = config.uploadsEnabled === true;
    $("setup-notice").hidden = ready;
    $("document-files").disabled = !ready;
    $("signature").disabled = !ready;
    renderFiles();
  } catch {
    message("No se pudo iniciar el portal. Recarga la página para volver a intentarlo.");
  }
}
initialize();
