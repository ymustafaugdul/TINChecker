from __future__ import annotations

import json
import os
import socket
import threading
import time
import webbrowser

from flask import Flask, jsonify, render_template_string, request
from werkzeug.serving import make_server

from tin_checker_core import load_countries, validate_bulk_entries, validate_many, validate_tin


I18N = {
    "tr": {
        "pageTitle": "Yabancı VKN Kontrolcüsü",
        "eyebrow": "Vergi Kimlik Doğrulama Motoru",
        "appTitle": "Yabancı VKN Kontrolcüsü",
        "singleTab": "Tekil Kontrol",
        "bulkTab": "Toplu Kontrol",
        "singleTitle": "Tekil Kontrol",
        "singleText": "Bir ülke seçin ve tek TIN/VKN değerini doğrulayın.",
        "country": "Ülke",
        "tin": "TIN / VKN",
        "validate": "Kontrol Et",
        "resetSingle": "Tekil Sonucu Sıfırla",
        "singleResult": "Sonuç",
        "singlePlaceholder": "Henüz sorgu yapılmadı.",
        "bulkTitle": "Toplu Girdi",
        "bulkText": "Birden fazla ülke bloğu ekleyin. Her ülke için TIN/VKN değerlerini satır satır yazın.",
        "addCountry": "Ülke Bloğu Ekle",
        "runBulk": "Toplu Kontrol",
        "resetBulk": "Toplu Sonucu Sıfırla",
        "downloadCsv": "CSV İndir",
        "remove": "Kaldır",
        "tinList": "TIN / VKN listesi",
        "tinListPlaceholder": "Her satıra bir TIN/VKN yazın",
        "summary": "Özet",
        "summaryPlaceholder": "Toplu kontrol çalıştığında özet burada görünür.",
        "details": "Detaylı Sonuçlar",
        "detailsPlaceholder": "Henüz toplu sonuç üretilmedi.",
        "total": "Toplam",
        "valid": "Doğru",
        "invalid": "Yanlış",
        "status": "Durum",
        "description": "Açıklama",
        "validStatus": "Doğru",
        "invalidStatus": "Yanlış",
        "noRows": "Doğrulanacak kayıt bulunamadı.",
        "csvNoData": "İndirilecek toplu sonuç yok.",
        "genericError": "Bir hata oluştu.",
    },
    "en": {
        "pageTitle": "TIN Checker",
        "eyebrow": "Tax ID Validation Engine",
        "appTitle": "TIN Checker",
        "singleTab": "Single Validation",
        "bulkTab": "Bulk Validation",
        "singleTitle": "Single Validation",
        "singleText": "Select a country and validate one TIN/VKN value.",
        "country": "Country",
        "tin": "TIN / VKN",
        "validate": "Validate",
        "resetSingle": "Reset Single Result",
        "singleResult": "Result",
        "singlePlaceholder": "No validation has been run yet.",
        "bulkTitle": "Bulk Input",
        "bulkText": "Add multiple country blocks. Enter one TIN/VKN per line for each country.",
        "addCountry": "Add Country Block",
        "runBulk": "Run Bulk Validation",
        "resetBulk": "Reset Bulk Result",
        "downloadCsv": "Download CSV",
        "remove": "Remove",
        "tinList": "TIN / VKN list",
        "tinListPlaceholder": "Enter one TIN/VKN per line",
        "summary": "Summary",
        "summaryPlaceholder": "Run a bulk validation to show the summary here.",
        "details": "Detailed Results",
        "detailsPlaceholder": "No bulk results have been generated yet.",
        "total": "Total",
        "valid": "Valid",
        "invalid": "Invalid",
        "status": "Status",
        "description": "Description",
        "validStatus": "Valid",
        "invalidStatus": "Invalid",
        "noRows": "No records to validate.",
        "csvNoData": "There are no bulk results to download.",
        "genericError": "An error occurred.",
    },
}


HTML = r"""
<!doctype html>
<html lang="tr">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>TIN Checker</title>
  <style>
    :root {
      --bg: #07111f;
      --panel: #0c182b;
      --text: #edf5ff;
      --muted: #9eb0ca;
      --line: rgba(126, 159, 217, 0.24);
      --accent: #55d6ff;
      --accent-2: #5de4bd;
      --danger: #ff6f91;
      --success: #59f2ae;
      --shadow: 0 24px 70px rgba(0, 0, 0, 0.38);
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      min-height: 100vh;
      color: var(--text);
      font-family: "Segoe UI", Tahoma, sans-serif;
      background:
        radial-gradient(circle at 18% 12%, rgba(85, 214, 255, 0.18), transparent 28%),
        radial-gradient(circle at 86% 8%, rgba(93, 228, 189, 0.14), transparent 22%),
        linear-gradient(180deg, #102746 0%, var(--bg) 100%);
    }
    body::before {
      content: "";
      position: fixed;
      inset: 0;
      pointer-events: none;
      background-image:
        linear-gradient(rgba(255,255,255,0.045) 1px, transparent 1px),
        linear-gradient(90deg, rgba(255,255,255,0.045) 1px, transparent 1px);
      background-size: 32px 32px;
      mask-image: linear-gradient(180deg, rgba(0,0,0,0.82), rgba(0,0,0,0.14));
    }
    .wrap {
      width: min(1220px, calc(100% - 32px));
      margin: 18px auto 40px;
      position: relative;
      z-index: 1;
    }
    .shell {
      background: linear-gradient(180deg, rgba(12, 24, 43, 0.96), rgba(5, 12, 24, 0.97));
      border: 1px solid var(--line);
      border-radius: 22px;
      box-shadow: var(--shadow);
      padding: 22px;
    }
    .topbar {
      display: flex;
      justify-content: space-between;
      align-items: flex-start;
      gap: 18px;
      margin-bottom: 20px;
    }
    .eyebrow {
      display: inline-flex;
      padding: 8px 13px;
      border: 1px solid rgba(85, 214, 255, 0.28);
      border-radius: 999px;
      color: var(--accent-2);
      background: rgba(85, 214, 255, 0.08);
      font-size: 12px;
      font-weight: 700;
      letter-spacing: 0.11em;
      text-transform: uppercase;
    }
    h1 {
      margin: 14px 0 8px;
      font-size: clamp(2.4rem, 6vw, 4.2rem);
      line-height: 0.95;
    }
    .subtitle {
      margin: 0;
      max-width: 800px;
      color: var(--muted);
      line-height: 1.6;
      font-size: 1.03rem;
    }
    .actions, .tabs, .button-row, .bulk-actions {
      display: flex;
      gap: 10px;
      flex-wrap: wrap;
    }
    .actions { justify-content: flex-end; }
    button, select, input, textarea {
      font: inherit;
      border-radius: 12px;
    }
    button {
      min-height: 44px;
      cursor: pointer;
      border: 1px solid rgba(126, 159, 217, 0.26);
      color: var(--text);
      background: rgba(255,255,255,0.035);
      padding: 10px 14px;
      font-weight: 700;
      transition: transform 0.16s ease, border-color 0.16s ease, background 0.16s ease;
    }
    button:hover {
      transform: translateY(-1px);
      border-color: rgba(85, 214, 255, 0.48);
      background: rgba(85, 214, 255, 0.08);
    }
    button.primary {
      color: #03141e;
      border: none;
      background: linear-gradient(135deg, var(--accent-2), var(--accent));
      box-shadow: 0 10px 22px rgba(85, 214, 255, 0.22);
    }
    button.active {
      border-color: rgba(85, 214, 255, 0.48);
      background: linear-gradient(135deg, rgba(85, 214, 255, 0.15), rgba(93, 228, 189, 0.1));
    }
    .lang-btn { min-width: 54px; }
    .tabs { margin: 22px 0 16px; }
    .panel { display: none; }
    .panel.active { display: grid; gap: 16px; }
    .section {
      background: linear-gradient(180deg, rgba(16, 31, 54, 0.88), rgba(8, 17, 32, 0.94));
      border: 1px solid var(--line);
      border-radius: 16px;
      padding: 17px;
    }
    .section h2 {
      margin: 0 0 6px;
      font-size: 1.08rem;
    }
    .section p {
      margin: 0 0 14px;
      color: var(--muted);
      line-height: 1.5;
    }
    .single-grid {
      display: grid;
      grid-template-columns: 1.1fr 1.8fr auto;
      gap: 12px;
      align-items: end;
    }
    .bulk-group {
      display: grid;
      gap: 12px;
      padding: 14px;
      border: 1px solid var(--line);
      border-radius: 14px;
      background: rgba(4, 12, 26, 0.58);
      margin-top: 12px;
    }
    .bulk-head {
      display: grid;
      grid-template-columns: 1fr auto;
      gap: 12px;
      align-items: end;
    }
    label {
      display: grid;
      gap: 8px;
      color: var(--muted);
      font-size: 0.95rem;
    }
    select, input, textarea {
      width: 100%;
      color: var(--text);
      background: rgba(2, 8, 18, 0.9);
      border: 1px solid rgba(126, 159, 217, 0.26);
      outline: none;
      padding: 13px 14px;
    }
    textarea {
      min-height: 150px;
      resize: vertical;
      font-family: Consolas, "Courier New", monospace;
    }
    select:focus, input:focus, textarea:focus {
      border-color: rgba(85, 214, 255, 0.52);
      box-shadow: 0 0 0 3px rgba(85, 214, 255, 0.13);
    }
    .result-card, .summary-card, .country-result {
      background: rgba(4, 12, 26, 0.82);
      border: 1px solid var(--line);
      border-radius: 14px;
      padding: 15px;
    }
    .result-head, .summary-card {
      display: flex;
      justify-content: space-between;
      align-items: center;
      gap: 12px;
      flex-wrap: wrap;
    }
    .country-result { margin-bottom: 14px; }
    .country-result h3 {
      margin: 0 0 6px;
      font-size: 1rem;
    }
    .badge {
      display: inline-flex;
      padding: 6px 10px;
      border-radius: 999px;
      font-size: 12px;
      font-weight: 800;
    }
    .badge.ok {
      color: var(--success);
      background: rgba(89, 242, 174, 0.12);
      border: 1px solid rgba(89, 242, 174, 0.24);
    }
    .badge.bad {
      color: var(--danger);
      background: rgba(255, 111, 145, 0.12);
      border: 1px solid rgba(255, 111, 145, 0.24);
    }
    table {
      width: 100%;
      border-collapse: collapse;
      margin-top: 12px;
    }
    th, td {
      text-align: left;
      padding: 10px 8px;
      border-bottom: 1px solid rgba(255,255,255,0.08);
      vertical-align: top;
      font-size: 0.95rem;
    }
    th { color: var(--muted); font-weight: 700; }
    .mono {
      font-family: Consolas, "Courier New", monospace;
      word-break: break-all;
    }
    .muted { color: var(--muted); }
    .error {
      min-height: 24px;
      color: var(--danger);
      font-weight: 700;
      margin-top: 12px;
    }
    .footer {
      margin-top: 22px;
      padding-top: 16px;
      border-top: 1px solid rgba(126, 159, 217, 0.16);
      color: var(--muted);
      text-align: center;
      line-height: 1.7;
      font-size: 0.92rem;
    }
    .footer-author {
      color: var(--text);
      font-weight: 700;
      letter-spacing: 0;
    }
    @media (max-width: 850px) {
      .topbar { flex-direction: column; }
      .actions { justify-content: flex-start; }
      .single-grid, .bulk-head { grid-template-columns: 1fr; }
      button.primary { width: 100%; }
    }
  </style>
</head>
<body>
  <div class="wrap">
    <main class="shell">
      <div class="topbar">
        <div>
          <div class="eyebrow" data-i18n="eyebrow"></div>
          <h1 data-i18n="appTitle"></h1>
        </div>
        <div class="actions">
          <button type="button" class="lang-btn active" data-lang="tr">TR</button>
          <button type="button" class="lang-btn" data-lang="en">EN</button>
        </div>
      </div>

      <div class="tabs">
        <button type="button" class="tab active" data-panel="single" data-i18n="singleTab"></button>
        <button type="button" class="tab" data-panel="bulk" data-i18n="bulkTab"></button>
      </div>

      <section id="single" class="panel active">
        <article class="section">
          <h2 data-i18n="singleTitle"></h2>
          <p data-i18n="singleText"></p>
          <form id="single-form" class="single-grid">
            <label>
              <span data-i18n="country"></span>
              <select id="single-country"></select>
            </label>
            <label>
              <span data-i18n="tin"></span>
              <input id="single-tin" type="text" required>
            </label>
            <div class="button-row">
              <button type="submit" class="primary" data-i18n="validate"></button>
              <button type="button" id="reset-single" data-i18n="resetSingle"></button>
            </div>
          </form>
        </article>
        <article class="section">
          <h2 data-i18n="singleResult"></h2>
          <div id="single-output" class="muted"></div>
        </article>
      </section>

      <section id="bulk" class="panel">
        <article class="section">
          <h2 data-i18n="bulkTitle"></h2>
          <p data-i18n="bulkText"></p>
          <div class="bulk-actions">
            <button type="button" id="add-country" data-i18n="addCountry"></button>
            <button type="button" id="run-bulk" class="primary" data-i18n="runBulk"></button>
            <button type="button" id="reset-bulk" data-i18n="resetBulk"></button>
            <button type="button" id="download-csv" data-i18n="downloadCsv"></button>
          </div>
          <div id="bulk-groups"></div>
        </article>
        <article class="section">
          <h2 data-i18n="summary"></h2>
          <div id="bulk-summary" class="muted"></div>
        </article>
        <article class="section">
          <h2 data-i18n="details"></h2>
          <div id="bulk-output" class="muted"></div>
        </article>
      </section>

      <div id="error" class="error"></div>
      <footer class="footer">
        <div>Made by</div>
        <div class="footer-author">Mustafa ÜĞDÜL</div>
        <div>2026</div>
      </footer>
    </main>
  </div>

  <script>
    const countriesByLang = {{ countries_by_lang|safe }};
    const i18n = {{ i18n|safe }};
    const state = { lang: "tr", lastBulkPayload: null };

    const $ = (id) => document.getElementById(id);

    function t(key) {
      return (i18n[state.lang] && i18n[state.lang][key]) || key;
    }

    function countries() {
      return countriesByLang[state.lang] || countriesByLang.en || [];
    }

    function escapeHtml(value) {
      return String(value ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#39;");
    }

    function optionsHtml(selectedCode = "") {
      return countries().map((item) => {
        const selected = item.code === selectedCode ? " selected" : "";
        return `<option value="${escapeHtml(item.code)}"${selected}>${escapeHtml(item.code)} - ${escapeHtml(item.label)}</option>`;
      }).join("");
    }

    function localizeCountrySelect(select) {
      const current = select.value || countries()[0]?.code || "";
      select.innerHTML = optionsHtml(current);
      if (!select.value && countries()[0]) select.value = countries()[0].code;
    }

    function refreshCountrySelects() {
      localizeCountrySelect($("single-country"));
      document.querySelectorAll(".bulk-country").forEach(localizeCountrySelect);
    }

    function setPlaceholder(id, key) {
      const node = $(id);
      node.textContent = t(key);
      node.className = "muted";
    }

    function resetSingle() {
      $("single-tin").value = "";
      setPlaceholder("single-output", "singlePlaceholder");
      $("error").textContent = "";
    }

    function resetBulk() {
      $("bulk-groups").innerHTML = "";
      state.lastBulkPayload = null;
      addBulkGroup();
      setPlaceholder("bulk-summary", "summaryPlaceholder");
      setPlaceholder("bulk-output", "detailsPlaceholder");
      $("error").textContent = "";
    }

    function updateTexts() {
      document.documentElement.lang = state.lang;
      document.title = t("pageTitle");
      document.querySelectorAll("[data-i18n]").forEach((node) => {
        node.textContent = t(node.dataset.i18n);
      });
      document.querySelectorAll(".lang-btn").forEach((button) => {
        button.classList.toggle("active", button.dataset.lang === state.lang);
      });
      document.querySelectorAll(".bulk-values").forEach((node) => {
        node.placeholder = t("tinListPlaceholder");
      });
      document.querySelectorAll(".bulk-tin-label").forEach((node) => {
        node.textContent = t("tinList");
      });
      document.querySelectorAll(".bulk-country-label").forEach((node) => {
        node.textContent = t("country");
      });
      document.querySelectorAll(".remove-group").forEach((node) => {
        node.textContent = t("remove");
      });
      refreshCountrySelects();
      if (!$("single-output").querySelector(".result-card")) setPlaceholder("single-output", "singlePlaceholder");
      if (!state.lastBulkPayload) {
        setPlaceholder("bulk-summary", "summaryPlaceholder");
        setPlaceholder("bulk-output", "detailsPlaceholder");
      } else {
        renderBulk(state.lastBulkPayload);
      }
    }

    async function postJson(url, payload) {
      const response = await fetch(url, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
      const data = await response.json();
      if (!response.ok) throw new Error(data.error || t("genericError"));
      return data;
    }

    function resultMessage(result) {
      return result.isValid ? result.formatInfo : (result.errorMsg || result.formatInfo || "");
    }

    function renderResult(result) {
      const ok = Boolean(result.isValid);
      return `
        <div class="result-card">
          <div class="result-head">
            <strong>${escapeHtml(result.countryCode)}</strong>
            <span class="badge ${ok ? "ok" : "bad"}">${escapeHtml(ok ? t("validStatus") : t("invalidStatus"))}</span>
          </div>
          <p class="mono">${escapeHtml(result.vkn)}</p>
          <p>${escapeHtml(resultMessage(result) || "-")}</p>
        </div>`;
    }

    function renderBulk(payload) {
      state.lastBulkPayload = payload;
      const summary = payload.summary || { total: 0, valid: 0, invalid: 0 };
      $("bulk-summary").innerHTML = `
        <div class="summary-card">
          <span>${escapeHtml(t("total"))}: <strong>${summary.total}</strong></span>
          <span>${escapeHtml(t("valid"))}: <strong>${summary.valid}</strong></span>
          <span>${escapeHtml(t("invalid"))}: <strong>${summary.invalid}</strong></span>
        </div>`;

      const groups = (payload.groups || []).map((group) => {
        const rows = (group.results || []).map((result) => {
          const ok = Boolean(result.isValid);
          return `<tr>
            <td class="mono">${escapeHtml(result.vkn)}</td>
            <td><span class="badge ${ok ? "ok" : "bad"}">${escapeHtml(ok ? t("validStatus") : t("invalidStatus"))}</span></td>
            <td>${escapeHtml(resultMessage(result) || "-")}</td>
          </tr>`;
        }).join("");
        return `
          <section class="country-result">
            <h3>${escapeHtml(group.countryCode || "-")}</h3>
            ${rows
              ? `<table><thead><tr><th>${escapeHtml(t("tin"))}</th><th>${escapeHtml(t("status"))}</th><th>${escapeHtml(t("description"))}</th></tr></thead><tbody>${rows}</tbody></table>`
              : `<span class="muted">${escapeHtml(t("noRows"))}</span>`}
          </section>`;
      }).join("");

      $("bulk-output").innerHTML = groups || `<span class="muted">${escapeHtml(t("noRows"))}</span>`;
    }

    function addBulkGroup(defaultCode = "") {
      const container = $("bulk-groups");
      const section = document.createElement("section");
      section.className = "bulk-group";
      section.innerHTML = `
        <div class="bulk-head">
          <label>
            <span class="bulk-country-label">${escapeHtml(t("country"))}</span>
            <select class="bulk-country"></select>
          </label>
          <button type="button" class="remove-group">${escapeHtml(t("remove"))}</button>
        </div>
        <label>
          <span class="bulk-tin-label">${escapeHtml(t("tinList"))}</span>
          <textarea class="bulk-values" placeholder="${escapeHtml(t("tinListPlaceholder"))}"></textarea>
        </label>`;
      const select = section.querySelector(".bulk-country");
      select.innerHTML = optionsHtml(defaultCode || countries()[0]?.code || "");
      section.querySelector(".remove-group").addEventListener("click", () => {
        section.remove();
        if (!document.querySelector(".bulk-group")) addBulkGroup();
      });
      container.appendChild(section);
    }

    function gatherBulkEntries() {
      return Array.from(document.querySelectorAll(".bulk-group")).map((group) => ({
        countryCode: group.querySelector(".bulk-country")?.value || "",
        values: (group.querySelector(".bulk-values")?.value || "")
          .split(/\r?\n/)
          .map((item) => item.trim())
          .filter(Boolean),
      })).filter((entry) => entry.countryCode && entry.values.length);
    }

    function downloadCsv() {
      const payload = state.lastBulkPayload;
      if (!payload || !payload.groups?.length) {
        $("error").textContent = t("csvNoData");
        return;
      }
      const lines = [["countryCode", "tin", "isValid", "message"]];
      for (const group of payload.groups) {
        for (const result of group.results || []) {
          lines.push([result.countryCode, result.vkn, result.isValid, resultMessage(result)]);
        }
      }
      const csv = lines.map((row) => row.map((cell) => `"${String(cell ?? "").replace(/"/g, '""')}"`).join(",")).join("\n");
      const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = "tin-checker-results.csv";
      link.click();
      URL.revokeObjectURL(url);
    }

    function init() {
      $("single-country").innerHTML = optionsHtml();
      addBulkGroup();
      updateTexts();

      document.querySelectorAll(".lang-btn").forEach((button) => {
        button.addEventListener("click", () => {
          state.lang = button.dataset.lang;
          updateTexts();
        });
      });

      document.querySelectorAll(".tab").forEach((tab) => {
        tab.addEventListener("click", () => {
          document.querySelectorAll(".tab").forEach((item) => item.classList.remove("active"));
          document.querySelectorAll(".panel").forEach((item) => item.classList.remove("active"));
          tab.classList.add("active");
          $(tab.dataset.panel).classList.add("active");
        });
      });

      $("single-form").addEventListener("submit", async (event) => {
        event.preventDefault();
        $("error").textContent = "";
        try {
          const payload = await postJson("/api/validate/single", {
            lang: state.lang,
            countryCode: $("single-country").value,
            tin: $("single-tin").value,
          });
          $("single-output").className = "";
          $("single-output").innerHTML = renderResult(payload);
        } catch (error) {
          $("error").textContent = error.message;
        }
      });

      $("reset-single").addEventListener("click", resetSingle);
      $("add-country").addEventListener("click", () => addBulkGroup());
      $("reset-bulk").addEventListener("click", resetBulk);
      $("download-csv").addEventListener("click", downloadCsv);

      $("run-bulk").addEventListener("click", async () => {
        $("error").textContent = "";
        try {
          const payload = await postJson("/api/validate/bulk", {
            lang: state.lang,
            entries: gatherBulkEntries(),
          });
          $("bulk-summary").className = "";
          $("bulk-output").className = "";
          renderBulk(payload);
        } catch (error) {
          $("error").textContent = error.message;
        }
      });

    }

    window.addEventListener("DOMContentLoaded", init);
  </script>
</body>
</html>
"""


def create_app(shutdown_callback=None) -> Flask:
    app = Flask(__name__)

    @app.get("/")
    def index():
        return render_template_string(
            HTML,
            countries_by_lang=json.dumps(
                {"tr": load_countries("tr"), "en": load_countries("en")},
                ensure_ascii=False,
            ),
            i18n=json.dumps(I18N, ensure_ascii=False),
        )

    @app.get("/api/countries")
    def countries():
        lang = (request.args.get("lang") or "en").strip().lower()
        return jsonify({"countries": load_countries("tr" if lang == "tr" else "en")})

    @app.post("/api/validate/single")
    def validate_single():
        payload = request.get_json(force=True)
        lang = (payload.get("lang") or "en").strip().lower()
        return jsonify(validate_tin(payload.get("countryCode", ""), payload.get("tin", ""), lang))

    @app.post("/api/validate/bulk")
    def validate_bulk():
        payload = request.get_json(force=True)
        lang = (payload.get("lang") or "en").strip().lower()
        entries = payload.get("entries")
        if entries is None:
            entries = [{"countryCode": payload.get("countryCode", ""), "values": payload.get("values") or []}]
        return jsonify(validate_bulk_entries(entries, lang))

    @app.post("/api/shutdown")
    def shutdown():
        if shutdown_callback:
            threading.Thread(target=shutdown_callback, daemon=True).start()
        return jsonify({"ok": True})

    return app


def _find_url(server) -> str:
    return f"http://127.0.0.1:{server.server_port}/"


def run_desktop() -> None:
    holder = {}

    def stop_server() -> None:
        time.sleep(0.25)
        holder["server"].shutdown()

    app = create_app(stop_server)
    requested_port = int(os.getenv("TIN_CHECKER_PORT", "8765"))
    try:
        server = make_server("127.0.0.1", requested_port, app)
    except OSError:
        server = make_server("127.0.0.1", 0, app)
    holder["server"] = server
    url = _find_url(server)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    if os.getenv("TIN_CHECKER_NO_BROWSER") != "1":
        webbrowser.open(url)
    try:
        while True:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as probe:
                if probe.connect_ex(("127.0.0.1", server.server_port)) != 0:
                    break
            time.sleep(0.4)
    except KeyboardInterrupt:
        server.shutdown()


def main() -> None:
    run_desktop()


if __name__ == "__main__":
    main()
