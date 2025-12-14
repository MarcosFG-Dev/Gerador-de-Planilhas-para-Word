(() => {
  const LS_KEY = "relatorio_pro_site_v1";


  const goApp = document.getElementById("goApp");
  const goApp2 = document.getElementById("goApp2");
  const demoFill = document.getElementById("demoFill");


  const tbody = document.getElementById("tbody");
  const byTypeList = document.getElementById("byTypeList");
  const grandTotalEl = document.getElementById("grandTotal");
  const rowCountEl = document.getElementById("rowCount");
  const autosaveStateEl = document.getElementById("autosaveState");
  const alertBox = document.getElementById("alertBox");


  const exportWordBtn = document.getElementById("exportWord");

  const clearBtn = document.getElementById("clearAll");
  const validateBtn = document.getElementById("validate");


  const docTitleEl = document.getElementById("docTitle");
  const clienteEl = document.getElementById("cliente");
  const cnpjEl = document.getElementById("cnpj");
  const projetoEl = document.getElementById("projeto");
  const responsavelEl = document.getElementById("responsavel");
  const dataDocEl = document.getElementById("dataDoc");
  const observacoesEl = document.getElementById("observacoes");
  const condicoesEl = document.getElementById("condicoes");

  const INITIAL_ROWS = 10;


  function $(id) { return document.getElementById(id); }

  function brMoney(n) {
    return (Number(n) || 0).toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
  }

  function parseNumberBR(str) {
    const s = String(str ?? "").trim();
    if (!s) return 0;
    const cleaned = s.replace(/\s/g, "").replace(/[R$\u00A0]/g, "");
    if (cleaned.includes(",")) {
      const noThousands = cleaned.replace(/\./g, "");
      return Number(noThousands.replace(",", ".")) || 0;
    }
    return Number(cleaned) || 0;
  }


  function parseMoneyInput(raw) {
    const s = String(raw ?? "").trim();
    if (!s) return 0;
    if (!/\d/.test(s)) return 0;
    if (/^\d+$/.test(s)) return Number(s) / 100;

    const filtered = s.replace(/[^\d,.\-]/g, "");
    if (filtered.includes(",")) {
      const noThousands = filtered.replace(/\./g, "");
      return Number(noThousands.replace(",", ".")) || 0;
    }
    return Number(filtered) || 0;
  }

  function normalizeTipo(s) {
    return String(s ?? "").trim().toLowerCase();
  }

  function escapeHtml(s) {
    return String(s ?? "")
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;").replace(/'/g, "&#039;");
  }

  function fileSafeName(name) {
    const raw = String(name || "relatorio").trim();
    return raw.replace(/[\\/:*?"<>|]+/g, "-").replace(/\s+/g, " ").trim() || "relatorio";
  }

  function todayISO() {
    const d = new Date();
    const pad = n => String(n).padStart(2, "0");
    return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  }


  function getRowCells(tr) {
    return {
      tipo: tr.querySelector('[data-col="tipo"]'),
      qnt: tr.querySelector('[data-col="qnt"]'),
      valor: tr.querySelector('[data-col="valor"]'),
      total: tr.querySelector('[data-col="total"]'),
    };
  }

  function rowIsFilled(tr) {
    const { tipo, qnt, valor } = getRowCells(tr);
    return (tipo.textContent.trim() !== "" ||
      qnt.textContent.trim() !== "" ||
      valor.textContent.trim() !== "");
  }

  function recalcRow(tr) {
    const { qnt, valor, total } = getRowCells(tr);
    const q = parseNumberBR(qnt.textContent);
    const v = parseMoneyInput(valor.textContent);
    total.textContent = brMoney(q * v);
  }

  function ensureLastRowAlwaysExists() {
    const rows = [...tbody.querySelectorAll("tr")];
    const last = rows[rows.length - 1];
    if (last && rowIsFilled(last)) {
      tbody.appendChild(createRow());
    }
  }

  function placeCaretAtEnd(el) {
    try {
      const range = document.createRange();
      range.selectNodeContents(el);
      range.collapse(false);
      const sel = window.getSelection();
      sel.removeAllRanges();
      sel.addRange(range);
    } catch { }
  }

  function focusCell(tr, col) {
    const el = tr?.querySelector(`[data-col="${col}"]`);
    if (el && el.getAttribute("contenteditable") === "true") el.focus();
  }

  function handleEnter(tr, col) {
    const rows = [...tbody.querySelectorAll("tr")];
    const idx = rows.indexOf(tr);
    const next = rows[idx + 1];
    if (next) {
      focusCell(next, col === "total" ? "tipo" : col);
    }
  }


  function splitRowCells(line) {
    if (line.includes("\t")) return line.split("\t");
    return line.split(";");
  }

  function applyPaste(startTr, startCol, text) {
    const lines = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n");
    const startCols = ["tipo", "qnt", "valor"];
    let colIdx = Math.max(0, startCols.indexOf(startCol));
    if (colIdx === -1) colIdx = 0;

    let rIndex = [...tbody.querySelectorAll("tr")].indexOf(startTr);

    for (let i = 0; i < lines.length; i++) {
      if (lines[i] === "") continue;
      const parts = splitRowCells(lines[i]);

      while (tbody.querySelectorAll("tr").length <= rIndex) {
        tbody.appendChild(createRow());
      }
      const tr = tbody.querySelectorAll("tr")[rIndex];

      for (let j = 0; j < parts.length; j++) {
        const targetCol = startCols[colIdx + j];
        if (!targetCol) break;
        const cell = tr.querySelector(`[data-col="${targetCol}"]`);
        if (!cell) continue;

        const val = parts[j].trim();
        if (targetCol === "valor") {
          const num = parseMoneyInput(val);
          cell.textContent = num ? brMoney(num) : "";
        } else {
          cell.textContent = val;
        }
      }

      recalcRow(tr);
      rIndex++;
    }

    ensureLastRowAlwaysExists();
    recalcAll();
    scheduleAutosave();
  }

  function createRow() {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td><div class="cell" contenteditable="true" data-col="tipo"></div></td>
      <td><div class="cell r" contenteditable="true" data-col="qnt"></div></td>
      <td><div class="cell r" contenteditable="true" data-col="valor"></div></td>
      <td><div class="cell r" contenteditable="false" data-col="total">${brMoney(0)}</div></td>
    `;

    // Melhoria na digitação de preço: Focus = Limpo / Blur = Formatado
    const valCell = tr.querySelector('[data-col="valor"]');

    valCell.addEventListener("focus", (e) => {
      // Ao focar, mostra o número puro (ex: "12,90") para facilitar edição
      const current = parseMoneyInput(e.target.textContent);
      if (current === 0) {
        e.target.textContent = "";
      } else {
        // Mantém formatação decimal padrão PT-BR para edição (1.234,56) sem o "R$"
        e.target.textContent = current.toLocaleString("pt-BR", { minimumFractionDigits: 2 });
      }
      // Seleciona tudo para facilitar substituição
      setTimeout(() => {
        try { document.execCommand('selectAll', false, null); } catch { }
      }, 10);
    });

    valCell.addEventListener("blur", (e) => {
      // Ao sair, formata bonitinho em R$
      const num = parseMoneyInput(e.target.textContent);
      e.target.textContent = num ? brMoney(num) : "";
      recalcRow(tr);
      ensureLastRowAlwaysExists();
      recalcAll();
      scheduleAutosave();
    });

    tr.addEventListener("input", (e) => {
      const col = e.target?.dataset?.col;
      if (!col) return;

      // Se for valor, apenas recalcula totais (sem formatar o input do usuário)
      recalcRow(tr);
      ensureLastRowAlwaysExists();
      recalcAll();
      scheduleAutosave();
    });

    tr.addEventListener("keydown", (e) => {
      const col = e.target?.dataset?.col;
      if (!col) return;

      if (e.key === "Enter") {
        e.preventDefault();
        handleEnter(tr, col);
      } else if (e.key === "Tab") {
        // Deixa o Tab nativo funcionar, mas garante foco no próximo
        // O CSS user-select já deve cuidar, mas aqui garantimos o fluxo lógico
      }
    });

    tr.querySelectorAll('[contenteditable="true"]').forEach(cell => {
      cell.addEventListener("paste", (e) => {
        const text = (e.clipboardData || window.clipboardData).getData("text");
        if (!text) return;
        if (text.includes("\n") || text.includes("\t") || text.includes(";")) {
          e.preventDefault();
          applyPaste(tr, cell.dataset.col, text);
        }
      });
    });

    return tr;
  }


  function computeAggregates() {
    let grand = 0;
    let filled = 0;
    const map = new Map();

    [...tbody.querySelectorAll("tr")].forEach(tr => {
      recalcRow(tr);
      if (!rowIsFilled(tr)) return;

      filled++;
      const { tipo, total } = getRowCells(tr);
      const t = parseNumberBR(total.textContent);
      grand += t;

      const key = normalizeTipo(tipo.textContent);
      if (!key) return;

      const label = tipo.textContent.trim();
      const cur = map.get(key) ?? { label, sum: 0 };
      if (!cur.label) cur.label = label;
      cur.sum += t;
      map.set(key, cur);
    });

    const items = [...map.values()].sort((a, b) => b.sum - a.sum);
    return { grand, filled, items };
  }

  function recalcAll() {
    const { grand, filled, items } = computeAggregates();
    grandTotalEl.textContent = brMoney(grand);
    rowCountEl.textContent = String(filled);

    byTypeList.innerHTML = "";
    if (items.length === 0) {
      byTypeList.innerHTML = `<div class="byTypeItem"><span class="muted">Sem dados ainda</span><span class="muted">—</span></div>`;
      return;
    }
    for (const it of items) {
      const div = document.createElement("div");
      div.className = "byTypeItem";
      div.innerHTML = `<b>${escapeHtml(it.label)}</b><span class="v">${escapeHtml(brMoney(it.sum))}</span>`;
      byTypeList.appendChild(div);
    }
  }


  function validateAll(showAlert = true) {
    const problems = [];

    const title = docTitleEl.value.trim();
    const cliente = clienteEl.value.trim();
    const data = dataDocEl.value;

    if (!title) problems.push("Título do documento é obrigatório.");
    if (!cliente) problems.push("Cliente é obrigatório.");
    if (!data) problems.push("Data é obrigatória.");


    const cnpj = cnpjEl.value.trim();
    if (cnpj) {
      const digits = cnpj.replace(/[^\d]/g, "");
      if (digits.length !== 14) problems.push("CNPJ parece inválido (precisa ter 14 dígitos).");
    }


    const rows = [...tbody.querySelectorAll("tr")].filter(rowIsFilled);
    if (rows.length === 0) problems.push("Inclua pelo menos 1 item na tabela.");

    rows.forEach((tr, idx) => {
      const { tipo, qnt, valor } = getRowCells(tr);
      const tipoTxt = tipo.textContent.trim();
      const q = parseNumberBR(qnt.textContent);
      const v = parseMoneyInput(valor.textContent);

      if (!tipoTxt) problems.push(`Linha ${idx + 1}: TIPO é obrigatório.`);
      if (q <= 0) problems.push(`Linha ${idx + 1}: QNT deve ser maior que 0.`);
      if (v <= 0) problems.push(`Linha ${idx + 1}: VALOR deve ser maior que 0.`);
    });

    if (showAlert) {
      if (problems.length) {
        alertBox.hidden = false;
        alertBox.textContent = "Verifique:\n• " + problems.join("\n• ");
      } else {
        alertBox.hidden = false;
        alertBox.style.borderColor = "rgba(46, 204, 113, .35)";
        alertBox.style.background = "rgba(46, 204, 113, .10)";
        alertBox.style.color = "#caffdf";
        alertBox.textContent = "Tudo certo! Documento válido para exportação.";
        setTimeout(() => {
          alertBox.hidden = true;
          alertBox.removeAttribute("style");
        }, 2200);
      }
    }

    return { ok: problems.length === 0, problems };
  }


  let autosaveTimer = null;

  function snapshot() {
    const rows = [...tbody.querySelectorAll("tr")]
      .filter(tr => rowIsFilled(tr))
      .map(tr => {
        const { tipo, qnt, valor, total } = getRowCells(tr);
        return {
          tipo: tipo.textContent.trim(),
          qnt: qnt.textContent.trim(),
          valor: valor.textContent.trim(),
          total: total.textContent.trim(),
        };
      });

    return {
      v: 1,
      savedAt: Date.now(),
      fields: {
        docTitle: docTitleEl.value,
        cliente: clienteEl.value,
        cnpj: cnpjEl.value,
        projeto: projetoEl.value,
        responsavel: responsavelEl.value,
        dataDoc: dataDocEl.value,
        observacoes: observacoesEl.value,
        condicoes: condicoesEl.value,
      },
      rows
    };
  }

  function scheduleAutosave() {
    if (autosaveTimer) clearTimeout(autosaveTimer);
    autosaveTimer = setTimeout(() => {
      localStorage.setItem(LS_KEY, JSON.stringify(snapshot()));
      autosaveStateEl.textContent = "Auto-save: " + new Date().toLocaleTimeString("pt-BR");
      autosaveTimer = null;
    }, 250);
  }

  function restore(state) {
    if (state?.fields) {
      docTitleEl.value = state.fields.docTitle ?? "";
      clienteEl.value = state.fields.cliente ?? "";
      cnpjEl.value = state.fields.cnpj ?? "";
      projetoEl.value = state.fields.projeto ?? "";
      responsavelEl.value = state.fields.responsavel ?? "";
      dataDocEl.value = state.fields.dataDoc ?? "";
      observacoesEl.value = state.fields.observacoes ?? "";
      condicoesEl.value = state.fields.condicoes ?? "";
    }

    tbody.innerHTML = "";
    const list = Array.isArray(state?.rows) ? state.rows : [];
    const count = Math.max(INITIAL_ROWS, list.length + 1);

    for (let i = 0; i < count; i++) {
      const tr = createRow();
      tbody.appendChild(tr);
      if (i < list.length) {
        const { tipo, qnt, valor } = getRowCells(tr);
        tipo.textContent = list[i].tipo ?? "";
        qnt.textContent = list[i].qnt ?? "";
        const money = parseMoneyInput(list[i].valor ?? "");
        valor.textContent = (list[i].valor && /R\$\s?/.test(list[i].valor)) ? list[i].valor : (money ? brMoney(money) : "");
        recalcRow(tr);
      }
    }

    ensureLastRowAlwaysExists();
    recalcAll();
  }

  function loadState() {
    try {
      const raw = localStorage.getItem(LS_KEY);
      if (!raw) return null;
      return JSON.parse(raw);
    } catch { return null; }
  }


  function escapeWordChars(s) {
    return String(s ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  }


  function buildWordHTML() {
    const title = docTitleEl.value.trim();
    const cliente = clienteEl.value.trim();
    const cnpj = cnpjEl.value.trim();
    const projeto = projetoEl.value.trim();
    const responsavel = responsavelEl.value.trim();


    let dataFormatada = "";
    if (dataDocEl.value) {
      const [ano, mes, dia] = dataDocEl.value.split("-");
      dataFormatada = `${dia}/${mes}/${ano}`;
    }

    const obs = observacoesEl.value.trim();
    const cond = condicoesEl.value.trim();

    const rows = [...tbody.querySelectorAll("tr")].filter(rowIsFilled).map(tr => {
      const { tipo, qnt, valor, total } = getRowCells(tr);
      return {
        tipo: escapeWordChars(tipo.textContent.trim()),
        qnt: escapeWordChars(qnt.textContent.trim()),
        valor: escapeWordChars(valor.textContent.trim()),
        total: escapeWordChars(total.textContent.trim()),
      };
    });

    const { grand, items } = computeAggregates();
    const grandFormatted = brMoney(grand);


    const tableRows = rows.length
      ? rows.map(r => `
        <tr>
          <td>${r.tipo}</td>
          <td class="r">${r.qnt}</td>
          <td class="r">${r.valor}</td>
          <td class="r strong">${r.total}</td>
        </tr>`).join("")
      : `<tr><td colspan="4" style="color:#999; text-align:center; padding:20px;">Nenhum item adicionado.</td></tr>`;


    const summaryRows = items.length
      ? items.map(it => `
          <tr>
            <td class="label">${escapeWordChars(it.label)}</td>
            <td class="value">${escapeWordChars(brMoney(it.sum))}</td>
          </tr>
        `).join("")
      : "";


    const css = `
      @page { margin: 1.0cm; }
      body { font-family: 'Segoe UI', Helvetica, Arial, sans-serif; font-size: 9pt; color: #333; margin: 0; background: #fff; line-height: 1.3; }
      
      /* Header Compacto */
      .dash-header-bg {
        background: #1e293b;
        color: white; 
        padding: 15px 20px; 
        border-bottom: 4px solid #10b981; 
        margin-bottom: 15px;
        display: flex; justify-content: space-between; align-items: center;
      }
      .dash-title { font-size: 16pt; font-weight: 800; text-transform: uppercase; margin: 0; }
      .dash-subtitle { color: #cbd5e1; font-size: 8.5pt; margin-top: 2px; }
      .status-pill {
        background: #059669; color: #fff; 
        padding: 4px 10px; border-radius: 4px; 
        font-weight: bold; font-size: 8pt; 
        border: 1px solid #34d399;
        float: right;
      }

      /* KPI Cards Compactos */
      .kpi-grid { width: 100%; border-collapse: separate; border-spacing: 10px 0; margin-bottom: 20px; }
      .kpi-box { 
        padding: 10px 12px; 
        border-radius: 6px; 
        color: #fff; 
        box-shadow: 2px 2px 5px rgba(0,0,0,0.1);
        position: relative; overflow: hidden;
      }
      .kpi-blue { background: #2563eb; border-left: 4px solid #1e40af; }
      .kpi-indigo { background: #4f46e5; border-left: 4px solid #3730a3; }
      .kpi-emerald { background: #059669; border-left: 4px solid #064e3b; }
      
      .kpi-label { font-size: 7pt; opacity: 0.9; text-transform: uppercase; font-weight: bold; margin-bottom: 2px; }
      .kpi-val { font-size: 11pt; font-weight: bold; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
      .kpi-sub { font-size: 8pt; opacity: 0.9; margin-top: 2px; border-top: 1px solid rgba(255,255,255,0.2); padding-top: 2px; }

      /* Tabela Densa */
      .grid-table { width: 100%; border-collapse: collapse; margin-bottom: 15px; font-size: 8.5pt; }
      .grid-table th { 
        background: #f1f5f9; 
        color: #475569; 
        font-size: 8pt; 
        text-transform: uppercase; 
        font-weight: 700; 
        text-align: left; 
        padding: 6px 8px; 
        border-bottom: 2px solid #cbd5e1;
      }
      .grid-table td { 
        padding: 5px 8px; 
        border-bottom: 1px solid #e2e8f0; 
        vertical-align: middle;
      }
      .tag { 
        display: inline-block; padding: 2px 6px; border-radius: 4px; 
        font-size: 7pt; font-weight: 700; background: #e0f2fe; color: #0369a1; border: 1px solid #bae6fd;
        text-transform: uppercase;
      }
      .bar-track { width: 50px; height: 5px; background: #e2e8f0; display: inline-block; vertical-align: middle; margin-right: 5px; }
      .bar-fill { height: 100%; background: #3b82f6; }

      /* Rodapé Compacto */
      .footer-grid { width: 100%; margin-top: 10px; border-top: 2px solid #e2e8f0; padding-top: 10px; }
      .footer-box { 
        background: #fffbeb; 
        border: 1px solid #fcd34d; 
        padding: 8px; 
        border-radius: 4px; 
        font-size: 8pt; color: #92400e;
      }
      .totals-card { background: #f8fafc; padding: 10px; border-radius: 4px; border: 1px solid #e2e8f0; }
      .sums-table td { padding: 2px 0; font-size: 8.5pt; }
      .dash-total { font-size: 12pt; color: #0f172a; font-weight: 800; }
    `;


    const maxVal = Math.max(...rows.map(r => parseMoneyInput(r.total)), 1);

    const richTableRows = rows.map(r => {
      const valNum = parseMoneyInput(r.total);
      const pct = Math.min(100, Math.round((valNum / maxVal) * 100));
      return `
        <tr>
          <td width="15%"><span class="tag">${escapeWordChars(r.tipo.slice(0, 12))}</span></td>
          <td width="30%"><b>${escapeWordChars(r.tipo)}</b></td>
          <td width="10%" class="r">${r.qnt}</td>
          <td width="15%" class="r">${r.valor}</td>
          <td width="15%" class="r"><b>${r.total}</b></td>
          <td width="15%">
            <div class="bar-track"><div class="bar-fill" style="width:${pct}%;"></div></div>
            <span style="font-size:7pt; color:#64748b;">${pct}%</span>
          </td>
        </tr>
      `;
    }).join("");

    return `
      <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
      <head>
        <meta charset="utf-8">
        <style>${css}</style>
      </head>
      <body>
        
        <!-- Header -->
        <table width="100%" class="dash-header-bg">
          <tr>
            <td>
              <div class="dash-title">${escapeWordChars(title)}</div>
              <div class="dash-subtitle">ID: #${Date.now().toString().slice(-6)} • ${dataFormatada || "—"}</div>
            </td>
            <td align="right">
              <span class="status-pill">EMITIDO</span>
            </td>
          </tr>
        </table>

        <!-- KPI Cards -->
        <table class="kpi-grid">
          <tr>
            <td width="33%">
              <div class="kpi-box kpi-blue">
                <div class="kpi-label">CLIENTE</div>
                <div class="kpi-val">${escapeWordChars(cliente.slice(0, 25))}</div>
                <div class="kpi-sub">${cnpj ? escapeWordChars(cnpj) : "—"}</div>
              </div>
            </td>
            <td width="33%">
              <div class="kpi-box kpi-indigo">
                <div class="kpi-label">PROJETO</div>
                <div class="kpi-val">${projeto ? escapeWordChars(projeto.slice(0, 25)) : "Geral"}</div>
                <div class="kpi-sub">${responsavel ? escapeWordChars(responsavel) : "—"}</div>
              </div>
            </td>
            <td width="33%">
              <div class="kpi-box kpi-emerald">
                <div class="kpi-label">TOTAL</div>
                <div class="kpi-val">${escapeWordChars(grandFormatted)}</div>
                <div class="kpi-sub">${rows.length} itens</div>
              </div>
            </td>
          </tr>
        </table>

        <!-- Tabela -->
        <div style="margin-bottom:5px; font-weight:700; color:#64748b; font-size:8pt; text-transform:uppercase;">Itens</div>
        <table class="grid-table">
          <thead>
            <tr>
              <th>TAG</th>
              <th>DESCRIÇÃO</th>
              <th class="r">QTD</th>
              <th class="r">VALOR</th>
              <th class="r">TOTAL</th>
              <th>IMP.</th>
            </tr>
          </thead>
          <tbody>
            ${richTableRows.length ? richTableRows : `<tr><td colspan="6" style="padding:10px; text-align:center;">Vazio</td></tr>`}
          </tbody>
        </table>

        <!-- Footer Grid -->
        <table class="footer-grid">
          <tr>
            <td width="60%" valign="top" style="padding-right:15px;">
              ${(obs || cond) ? `
                <div class="footer-box">
                  ${obs ? `<div><b>OBS:</b> ${escapeWordChars(obs)}</div>` : ""}
                  ${cond ? `<div style="margin-top:5px;"><b>COND:</b> ${escapeWordChars(cond)}</div>` : ""}
                </div>
              ` : ""}
              <div style="margin-top:15px; font-size:7pt; color:#94a3b8;">
                Relatório Pro • Gerado automaticamente
              </div>
            </td>
            <td width="40%" valign="top">
              <div class="totals-card">
                <table class="sums-table">
                  ${items.map(it => `
                    <tr>
                      <td style="color:#64748b;">${escapeWordChars(it.label)}</td>
                      <td style="text-align:right; font-weight:600;">${escapeWordChars(brMoney(it.sum))}</td>
                    </tr>
                  `).join("")}
                  <tr><td colspan="2" style="border-top:1px solid #e2e8f0; margin:5px 0;"></td></tr>
                  <tr>
                    <td style="font-weight:700;">TOTAL</td>
                    <td style="text-align:right;" class="dash-total">${escapeWordChars(grandFormatted)}</td>
                  </tr>
                </table>
              </div>
            </td>
          </tr>
        </table>

      </body>
      </html>
    `;
  }



  function downloadWord(html) {
    const blob = new Blob(["\ufeff", html], { type: "application/msword;charset=utf-8" });
    const url = URL.createObjectURL(blob);

    const a = document.createElement("a");
    a.href = url;
    a.download = `${fileSafeName(docTitleEl.value)}.doc`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }


  function scrollToApp() {
    document.getElementById("app").scrollIntoView({ behavior: "smooth", block: "start" });
    setTimeout(() => docTitleEl.focus(), 250);
  }

  goApp?.addEventListener("click", scrollToApp);
  goApp2?.addEventListener("click", scrollToApp);

  function setDemo() {
    docTitleEl.value = "Relatório de Itens";
    clienteEl.value = "ACME Ltda";
    cnpjEl.value = "00.000.000/0000-00";
    projetoEl.value = "Reforma";
    responsavelEl.value = "Marcos";
    dataDocEl.value = todayISO();
    observacoesEl.value = "Valores estimados. Conferir quantidades na entrega.";
    condicoesEl.value = "Validade: 7 dias • Pagamento: 50% entrada / 50% entrega";

    tbody.innerHTML = "";
    const data = [
      ["Material", "10", "1290", ""],
      ["Serviço", "2", "R$ 450,00", ""],
      ["Transporte", "1", "12000", ""],
    ];
    for (let i = 0; i < Math.max(INITIAL_ROWS, data.length + 1); i++) {
      const tr = createRow();
      tbody.appendChild(tr);
      if (i < data.length) {
        const { tipo, qnt, valor } = getRowCells(tr);
        tipo.textContent = data[i][0];
        qnt.textContent = data[i][1];
        const v = parseMoneyInput(data[i][2]);
        valor.textContent = v ? brMoney(v) : "";
        recalcRow(tr);
      }
    }
    ensureLastRowAlwaysExists();
    recalcAll();
    scheduleAutosave();
    scrollToApp();
  }

  demoFill?.addEventListener("click", setDemo);


  [docTitleEl, clienteEl, cnpjEl, projetoEl, responsavelEl, dataDocEl, observacoesEl, condicoesEl]
    .forEach(el => el.addEventListener("input", scheduleAutosave));

  validateBtn.addEventListener("click", () => validateAll(true));

  exportWordBtn.addEventListener("click", () => {
    const { ok } = validateAll(true);
    if (!ok) return;
    const html = buildWordHTML();
    downloadWord(html);
  });



  clearBtn.addEventListener("click", () => {
    localStorage.removeItem(LS_KEY);
    alertBox.hidden = true;

    docTitleEl.value = "";
    clienteEl.value = "";
    cnpjEl.value = "";
    projetoEl.value = "";
    responsavelEl.value = "";
    dataDocEl.value = todayISO();
    observacoesEl.value = "";
    condicoesEl.value = "";

    tbody.innerHTML = "";
    for (let i = 0; i < INITIAL_ROWS; i++) tbody.appendChild(createRow());
    ensureLastRowAlwaysExists();
    recalcAll();
    autosaveStateEl.textContent = "Auto-save: limpo";
  });


  function init() {

    if (!dataDocEl.value) dataDocEl.value = todayISO();

    const state = loadState();
    if (state) {
      restore(state);
      autosaveStateEl.textContent = "Auto-save: restaurado";
    } else {
      tbody.innerHTML = "";
      for (let i = 0; i < INITIAL_ROWS; i++) tbody.appendChild(createRow());
      ensureLastRowAlwaysExists();
      recalcAll();
      autosaveStateEl.textContent = "Auto-save: pronto";
    }

    // quick: if empty title, set a good default
    if (!docTitleEl.value) docTitleEl.value = "Relatório de Itens";
    scheduleAutosave();
  }

  init();
})();
