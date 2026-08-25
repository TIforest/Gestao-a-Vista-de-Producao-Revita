(() => {
  "use strict";

  const state = {
    turma: null,
    maquina: null,
  };

  const REFRESH_MS = 30_000; // o servidor checa o SharePoint a cada 1 min; o painel rebusca o dashboard a cada 30s.

  function fmtMil(valor) {
    const mil = (valor || 0) / 1000;
    return (
      mil.toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 3 }) + " Mil"
    );
  }

  function fmtKg(valor) {
    return (valor || 0).toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  function fmtHora(isoDataHora) {
    const t = isoDataHora?.split("T")[1];
    return t ? t.slice(0, 8) : "—";
  }

  const MONTH_NAMES = [
    "JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO",
    "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO",
  ];

  async function apiGet(path) {
    const res = await fetch(path);
    if (!res.ok) throw new Error(`Erro ${res.status} ao consultar ${path}`);
    return res.json();
  }

  function renderHBarRow({ label, valor, meta, onClick, selected }) {
    const row = document.createElement("div");
    row.className = "hbar-row" + (onClick ? " clickable" : "") + (selected ? " selected" : "");

    const labelEl = document.createElement("div");
    labelEl.className = "hbar-row-label";
    labelEl.textContent = label;
    row.appendChild(labelEl);

    const track = document.createElement("div");
    track.className = "hbar-track";

    const pct = meta > 0 ? Math.min(100, (valor / meta) * 100) : 0;
    const fillValor = document.createElement("div");
    fillValor.className = "hbar-fill-valor";
    fillValor.style.width = pct + "%";
    track.appendChild(fillValor);

    const badge = document.createElement("div");
    badge.className = "hbar-value-badge";
    badge.style.left = `calc(${pct}% + 6px)`;
    badge.textContent = fmtMil(valor);
    track.appendChild(badge);

    row.appendChild(track);

    const metaLabel = document.createElement("div");
    metaLabel.className = "hbar-meta-label";
    metaLabel.textContent = fmtMil(meta);
    row.appendChild(metaLabel);

    if (onClick) row.addEventListener("click", onClick);
    return row;
  }

  function renderGauge(container, valor, meta, color) {
    container.innerHTML = "";
    const size = 160;
    const svgNS = "http://www.w3.org/2000/svg";
    const svg = document.createElementNS(svgNS, "svg");
    svg.setAttribute("viewBox", `0 0 ${size} ${size / 2 + 10}`);
    svg.setAttribute("width", "100%");
    svg.setAttribute("height", (size / 2 + 10) + "px");

    const cx = size / 2;
    const cy = size / 2;
    const r = size / 2 - 10;

    const bg = document.createElementNS(svgNS, "path");
    bg.setAttribute("d", describeArc(cx, cy, r, 180, 360));
    bg.setAttribute("stroke", "#b19873"); // bege kraft — trilho do gauge sempre visível, mesmo com valor 0
    bg.setAttribute("stroke-width", "16");
    bg.setAttribute("fill", "none");
    svg.appendChild(bg);

    const pct = meta > 0 ? Math.min(1, valor / meta) : 0;
    const endAngle = 180 + pct * 180;
    const fg = document.createElementNS(svgNS, "path");
    fg.setAttribute("d", describeArc(cx, cy, r, 180, endAngle));
    fg.setAttribute("stroke", color);
    fg.setAttribute("stroke-width", "16");
    fg.setAttribute("stroke-linecap", "round");
    fg.setAttribute("fill", "none");
    svg.appendChild(fg);

    container.appendChild(svg);

    // Texto sempre em verde-escuro: verde-claro/roxo-destaque têm baixo
    // contraste como texto sobre fundo branco, então servem só para o arco.
    const value = document.createElement("div");
    value.className = "gauge-value";
    value.textContent = fmtMil(valor);
    container.appendChild(value);

    const bounds = document.createElement("div");
    bounds.className = "gauge-bounds";
    bounds.innerHTML = `<span>0,000 Mil</span><span>${fmtMil(meta)}</span>`;
    container.appendChild(bounds);
  }

  function polarToCartesian(cx, cy, r, angleDeg) {
    const a = ((angleDeg - 90) * Math.PI) / 180;
    return { x: cx + r * Math.cos(a), y: cy + r * Math.sin(a) };
  }

  function describeArc(cx, cy, r, startAngle, endAngle) {
    const start = polarToCartesian(cx, cy, r, endAngle);
    const end = polarToCartesian(cx, cy, r, startAngle);
    const largeArc = endAngle - startAngle <= 180 ? "0" : "1";
    return `M ${start.x} ${start.y} A ${r} ${r} 0 ${largeArc} 0 ${end.x} ${end.y}`;
  }

  function renderMetaBar(container, pct) {
    container.innerHTML = "";
    const clamped = Math.max(0, Math.min(100, pct));
    const track = document.createElement("div");
    track.className = "meta-bar-track";

    const atingido = document.createElement("div");
    atingido.className = "meta-bar-atingido";
    atingido.style.width = clamped + "%";
    atingido.textContent = clamped.toFixed(2).replace(".", ",") + "%";
    track.appendChild(atingido);

    const falta = document.createElement("div");
    falta.className = "meta-bar-falta";
    falta.textContent = (100 - clamped).toFixed(2).replace(".", ",") + "%";
    track.appendChild(falta);

    container.appendChild(track);

    const ticks = document.createElement("div");
    ticks.className = "meta-bar-ticks";
    ticks.innerHTML = "<span>0%</span><span>50%</span><span>100%</span>";
    container.appendChild(ticks);
  }

  function renderTurmas(payload) {
    const list = document.getElementById("turmasList");
    list.innerHTML = "";

    const todas = document.createElement("div");
    todas.className = "turma-box todas" + (state.turma === null ? " active" : "");
    todas.textContent = "TODAS";
    todas.addEventListener("click", () => { state.turma = null; load(); });
    list.appendChild(todas);

    for (const t of payload.turmasDisponiveis) {
      const box = document.createElement("div");
      box.className = "turma-box" + (state.turma === t ? " active" : "");
      box.textContent = t;
      box.addEventListener("click", () => { state.turma = state.turma === t ? null : t; load(); });
      list.appendChild(box);
    }
  }

  function render(payload) {
    renderTurmas(payload);

    document.getElementById("producaoMesValor").textContent = fmtMil(payload.producaoMes);
    const [y, m] = payload.data.split("-");
    document.getElementById("mesLabel").textContent = `${MONTH_NAMES[Number(m) - 1]} ${y}`;

    const turmaChart = document.getElementById("producaoTurmaChart");
    turmaChart.innerHTML = "";
    const linhasTurma = state.turma
      ? payload.producaoPorTurma.filter((r) => r.turma === state.turma)
      : payload.producaoPorTurma;
    for (const r of linhasTurma) {
      turmaChart.appendChild(
        renderHBarRow({ label: r.turma, valor: r.valor, meta: payload.metaTurno || payload.metaDia })
      );
    }

    const desaguadorasChart = document.getElementById("desaguadorasChart");
    desaguadorasChart.innerHTML = "";
    const metaPorDesaguadora =
      payload.desaguadorasDisponiveis.length > 0
        ? (payload.metaTurno || payload.metaDia) / payload.desaguadorasDisponiveis.length
        : 0;
    for (const r of payload.producaoPorDesaguadora) {
      desaguadorasChart.appendChild(
        renderHBarRow({
          label: `DESAGUADORA ${r.maquina}`,
          valor: r.valor,
          meta: metaPorDesaguadora,
          selected: state.maquina === r.maquina,
          onClick: () => { state.maquina = state.maquina === r.maquina ? null : r.maquina; load(); },
        })
      );
    }

    const horaChart = document.getElementById("horaChart");
    horaChart.innerHTML = "";
    horaChart.appendChild(
      renderHBarRow({ label: state.turma || "TODAS", valor: payload.producaoMediaHora, meta: payload.metaHora })
    );

    renderGauge(document.getElementById("gaugeTurno"), payload.producaoTurno, payload.metaTurno, "#c5f249");
    renderGauge(document.getElementById("gaugeDia"), payload.producaoDia, payload.metaDia, "#c1b4ee");
    renderMetaBar(document.getElementById("metaAtingidaChart"), payload.percentualMetaAtingida);

    const tbody = document.querySelector("#tabelaApontamentos tbody");
    tbody.innerHTML = "";
    for (const a of payload.ultimosApontamentos) {
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${a.lote}</td>
        <td>${a.cliente}</td>
        <td>${a.numero_fardo ?? ""}</td>
        <td>${a.turma}</td>
        <td>${fmtKg(a.peso_seco)}</td>
        <td>${fmtHora(a.data_hora)}</td>
        <td>DESAGUADORA ${a.maquina}</td>
        <td>${a.produto}</td>`;
      tbody.appendChild(tr);
    }

    const syncInfo = document.getElementById("syncInfo");
    const statusLabel = {
      sincronizado: "sincronizado",
      sincronizado_manual: "sincronizado (upload manual)",
      sem_alteracao: "sem alterações no Excel",
      erro: "erro na sincronização",
      nunca_sincronizado: "ainda não sincronizado",
    }[payload.sync.status] || payload.sync.status;
    const quando = payload.sync.ultimaSincronizacao
      ? new Date(payload.sync.ultimaSincronizacao + "Z").toLocaleString("pt-BR")
      : "—";
    syncInfo.textContent = `Última sincronização: ${quando} · ${statusLabel}`;
    syncInfo.title = payload.sync.erro || "";
  }

  async function load() {
    const params = new URLSearchParams();
    if (state.turma) params.set("turma", state.turma);
    if (state.maquina) params.set("maquina", state.maquina);
    try {
      const payload = await apiGet("/api/dashboard?" + params.toString());
      render(payload);
    } catch (err) {
      console.error(err);
    }
  }

  document.getElementById("btnRefresh").addEventListener("click", load);

  load();
  setInterval(load, REFRESH_MS);
})();
