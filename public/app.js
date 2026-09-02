(() => {
  "use strict";

  const state = {
    turma: null,
    maquina: null,
  };

  const REFRESH_MS = 30_000; // o servidor checa o SharePoint a cada 1 min; o painel rebusca o dashboard a cada 30s.
  // 1 min: enquanto o cron do servidor está travado por um bug da própria
  // Cloudflare (ver SUPORTE_CLOUDFLARE.txt), o navegador que ficar com essa
  // aba aberta (a TV) vira o mecanismo principal de atualização automática.
  // Seguro rodar a cada 1 min mesmo assim: sem "force" (ver mais abaixo),
  // quando o arquivo não mudou isso é só 1 checagem barata, não reescreve
  // o banco à toa.
  const FORCE_SYNC_MS = 60_000;

  function fmtTon(valor) {
    const ton = (valor || 0) / 1000; // peso_seco vem em kg — dividir por 1000 já dá toneladas
    return (
      ton.toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 3 }) + " Ton"
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

  // Timeout pra nunca ficar "travado" sem feedback — se a rede/servidor
  // não responder em 12s, desiste e mostra erro em vez de ficar pendurado.
  async function fetchComTimeout(path, opts, timeoutMs = 12000) {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), timeoutMs);
    try {
      return await fetch(path, { ...opts, signal: controller.signal });
    } finally {
      clearTimeout(timer);
    }
  }

  async function apiGet(path) {
    const res = await fetchComTimeout(path);
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
    badge.textContent = fmtTon(valor);
    track.appendChild(badge);

    row.appendChild(track);

    const metaLabel = document.createElement("div");
    metaLabel.className = "hbar-meta-label";
    metaLabel.textContent = fmtTon(meta);
    row.appendChild(metaLabel);

    if (onClick) row.addEventListener("click", onClick);

    // Posiciona o rótulo do valor depois de estar no DOM: fora da barra
    // (após o preenchimento) se couber sem invadir a meta ao lado; senão,
    // para dentro da barra, encostado à direita do preenchimento.
    requestAnimationFrame(() => {
      const trackWidth = track.clientWidth;
      if (!trackWidth) return;
      const fillWidthPx = (pct / 100) * trackWidth;
      const badgeWidth = badge.offsetWidth;
      const GAP = 6;
      if (fillWidthPx + GAP + badgeWidth <= trackWidth) {
        badge.style.left = fillWidthPx + GAP + "px";
      } else {
        badge.style.left = Math.max(GAP, fillWidthPx - GAP - badgeWidth) + "px";
      }
      badge.style.visibility = "visible";
    });

    return row;
  }

  // Ponto na borda do semicírculo (topo do medidor). theta: 180° = esquerda
  // (0%), 90° = topo (50%), 0° = direita (100%).
  function pontoGauge(cx, cy, r, thetaDeg) {
    const rad = (thetaDeg * Math.PI) / 180;
    return { x: cx + r * Math.cos(rad), y: cy - r * Math.sin(rad) };
  }

  // Um arco de no máximo 90° entre dois ângulos (thetaFrom > thetaTo).
  // Nunca desenha um arco de exatamente 180° — esse caso é numericamente
  // instável em SVG (os dois pontos ficam diametralmente opostos e o
  // renderizador pode "confundir" de que lado desenhar o traço).
  function arcoSVG(cx, cy, r, thetaFrom, thetaTo) {
    const p1 = pontoGauge(cx, cy, r, thetaFrom);
    const p2 = pontoGauge(cx, cy, r, thetaTo);
    return `M ${p1.x} ${p1.y} A ${r} ${r} 0 0 1 ${p2.x} ${p2.y}`;
  }

  // Desenha o trecho do medidor entre 0% e `pct` (0..1) como 1 ou 2 arcos de
  // até 90° cada, sempre que o trecho cruzar o topo (50%).
  function criarArcoMedidor(svgNS, cx, cy, r, pct) {
    const grupo = document.createElementNS(svgNS, "g");
    const thetaFim = 180 - Math.min(1, Math.max(0, pct)) * 180;
    if (thetaFim < 90) {
      grupo.appendChild(criarPath(svgNS, arcoSVG(cx, cy, r, 180, 90)));
      grupo.appendChild(criarPath(svgNS, arcoSVG(cx, cy, r, 90, thetaFim)));
    } else {
      grupo.appendChild(criarPath(svgNS, arcoSVG(cx, cy, r, 180, thetaFim)));
    }
    return grupo;
  }

  function criarPath(svgNS, d) {
    const path = document.createElementNS(svgNS, "path");
    path.setAttribute("d", d);
    path.setAttribute("fill", "none");
    return path;
  }

  function renderGauge(container, valor, meta, color) {
    container.innerHTML = "";
    const size = 160;
    const strokeWidth = 24; // traço mais grosso, pra barra parecer maior que o número embaixo
    const pad = strokeWidth / 2 + 2; // >= metade do traço, senão a ponta arredondada corta
    const svgNS = "http://www.w3.org/2000/svg";
    const svg = document.createElementNS(svgNS, "svg");

    const cx = size / 2;
    const r = cx - pad;
    const cy = r + pad;
    const viewBoxHeight = cy + pad;
    svg.setAttribute("viewBox", `0 0 ${size} ${viewBoxHeight}`);
    // largura fixa (não mais 100%) — assim o arco tem um tamanho de caixa
    // conhecido, e o número/rótulos abaixo conseguem se alinhar exatamente
    // com ele em vez de espalhar pela largura toda do painel.
    svg.setAttribute("width", size + "px");
    svg.setAttribute("height", viewBoxHeight + "px"); // altura fixa — tamanho original, não estica com o painel

    const bg = criarArcoMedidor(svgNS, cx, cy, r, 1); // trilho sempre completo, 0% a 100%
    for (const path of bg.children) {
      path.setAttribute("stroke", "#a6a6a6"); // cinza — trilho do gauge sempre visível, mesmo com valor 0
      path.setAttribute("stroke-width", String(strokeWidth));
    }
    svg.appendChild(bg);

    const pct = meta > 0 ? Math.min(1, valor / meta) : 0;
    const fg = criarArcoMedidor(svgNS, cx, cy, r, pct);
    for (const path of fg.children) {
      path.setAttribute("stroke", color);
      path.setAttribute("stroke-width", String(strokeWidth));
      path.setAttribute("stroke-linecap", "round");
    }
    svg.appendChild(fg);

    const arcWrap = document.createElement("div");
    arcWrap.className = "gauge-arc-wrap";
    arcWrap.style.width = size + "px";
    arcWrap.appendChild(svg);
    container.appendChild(arcWrap);

    // Texto sempre em verde-escuro: verde-claro/roxo-destaque têm baixo
    // contraste como texto sobre fundo branco, então servem só para o arco.
    // Fica logo abaixo do arco (não sobreposto): com o texto comprido
    // ("28,789 Ton"), não existe altura onde ele caiba dentro do miolo do
    // arco sem tocar a ponta arredondada do traço — sobrepor sempre "pegava
    // a barra" em algum ponto. Fica só bem colado, sem espaço sobrando.
    const value = document.createElement("div");
    value.className = "gauge-value";
    value.textContent = fmtTon(valor);
    container.appendChild(value);

    // Mesma largura fixa do arco, pra "0,000 Ton" e a meta ficarem perto
    // das pontas do arco em vez de espalhados pela largura do painel.
    const bounds = document.createElement("div");
    bounds.className = "gauge-bounds";
    bounds.style.width = size + "px";
    bounds.innerHTML = `<span>0,000 Ton</span><span>${fmtTon(meta)}</span>`;
    container.appendChild(bounds);
  }

  function renderMetaBar(container, pct) {
    container.innerHTML = "";
    const real = Math.max(0, pct); // número exibido não tem teto — pode passar de 100%
    const largura = Math.min(100, real); // a barra em si não estoura o quadro
    const track = document.createElement("div");
    track.className = "meta-bar-track";

    const atingido = document.createElement("div");
    atingido.className = "meta-bar-atingido";
    atingido.style.width = largura + "%";
    atingido.textContent = real.toFixed(2).replace(".", ",") + "%";
    track.appendChild(atingido);

    const falta = document.createElement("div");
    falta.className = "meta-bar-falta";
    falta.textContent = Math.max(0, 100 - real).toFixed(2).replace(".", ",") + "%";
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

    document.getElementById("producaoMesValor").textContent = fmtTon(payload.producaoMes);
    const [y, m] = payload.data.split("-");
    document.getElementById("mesLabel").textContent = `${MONTH_NAMES[Number(m) - 1]} ${y}`;

    const turmaChart = document.getElementById("producaoTurmaChart");
    turmaChart.innerHTML = "";
    // Com só 1 turma filtrada sobra espaço vazio no resto do quadrado — centraliza
    // a linha única verticalmente e dá um leve destaque nela.
    turmaChart.classList.toggle("hbar-chart--single", !!state.turma);
    const linhasTurma = state.turma
      ? payload.producaoPorTurma.filter((r) => r.turma === state.turma)
      : payload.producaoPorTurma;
    for (const r of linhasTurma) {
      turmaChart.appendChild(
        renderHBarRow({ label: r.turma, valor: r.valor, meta: r.meta, selected: !!state.turma })
      );
    }

    const desaguadorasChart = document.getElementById("desaguadorasChart");
    desaguadorasChart.innerHTML = "";
    for (const r of payload.producaoPorDesaguadora) {
      desaguadorasChart.appendChild(
        renderHBarRow({
          label: `DESAGUADORA ${r.maquina}`,
          valor: r.valor,
          meta: r.meta,
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
      const syncInfo = document.getElementById("syncInfo");
      syncInfo.textContent = "Não consegui atualizar agora (" + err.message + ") — tentando de novo sozinho.";
    }
  }

  // Força o servidor a checar o SharePoint na hora, além de recarregar a
  // tela — se essa parte falhar (rede lenta, SharePoint/D1 fora do ar),
  // ainda assim recarrega com o que já tiver no banco.
  async function forcarSincronizacao(force) {
    const path = force ? "/api/sync?force=1" : "/api/sync";
    await fetchComTimeout(path, { method: "POST" }, 15000).catch((err) => console.warn("Forçar sync falhou:", err.message));
    await load();
  }

  const btnRefresh = document.getElementById("btnRefresh");
  btnRefresh.addEventListener("click", async () => {
    btnRefresh.disabled = true;
    const textoOriginal = btnRefresh.textContent;
    btnRefresh.textContent = "↻ Atualizando…";
    try {
      await forcarSincronizacao(true);
    } finally {
      btnRefresh.disabled = false;
      btnRefresh.textContent = textoOriginal;
    }
  });

  load();
  setInterval(load, REFRESH_MS);
  // Redundância: mesmo com o cron do servidor rodando de 1 em 1 minuto,
  // o navegador (que fica aberto o dia todo, numa TV) também checa uma
  // sincronização de tempos em tempos por conta própria — não depende só
  // do cron do lado do servidor pra insistir quando algo falha. Sem force
  // (ver comentário do botão Atualizar): barato quando nada mudou.
  setInterval(() => forcarSincronizacao(false), FORCE_SYNC_MS);
})();
