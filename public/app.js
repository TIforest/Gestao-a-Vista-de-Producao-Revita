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

  // Cores do BI original (mesmas do styles.css) — usadas só onde o SVG
  // precisa delas em atributo.
  const COR = {
    verde: "#18c13b",
    verdeTexto: "#2e9e3a",
    azulArco: "#1f5fe6",
    azulTexto: "#2b33d6",
    trilho: "#e6e6e6",
  };

  // BI mostra sempre 3 casas: "33,222 Ton", "40,000 Ton".
  function fmtTonNum(valor) {
    const ton = (valor || 0) / 1000; // peso_seco vem em kg — dividir por 1000 já dá toneladas
    return ton.toLocaleString("pt-BR", { minimumFractionDigits: 3, maximumFractionDigits: 3 });
  }
  function fmtTon(valor) {
    return fmtTonNum(valor) + " Ton";
  }

  function fmtKg(valor) {
    return (valor || 0).toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  function fmtPct(p) {
    return p.toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + "%";
  }

  function fmtHora(isoDataHora) {
    const t = isoDataHora?.split("T")[1];
    return t ? t.slice(0, 8) : "—";
  }

  const MONTH_NAMES = [
    "JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO",
    "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO",
  ];

  function el(tag, className, text) {
    const node = document.createElement(tag);
    if (className) node.className = className;
    if (text !== undefined) node.textContent = text;
    return node;
  }

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

  // ---------------------------------------------------------------------
  // Barras em par (como no BI): produção em verde em cima, meta em azul
  // embaixo, as duas na mesma escala dentro do gráfico. `escala` é o maior
  // valor do gráfico (produção ou meta) — essa barra ocupa ~64% da largura,
  // o resto fica pro rótulo à direita.
  // ---------------------------------------------------------------------
  const LARGURA_MAX_PCT = 64;

  function escalaDoGrafico(linhas) {
    let max = 0;
    for (const l of linhas) max = Math.max(max, l.valor || 0, l.meta || 0);
    return max;
  }

  function linhaDeBarra(classe, valor, escala, italico, inside) {
    const line = el("div", "pbar-line");
    const fill = el("div", "pbar-fill " + classe);
    const pct = escala > 0 ? Math.max(0, Math.min(100, ((valor || 0) / escala) * 100)) : 0;
    fill.style.width = (pct * LARGURA_MAX_PCT) / 100 + "%";
    const text = el("span", "pbar-text" + (italico ? " italic" : ""), fmtTon(valor));
    line.appendChild(fill);
    // No modo "inside" (gráfico de meta por hora) o rótulo vai dentro da
    // barra, em branco; nos outros fica fora, à direita do preenchimento.
    if (inside) fill.appendChild(text);
    else line.appendChild(text);
    return line;
  }

  function renderPairBar({ label, valor, meta, escala, italico, inside, onClick, selected }) {
    const row = el("div", "pbar-row" + (onClick ? " clickable" : "") + (selected ? " selected" : ""));
    row.appendChild(el("div", "pbar-label", label));

    const bars = el("div", "pbar-bars");
    bars.appendChild(linhaDeBarra("pbar-valor", valor, escala, italico, inside));
    bars.appendChild(linhaDeBarra("pbar-meta", meta, escala, italico, inside));
    row.appendChild(bars);

    if (onClick) row.addEventListener("click", onClick);
    return row;
  }

  // ---------------------------------------------------------------------
  // Gauge (semicírculo). Ponto na borda: theta 180° = esquerda (0%),
  // 90° = topo (50%), 0° = direita (100%).
  // ---------------------------------------------------------------------
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

  // Como no BI: arco grande, traço grosso, sem ponta arredondada, número
  // colorido dentro do arco (entre as duas pontas) e limites "0,000" /
  // "40,000 Ton" em itálico logo abaixo, na mesma cor.
  function renderGauge(container, valor, meta, corArco, corTexto) {
    container.innerHTML = "";
    const size = 210;
    const strokeWidth = 30;
    const pad = strokeWidth / 2 + 2;
    const svgNS = "http://www.w3.org/2000/svg";
    const svg = document.createElementNS(svgNS, "svg");

    const cx = size / 2;
    const r = cx - pad;
    const cy = r + pad;
    const viewBoxHeight = cy + pad;
    svg.setAttribute("viewBox", `0 0 ${size} ${viewBoxHeight}`);
    svg.setAttribute("width", size + "px");
    svg.setAttribute("height", viewBoxHeight + "px");

    const bg = criarArcoMedidor(svgNS, cx, cy, r, 1); // trilho sempre completo, 0% a 100%
    for (const path of bg.children) {
      path.setAttribute("stroke", COR.trilho);
      path.setAttribute("stroke-width", String(strokeWidth));
    }
    svg.appendChild(bg);

    const pct = meta > 0 ? Math.min(1, valor / meta) : 0;
    const fg = criarArcoMedidor(svgNS, cx, cy, r, pct);
    for (const path of fg.children) {
      path.setAttribute("stroke", corArco);
      path.setAttribute("stroke-width", String(strokeWidth));
    }
    svg.appendChild(fg);

    const arcWrap = el("div", "gauge-arc-wrap");
    arcWrap.style.width = size + "px";
    arcWrap.style.height = viewBoxHeight + "px";
    arcWrap.appendChild(svg);

    const value = el("div", "gauge-value", fmtTonNum(valor));
    value.style.color = corTexto;
    arcWrap.appendChild(value);
    container.appendChild(arcWrap);

    const bounds = el("div", "gauge-bounds");
    bounds.style.width = size + "px";
    bounds.style.color = corTexto;
    bounds.innerHTML = `<span>0,000</span><span>${fmtTon(meta)}</span>`;
    container.appendChild(bounds);
  }

  // ---------------------------------------------------------------------
  // % da meta atingida: uma barra por turma (verde = atingido, vermelho =
  // falta), com as linhas pontilhadas de 0% / 50% / 100% atrás, como no BI.
  // ---------------------------------------------------------------------
  function renderMetaChart(container, linhas) {
    container.innerHTML = "";
    const rows = el("div", "meta-rows");
    for (const pos of [0, 50, 100]) {
      const v = el("div", "meta-vline");
      v.style.left = pos + "%";
      rows.appendChild(v);
    }
    for (const l of linhas) {
      const pct = l.meta > 0 ? (l.valor / l.meta) * 100 : 0;
      const real = Math.max(0, pct); // número exibido não tem teto — pode passar de 100%
      const largura = Math.min(100, real); // a barra em si não estoura o quadro

      const row = el("div", "meta-row");
      row.appendChild(el("div", "meta-row-label", l.turma));
      const track = el("div", "meta-track");
      const atingido = el("div", "meta-atingido", fmtPct(real));
      atingido.style.width = largura + "%";
      track.appendChild(atingido);
      const falta = el("div", "meta-falta", fmtPct(Math.max(0, 100 - real)));
      if (largura >= 100) falta.style.display = "none";
      track.appendChild(falta);
      row.appendChild(track);
      rows.appendChild(row);
    }
    container.appendChild(rows);

    const ticks = el("div", "meta-ticks");
    ticks.innerHTML = "<span>0%</span><span>50%</span><span>100%</span>";
    container.appendChild(ticks);
  }

  // Lista só as turmas que têm apontamento no dia (o servidor já manda
  // assim — igual ao segmentador do BI). Sem turma selecionada, todas as
  // caixas ficam verdes ("todas"); clicando numa, só ela fica verde;
  // clicando de novo, volta pra todas.
  function renderTurmas(payload) {
    const list = document.getElementById("turmasList");
    list.innerHTML = "";
    for (const t of payload.turmasDisponiveis) {
      const ativa = state.turma === null || state.turma === t;
      const box = el("div", "turma-box" + (ativa ? " active" : ""), t);
      box.title = state.turma === t ? "Clique pra voltar a mostrar todas as turmas" : "Mostrar só a turma " + t;
      box.addEventListener("click", () => { state.turma = state.turma === t ? null : t; load(); });
      list.appendChild(box);
    }
  }

  function render(payload) {
    renderTurmas(payload);

    document.getElementById("producaoMesValor").textContent = fmtTon(payload.producaoMes);
    const [y, m] = payload.data.split("-");
    const mesLabel = document.getElementById("mesLabel");
    mesLabel.textContent = MONTH_NAMES[Number(m) - 1];
    mesLabel.title = `${MONTH_NAMES[Number(m) - 1]} ${y}`;

    // Produção x turma: só as turmas ativas (ou a selecionada), cada uma
    // contra a meta fixa do turno — tudo na mesma escala.
    const turmaChart = document.getElementById("producaoTurmaChart");
    turmaChart.innerHTML = "";
    const linhasTurma = state.turma
      ? payload.producaoPorTurma.filter((r) => r.turma === state.turma)
      : payload.producaoPorTurma;
    const escalaTurma = escalaDoGrafico(linhasTurma);
    for (const r of linhasTurma) {
      turmaChart.appendChild(
        renderPairBar({ label: r.turma, valor: r.valor, meta: r.meta, escala: escalaTurma, italico: true, selected: !!state.turma })
      );
    }

    const desaguadorasChart = document.getElementById("desaguadorasChart");
    desaguadorasChart.innerHTML = "";
    const escalaDesag = escalaDoGrafico(payload.producaoPorDesaguadora);
    for (const r of payload.producaoPorDesaguadora) {
      desaguadorasChart.appendChild(
        renderPairBar({
          label: `DESAGUADORA ${r.maquina}`,
          valor: r.valor,
          meta: r.meta,
          escala: escalaDesag,
          selected: state.maquina === r.maquina,
          onClick: () => { state.maquina = state.maquina === r.maquina ? null : r.maquina; load(); },
        })
      );
    }

    const horaChart = document.getElementById("horaChart");
    horaChart.innerHTML = "";
    horaChart.appendChild(
      renderPairBar({
        label: state.turma || "TODAS",
        valor: payload.producaoMediaHora,
        meta: payload.metaHora,
        escala: escalaDoGrafico([{ valor: payload.producaoMediaHora, meta: payload.metaHora }]),
        inside: true,
      })
    );

    renderGauge(document.getElementById("gaugeTurno"), payload.producaoTurno, payload.metaTurno, COR.verde, COR.verdeTexto);
    renderGauge(document.getElementById("gaugeDia"), payload.producaoDia, payload.metaDia, COR.azulArco, COR.azulTexto);

    // % da meta: uma barra por turma ativa (ou só a selecionada), cada uma
    // contra a meta do turno — como o eixo "TURMA" do BI.
    renderMetaChart(document.getElementById("metaAtingidaChart"), linhasTurma);

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
      // A turma selecionada sumiu da lista (virou o dia, ou ela não tem mais
      // apontamento): volta pra "todas" em vez de ficar numa tela zerada.
      if (state.turma && !payload.turmasDisponiveis.includes(state.turma)) {
        state.turma = null;
        return load();
      }
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
