/* app.js — CREA-BA: Instituições, Cursos, Títulos e Atribuições
 * Le dados/dados.json (gerado por scripts/gerar_json.py) e monta:
 *  - indicadores
 *  - busca geral + filtros combinados
 *  - tabela paginada/ordenavel/exportavel
 *  - mapa Leaflet por municipio, sincronizado com a tabela
 *  - modal de detalhe
 * Sem framework, sem dependencias alem do Leaflet (CDN).
 */
(function () {
  "use strict";

  const PAGE_SIZE = 25;
  const CAMPOS_TABELA = [
    { chave: "nome_ies", rotulo: "Instituição" },
    { chave: "municipio", rotulo: "Município" },
    { chave: "nome_campus", rotulo: "Campus" },
    { chave: "nome_curso", rotulo: "Curso" },
    { chave: "titulo_profissional", rotulo: "Título profissional" },
    { chave: "modalidade_oferta", rotulo: "Modalidade" },
    { chave: "categoria_academica", rotulo: "Categoria" },
    { chave: "status_validacao_cadastral", rotulo: "Situação cadastral" },
    { chave: "status_validacao_legal", rotulo: "Validação normativa" },
  ];

  const estado = {
    todos: [],
    filtrados: [],
    pagina: 1,
    ordenarPor: "nome_ies",
    ordemAsc: true,
    filtros: {},
    somentePublicaveis: true,
    busca: "",
    mapa: null,
    marcadores: new Map(),
    municipioSelecionado: "",
  };

  function normalizaTexto(s) {
    return (s || "")
      .toString()
      .normalize("NFD")
      .replace(/[̀-ͯ]/g, "")
      .toLowerCase();
  }

  function statusVisual(registro) {
    const legal = registro.status_validacao_legal;
    const cadastral = registro.status_validacao_cadastral;
    if (cadastral === "PUBLICÁVEL" && legal && legal !== "PENDENTE DE VALIDAÇÃO") return "validado";
    if (cadastral === "PUBLICÁVEL" || (legal && legal !== "PENDENTE DE VALIDAÇÃO")) return "parcial";
    return "pendente";
  }

  function rotuloStatus(nivel) {
    return { validado: "Validado", parcial: "Parcial", pendente: "Pendente" }[nivel];
  }

  // ---------------------------------------------------------------------
  // Carregamento
  // ---------------------------------------------------------------------
  async function carregar() {
    const resp = await fetch("../dados/dados.json");
    if (!resp.ok) throw new Error("Não foi possível carregar dados/dados.json (" + resp.status + ")");
    const payload = await resp.json();
    estado.todos = payload.registros || [];
    estado.meta = payload.meta || {};
    preencherIndicadores(estado.meta);
    preencherFiltros(estado.todos);
    aplicarFiltros();
    inicializarMapa();
    atualizarMapa();
  }

  function preencherIndicadores(meta) {
    document.getElementById("meta-atualizacao").textContent = meta.gerado_em || "—";
    document.getElementById("ind-instituicoes").textContent = meta.total_instituicoes ?? "—";
    document.getElementById("ind-campi").textContent = meta.total_campi ?? "—";
    document.getElementById("ind-municipios").textContent = meta.total_municipios ?? "—";
    document.getElementById("ind-cursos").textContent = meta.total_cursos ?? "—";
    document.getElementById("ind-titulos").textContent = meta.total_titulos ?? "—";
    document.getElementById("ind-base-legal").textContent = meta.total_cursos_com_base_legal_validada ?? "—";
    document.getElementById("ind-pendencias").textContent = meta.total_pendencias ?? "—";
    document.getElementById("ind-completude").textContent = (meta.percentual_completude_cadastral ?? "—") + "%";
  }

  function opcoesUnicas(lista, chave) {
    const s = new Set();
    lista.forEach((r) => { if (r[chave]) s.add(r[chave]); });
    return Array.from(s).sort((a, b) => a.localeCompare(b, "pt-BR"));
  }

  function preencherSelect(id, valores) {
    const sel = document.getElementById(id);
    valores.forEach((v) => {
      const opt = document.createElement("option");
      opt.value = v;
      opt.textContent = v;
      sel.appendChild(opt);
    });
  }

  function preencherFiltros(dados) {
    preencherSelect("f-municipio", opcoesUnicas(dados, "municipio"));
    preencherSelect("f-instituicao", opcoesUnicas(dados, "nome_ies"));
    preencherSelect("f-curso", opcoesUnicas(dados, "nome_curso"));
    preencherSelect("f-titulo", opcoesUnicas(dados, "titulo_profissional"));
    preencherSelect("f-modalidade", opcoesUnicas(dados, "modalidade_oferta"));
    preencherSelect("f-grau", opcoesUnicas(dados, "grau_academico"));
    preencherSelect("f-categoria", opcoesUnicas(dados, "categoria_academica"));
    preencherSelect("f-natureza", opcoesUnicas(dados, "natureza_administrativa"));
    preencherSelect("f-situacao-curso", opcoesUnicas(dados, "situacao_curso"));
    preencherSelect("f-situacao-cadastral", opcoesUnicas(dados, "status_validacao_cadastral"));
    preencherSelect("f-status-legal", opcoesUnicas(dados, "status_validacao_legal"));
  }

  // ---------------------------------------------------------------------
  // Filtros + busca combinada
  // ---------------------------------------------------------------------
  const CAMPOS_BUSCA = [
    "nome_ies", "sigla_ies", "nome_campus", "municipio", "nome_curso",
    "titulo_profissional", "codigo_titulo", "norma", "artigo", "atribuicao_resumida",
  ];

  function aplicarFiltros() {
    const buscaNorm = normalizaTexto(estado.busca);
    const f = estado.filtros;
    estado.filtrados = estado.todos.filter((r) => {
      if (estado.somentePublicaveis && r.publicar_dashboard !== "SIM") return false;
      if (estado.municipioSelecionado && r.municipio !== estado.municipioSelecionado) return false;
      for (const [campo, valor] of Object.entries(f)) {
        if (valor && r[campo] !== valor) return false;
      }
      if (buscaNorm) {
        const alvo = CAMPOS_BUSCA.map((c) => normalizaTexto(r[c])).join(" | ");
        if (!alvo.includes(buscaNorm)) return false;
      }
      return true;
    });
    ordenar();
    estado.pagina = 1;
    renderTabela();
    renderContagem();
    atualizarMapa();
  }

  function ordenar() {
    const chave = estado.ordenarPor;
    const dir = estado.ordemAsc ? 1 : -1;
    estado.filtrados.sort((a, b) => {
      const va = (a[chave] ?? "").toString();
      const vb = (b[chave] ?? "").toString();
      return va.localeCompare(vb, "pt-BR") * dir;
    });
  }

  function renderContagem() {
    document.getElementById("contagem-resultados").textContent =
      estado.filtrados.length + " registro(s) encontrado(s)";
    const aviso = document.getElementById("aviso-sem-coordenada");
    const semCoord = estado.filtrados.some((r) => r.publicar_dashboard === "SIM" && !r.latitude && r.municipio === "");
    aviso.hidden = !semCoord;
  }

  // ---------------------------------------------------------------------
  // Tabela
  // ---------------------------------------------------------------------
  function montarCabecalho() {
    const tr = document.getElementById("linha-cabecalho");
    tr.innerHTML = "";
    CAMPOS_TABELA.forEach((c) => {
      const th = document.createElement("th");
      th.textContent = c.rotulo;
      th.tabIndex = 0;
      th.setAttribute("role", "button");
      th.setAttribute("aria-sort", estado.ordenarPor === c.chave ? (estado.ordemAsc ? "ascending" : "descending") : "none");
      th.addEventListener("click", () => alternarOrdenacao(c.chave));
      th.addEventListener("keydown", (e) => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); alternarOrdenacao(c.chave); } });
      tr.appendChild(th);
    });
  }

  function alternarOrdenacao(chave) {
    if (estado.ordenarPor === chave) {
      estado.ordemAsc = !estado.ordemAsc;
    } else {
      estado.ordenarPor = chave;
      estado.ordemAsc = true;
    }
    ordenar();
    renderTabela();
  }

  function renderTabela() {
    const corpo = document.getElementById("corpo-tabela");
    corpo.innerHTML = "";
    const totalPaginas = Math.max(1, Math.ceil(estado.filtrados.length / PAGE_SIZE));
    if (estado.pagina > totalPaginas) estado.pagina = totalPaginas;
    const inicio = (estado.pagina - 1) * PAGE_SIZE;
    const pagina = estado.filtrados.slice(inicio, inicio + PAGE_SIZE);

    pagina.forEach((r) => {
      const tr = document.createElement("tr");
      tr.tabIndex = 0;
      tr.setAttribute("role", "button");
      tr.setAttribute("aria-label", "Ver detalhes de " + (r.nome_ies || "registro"));
      CAMPOS_TABELA.forEach((c) => {
        const td = document.createElement("td");
        if (c.chave === "status_validacao_cadastral" || c.chave === "status_validacao_legal") {
          const nivel = statusVisual(r);
          const selo = document.createElement("span");
          selo.className = "selo " + nivel;
          selo.textContent = (c.chave === "status_validacao_cadastral" ? r[c.chave] : rotuloStatus(nivel)) || "—";
          td.appendChild(selo);
        } else {
          td.textContent = r[c.chave] || "—";
        }
        tr.appendChild(td);
      });
      tr.addEventListener("click", () => abrirModal(r));
      tr.addEventListener("keydown", (e) => { if (e.key === "Enter") abrirModal(r); });
      corpo.appendChild(tr);
    });

    document.getElementById("pag-atual").textContent = "Página " + estado.pagina + " de " + totalPaginas;
    document.getElementById("btn-pag-anterior").disabled = estado.pagina <= 1;
    document.getElementById("btn-pag-proxima").disabled = estado.pagina >= totalPaginas;
    montarCabecalho();
  }

  // ---------------------------------------------------------------------
  // Modal de detalhe
  // ---------------------------------------------------------------------
  const CAMPOS_MODAL = [
    ["nome_ies", "Instituição"], ["sigla_ies", "Sigla"], ["cnpj", "CNPJ"], ["codigo_mec_ies", "Código MEC"],
    ["categoria_academica", "Categoria acadêmica"], ["natureza_administrativa", "Natureza administrativa"],
    ["nome_campus", "Campus"], ["municipio", "Município"], ["uf", "UF"],
    ["nome_curso", "Curso"], ["grau_academico", "Grau acadêmico"], ["modalidade_oferta", "Modalidade"],
    ["situacao_curso", "Situação do curso"],
    ["codigo_titulo", "Código do título"], ["titulo_profissional", "Título profissional"],
    ["norma", "Norma"], ["artigo", "Artigo"], ["atribuicao_resumida", "Atribuição (resumo)"],
    ["url_fonte_legal", "Fonte legal (link)"],
    ["status_validacao_cadastral", "Situação cadastral"], ["status_validacao_legal", "Status de validação normativa"],
    ["data_atualizacao", "Última atualização"],
  ];

  function abrirModal(r) {
    document.getElementById("modal-titulo").textContent = r.nome_ies || "Detalhe do registro";
    const dl = document.getElementById("modal-corpo");
    dl.innerHTML = "";
    CAMPOS_MODAL.forEach(([chave, rotulo]) => {
      const valor = r[chave];
      if (!valor) return;
      const dt = document.createElement("dt");
      dt.textContent = rotulo;
      const dd = document.createElement("dd");
      if (chave === "url_fonte_legal") {
        const a = document.createElement("a");
        a.href = valor; a.textContent = valor; a.target = "_blank"; a.rel = "noopener";
        dd.appendChild(a);
      } else {
        dd.textContent = valor;
      }
      dl.appendChild(dt);
      dl.appendChild(dd);
    });
    const overlay = document.getElementById("modal-overlay");
    overlay.hidden = false;
    document.getElementById("btn-fechar-modal").focus();
  }

  function fecharModal() {
    document.getElementById("modal-overlay").hidden = true;
  }

  // ---------------------------------------------------------------------
  // Exportação CSV
  // ---------------------------------------------------------------------
  function exportarCSV() {
    const colunas = CAMPOS_TABELA.map((c) => c.chave).concat(["codigo_titulo", "norma", "artigo", "url_fonte_legal", "cnpj", "codigo_mec_ies"]);
    const linhas = [colunas.join(",")];
    estado.filtrados.forEach((r) => {
      const linha = colunas.map((c) => {
        const v = (r[c] ?? "").toString().replace(/"/g, '""');
        return /[",\n]/.test(v) ? `"${v}"` : v;
      });
      linhas.push(linha.join(","));
    });
    const blob = new Blob(["﻿" + linhas.join("\n")], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "crea_ba_instituicoes_cursos.csv";
    a.click();
    URL.revokeObjectURL(url);
  }

  // ---------------------------------------------------------------------
  // Mapa (Leaflet) — 1 marcador por municipio, agrupado por instituicao
  // ---------------------------------------------------------------------
  const BAHIA_BOUNDS = L.latLngBounds([-18.35, -46.7], [-8.5, -37.3]);

  function inicializarMapa() {
    estado.mapa = L.map("mapa", { scrollWheelZoom: false });
    L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
      attribution: "&copy; colaboradores do OpenStreetMap",
      maxZoom: 18,
    }).addTo(estado.mapa);
    estado.mapa.fitBounds(BAHIA_BOUNDS);
    // Garante o dimensionamento correto do mapa apos o layout CSS assentar
    // (containers dentro de grid podem ter largura 0 no instante da criacao do Leaflet).
    setTimeout(() => {
      estado.mapa.invalidateSize();
      estado.mapa.fitBounds(BAHIA_BOUNDS);
    }, 150);
    window.addEventListener("resize", () => estado.mapa.invalidateSize());
  }

  function corPorNivel(nivel) {
    return { validado: "#1a7a4c", parcial: "#8a6300", pendente: "#a3241f" }[nivel];
  }

  function agruparPorMunicipio(registros) {
    const mapa = new Map();
    registros.forEach((r) => {
      if (!r.municipio || r.latitude == null || r.longitude == null) return;
      if (!mapa.has(r.municipio)) {
        mapa.set(r.municipio, { municipio: r.municipio, lat: r.latitude, lon: r.longitude, instituicoes: new Map() });
      }
      const grupo = mapa.get(r.municipio);
      if (!grupo.instituicoes.has(r.id_ies)) {
        grupo.instituicoes.set(r.id_ies, { nome: r.nome_ies, sigla: r.sigla_ies, cursos: [] });
      }
      if (r.id_curso) {
        grupo.instituicoes.get(r.id_ies).cursos.push({
          curso: r.nome_curso, titulo: r.titulo_profissional, modalidade: r.modalidade_oferta,
          status: r.status_validacao_cadastral,
        });
      }
    });
    return mapa;
  }

  function popupHtml(grupo) {
    const nInst = grupo.instituicoes.size;
    let nCursos = 0;
    grupo.instituicoes.forEach((i) => (nCursos += i.cursos.length));
    let html = `<div class="popup-municipio"><h3>${grupo.municipio.toUpperCase()}</h3>`;
    html += `<div class="resumo">${nInst} instituição(ões) · ${nCursos} curso(s) cadastrado(s)</div>`;
    grupo.instituicoes.forEach((inst) => {
      html += `<details class="popup-inst"><summary>${inst.sigla ? inst.sigla + " — " : ""}${inst.nome}</summary>`;
      if (inst.cursos.length === 0) {
        html += `<div class="curso-item">Sem curso cadastrado nesta base (apenas instituição/campus).</div>`;
      } else {
        inst.cursos.forEach((c) => {
          html += `<div class="curso-item">• ${c.curso || "—"}<br>Título: ${c.titulo || "—"}<br>Modalidade: ${c.modalidade || "—"} · ${c.status || "—"}</div>`;
        });
      }
      html += `</details>`;
    });
    html += `</div>`;
    return html;
  }

  function atualizarMapa() {
    if (!estado.mapa) return;
    estado.marcadores.forEach((m) => estado.mapa.removeLayer(m));
    estado.marcadores.clear();

    const grupos = agruparPorMunicipio(estado.filtrados);
    grupos.forEach((grupo, nomeMunicipio) => {
      let melhorNivel = "pendente";
      let temValidado = false, temParcial = false;
      estado.filtrados
        .filter((r) => r.municipio === nomeMunicipio)
        .forEach((r) => {
          const n = statusVisual(r);
          if (n === "validado") temValidado = true;
          if (n === "parcial") temParcial = true;
        });
      melhorNivel = temValidado ? "validado" : temParcial ? "parcial" : "pendente";

      const marker = L.circleMarker([grupo.lat, grupo.lon], {
        radius: 7 + Math.min(6, grupo.instituicoes.size),
        color: "#1f2328",
        weight: 1,
        fillColor: corPorNivel(melhorNivel),
        fillOpacity: 0.85,
      });
      marker.bindPopup(popupHtml(grupo));
      marker.on("click", () => {
        estado.municipioSelecionado = estado.municipioSelecionado === nomeMunicipio ? "" : nomeMunicipio;
        document.getElementById("f-municipio").value = estado.municipioSelecionado;
        aplicarFiltros();
      });
      marker.addTo(estado.mapa);
      estado.marcadores.set(nomeMunicipio, marker);
    });

    if (estado.municipioSelecionado && estado.marcadores.has(estado.municipioSelecionado)) {
      estado.marcadores.get(estado.municipioSelecionado).openPopup();
    }
  }

  // ---------------------------------------------------------------------
  // Eventos de UI
  // ---------------------------------------------------------------------
  function ligarEventos() {
    document.getElementById("busca").addEventListener("input", (e) => {
      estado.busca = e.target.value;
      aplicarFiltros();
    });
    document.getElementById("btn-limpar-busca").addEventListener("click", () => {
      document.getElementById("busca").value = "";
      estado.busca = "";
      aplicarFiltros();
    });

    const mapaFiltros = {
      "f-municipio": "municipio", "f-instituicao": "nome_ies", "f-curso": "nome_curso",
      "f-titulo": "titulo_profissional", "f-modalidade": "modalidade_oferta", "f-grau": "grau_academico",
      "f-categoria": "categoria_academica", "f-natureza": "natureza_administrativa",
      "f-situacao-curso": "situacao_curso", "f-situacao-cadastral": "status_validacao_cadastral",
      "f-status-legal": "status_validacao_legal",
    };
    Object.entries(mapaFiltros).forEach(([id, campo]) => {
      document.getElementById(id).addEventListener("change", (e) => {
        estado.filtros[campo] = e.target.value;
        if (campo === "municipio") estado.municipioSelecionado = e.target.value;
        aplicarFiltros();
      });
    });

    document.getElementById("f-somente-publicaveis").addEventListener("change", (e) => {
      estado.somentePublicaveis = e.target.checked;
      aplicarFiltros();
    });

    document.getElementById("btn-limpar-filtros").addEventListener("click", () => {
      estado.filtros = {};
      estado.busca = "";
      estado.municipioSelecionado = "";
      document.getElementById("busca").value = "";
      Object.keys(mapaFiltros).forEach((id) => (document.getElementById(id).value = ""));
      document.getElementById("f-somente-publicaveis").checked = true;
      estado.somentePublicaveis = true;
      aplicarFiltros();
    });

    document.getElementById("btn-exportar-csv").addEventListener("click", exportarCSV);

    document.getElementById("btn-pag-anterior").addEventListener("click", () => {
      if (estado.pagina > 1) { estado.pagina--; renderTabela(); }
    });
    document.getElementById("btn-pag-proxima").addEventListener("click", () => {
      estado.pagina++; renderTabela();
    });

    document.getElementById("btn-fechar-modal").addEventListener("click", fecharModal);
    document.getElementById("modal-overlay").addEventListener("click", (e) => {
      if (e.target.id === "modal-overlay") fecharModal();
    });
    document.addEventListener("keydown", (e) => {
      if (e.key === "Escape") fecharModal();
    });
  }

  ligarEventos();
  carregar().catch((err) => {
    console.error(err);
    document.getElementById("contagem-resultados").textContent = "Erro ao carregar dados: " + err.message;
  });
})();
