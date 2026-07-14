// ── Unidades ─────────────────────────────────────────────
const UNIDADES = [
  { cod: 'SP',     nome: 'São Paulo',      arquivo: 'dados/base/sao_paulo.xlsx' },
  { cod: 'GOIAS',  nome: 'Goiás',          arquivo: 'dados/base/goias.xlsx' },
  { cod: 'RJ',     nome: 'Rio de Janeiro', arquivo: 'dados/base/rio_de_janeiro.xlsx' },
  { cod: 'Santos', nome: 'Santos',         arquivo: 'dados/base/santos.xlsx' },
];

// ── Estado ───────────────────────────────────────────────
const DADOS_UNIDADE      = {};   // cod → array de rows
const PROMESSAS          = {};   // cod → Promise
const DATAS_ATUALIZACAO  = {};   // cod → string data formatada

const estado = {
  unidade:      null,
  unidadeRows:  [],
  departamento: null,
  depRows:      [],
  abaAtiva:     'Em Atraso',
};

const filtrosRel = { coordenador: '', responsavel: '', grupo: '', busca: '' };

// ── Filtros da tabela de Tarefas (estilo Excel) ────────────
function clienteLabel(r) {
  const cod = String(r.CodCliente || '').trim();
  const raz = String(r.RazaoSocial || '').trim();
  return cod && raz ? `${cod} - ${raz}` : (cod || raz);
}

// Usuários "genéricos" de baixa: quando UsuarioBaixa é um destes, a baixa não
// identifica quem realmente executou a tarefa (ex-funcionário, conta genérica),
// então usamos o UsuarioResponsavel original como se ele tivesse baixado.
const USUARIOS_BAIXA_GENERICOS = ['Juan Rodrigues', 'ex-funcionario', 'Nicole Oliveira', 'Gabriel Amaral'];

function usuarioBaixaEfetivo(r) {
  const baixa = String(r.UsuarioBaixa || '').trim();
  if (USUARIOS_BAIXA_GENERICOS.includes(baixa)) return String(r.UsuarioResponsavel || '').trim();
  return baixa;
}

const TODAS_CHAVES_TAB = ['cliente', 'grupo', 'tarefa', 'coordenador', 'responsavel', 'baixadopor', 'competencia', 'vencimento', 'previsao', 'databaixa', 'prioridade'];

// Colunas da tabela de Tarefas: na aba 'Baixado', 'Previsão' vira 'Data de baixa'
// (lendo DataBaixa) e 'Responsável' passa a ler UsuarioBaixa (quem baixou) em vez
// de UsuarioResponsavel — nas demais abas (Em Atraso/Em Aberto) a tarefa ainda não
// foi baixada, então previsão/responsável originais continuam fazendo sentido.
function colunasTab() {
  const colPrevisaoOuBaixa = estado.abaAtiva === 'Baixado'
    ? { key: 'databaixa', label: 'Data de baixa', valor: r => fmtData(r.DataBaixa), ordem: r => r.DataBaixa instanceof Date ? r.DataBaixa.getTime() : 0 }
    : { key: 'previsao',  label: 'Previsão',      valor: r => fmtData(r.DataPrevisaoConclusao), ordem: r => r.DataPrevisaoConclusao instanceof Date ? r.DataPrevisaoConclusao.getTime() : 0 };

  const colResponsavel = estado.abaAtiva === 'Baixado'
    ? { key: 'baixadopor', label: 'Responsável', valor: r => usuarioBaixaEfetivo(r) }
    : { key: 'responsavel', label: 'Responsável', valor: r => String(r.UsuarioResponsavel || '') };

  return [
    { key: 'cliente',     label: 'Cliente',     valor: r => clienteLabel(r) },
    { key: 'grupo',       label: 'Grupo',       valor: r => String(r.Grupo || '') },
    { key: 'tarefa',      label: 'Tarefa',      valor: r => String(r.Titulo || '') },
    { key: 'coordenador', label: 'Coordenação', valor: r => String(r.Coordenador || '') },
    colResponsavel,
    { key: 'competencia', label: 'Competência', valor: r => fmtCompetencia(r.Competencia), ordem: r => r.Competencia instanceof Date ? r.Competencia.getTime() : 0 },
    { key: 'vencimento',  label: 'Vencimento',  valor: r => fmtData(r.DataVencimento), ordem: r => r.DataVencimento instanceof Date ? r.DataVencimento.getTime() : 0 },
    colPrevisaoOuBaixa,
    { key: 'prioridade',  label: 'Prioridade',  valor: r => String(r.Prioridade || '') },
  ];
}

const filtrosTab = {
  busca: '',
  excluidos: Object.fromEntries(TODAS_CHAVES_TAB.map(k => [k, new Set()])),
};
let filtroColunaAberta = null;
let draftExcluidos = null;

let buscaTimerRel = null;
let buscaTimerTab = null;

// ── Utilitários ───────────────────────────────────────────
function esc(s) {
  if (s == null) return '';
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function fmtData(v) {
  if (!v) return '';
  if (v instanceof Date) return v.toLocaleDateString('pt-BR');
  return String(v);
}

function fmtCompetencia(v) {
  if (!v) return '';
  if (v instanceof Date) return v.toLocaleDateString('pt-BR', { month: '2-digit', year: 'numeric' });
  return String(v);
}

function contarStatus(rows) {
  let emAtraso = 0, emAberto = 0, baixadas = 0;
  for (const r of rows) {
    if      (r.Status === 'Em Atraso') emAtraso++;
    else if (r.Status === 'Em Aberto') emAberto++;
    else if (r.Status === 'Baixado')   baixadas++;
  }
  return { total: rows.length, emAtraso, emAberto, baixadas };
}

// ── Carregamento ──────────────────────────────────────────
function _carregarArquivo(cod, arquivo) {
  PROMESSAS[cod] = fetch(arquivo)
    .then(r => {
      if (!r.ok) throw new Error(`HTTP ${r.status}`);
      const lastMod = r.headers.get('Last-Modified');
      if (lastMod) {
        const d = new Date(lastMod);
        const hora = d.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
        DATAS_ATUALIZACAO[cod] = d.toLocaleDateString('pt-BR') + ' ' + hora;
      }
      return r.arrayBuffer();
    })
    .then(buf => {
      const wb = XLSX.read(buf, { type: 'array', cellDates: true });
      const ws = wb.Sheets['Pendencias'];
      DADOS_UNIDADE[cod] = XLSX.utils.sheet_to_json(ws, { raw: true });
    })
    .catch(err => {
      DADOS_UNIDADE[cod] = null;
      console.error(`Erro ao carregar ${arquivo}:`, err);
    });
}

async function iniciarCarregamento() {
  UNIDADES.forEach(u => _carregarArquivo(u.cod, u.arquivo));
  await Promise.all(UNIDADES.map(u => PROMESSAS[u.cod]));

  const dataBase = Object.values(DATAS_ATUALIZACAO).find(Boolean);
  if (dataBase) {
    const hoje = new Date();
    const fimMes = new Date(hoje.getFullYear(), hoje.getMonth() + 1, 0);
    const fimMesStr = `${String(fimMes.getDate()).padStart(2,'0')}/${String(fimMes.getMonth()+1).padStart(2,'0')}`;
    document.getElementById('topbar-atualizacao').innerHTML =
      `Base atualizada em ${dataBase}` +
      `<span class="topbar-sep">|</span>` +
      `Tarefas com vencimento até ${fimMesStr}` +
      `<span class="topbar-sep">|</span>` +
      `Para atualizar seus dados recarregue a página ou pressione 'Ctrl+F5'`;
  }

  document.getElementById('loading').style.display = 'none';
  renderCards();
}

// ── Breadcrumb ────────────────────────────────────────────
function setBreadcrumb(partes) {
  const bar = document.getElementById('unidade-bar');
  if (!partes.length) { bar.classList.add('hidden'); return; }

  bar.classList.remove('hidden');
  document.getElementById('bar-crumbs').innerHTML = partes
    .map((p, i) => (i === 0
      ? `<span class="unidade-crumb" onclick="mostrarTela1()">${esc(p)}</span>`
      : `<span class="unidade-sep">›</span><span class="unidade-nome">${esc(p)}</span>`
    ))
    .join('');
}

// ── Placares ─────────────────────────────────────────────
function renderPlacares(containerId, rows) {
  const c = contarStatus(rows);
  const defs = [
    { classe: 'total',   valor: c.total,    label: 'Total de Tarefas', desc: 'da competência' },
    { classe: 'atraso',  valor: c.emAtraso, label: 'Em Atraso',        desc: 'tarefas' },
    { classe: 'aberto',  valor: c.emAberto, label: 'Em Aberto',        desc: 'tarefas' },
    { classe: 'baixado', valor: c.baixadas, label: 'Baixadas',         desc: 'tarefas' },
  ];

  document.getElementById(containerId).innerHTML = defs.map(p => `
    <div class="placar ${p.classe}">
      <div class="placar-label">${p.label}</div>
      <div class="placar-valor">${p.valor.toLocaleString('pt-BR')}</div>
      <div class="placar-desc">${p.desc}</div>
    </div>
  `).join('');
}

// ── Tela 1 ───────────────────────────────────────────────
function renderCards() {
  document.getElementById('cards-grid').innerHTML = UNIDADES.map(u => `
    <div class="card" onclick="mostrarTela2('${u.cod}', '${esc(u.nome)}')">
      <div class="card-unit">${esc(u.nome)}</div>
      <div class="card-footer">Clique para ver os detalhes</div>
    </div>`).join('');
}

// ── Tela 2 ───────────────────────────────────────────────
async function mostrarTela2(cod, nome) {
  if (!DADOS_UNIDADE[cod]) {
    const loading = document.getElementById('loading');
    loading.querySelector('p').textContent = `Carregando ${nome}...`;
    loading.style.display = 'flex';
    try {
      await PROMESSAS[cod];
    } finally {
      loading.style.display = 'none';
    }
  }

  if (DADOS_UNIDADE[cod] === null) {
    alert(`Erro ao carregar os dados de ${nome}. Verifique o arquivo e recarregue a página.`);
    return;
  }

  estado.unidade     = { cod, nome };
  estado.unidadeRows = DADOS_UNIDADE[cod];

  setBreadcrumb(['Painel de controle', nome]);
  renderPlacares('placares-tela2', estado.unidadeRows);
  renderDepartamentos(estado.unidadeRows);
  trocarTela('tela-2');
}

function renderDepartamentos(rows) {
  const deps = [...new Set(rows.map(r => r.Departamento).filter(Boolean))].sort();

  document.getElementById('deps-grid').innerHTML = deps.map(dep => {
    const dr = rows.filter(r => r.Departamento === dep);
    const c  = contarStatus(dr);
    return `
      <div class="dep-card" onclick="mostrarTela3('${esc(dep)}')">
        <div class="dep-nome">${esc(dep)}</div>
        <div class="dep-linhas">
          <div class="dep-linha">
            <span class="dep-dot atraso"></span>
            <span class="dep-linha-texto">Em Atraso</span>
            <span class="dep-linha-valor">${c.emAtraso.toLocaleString('pt-BR')}</span>
          </div>
          <div class="dep-linha">
            <span class="dep-dot aberto"></span>
            <span class="dep-linha-texto">Em Aberto</span>
            <span class="dep-linha-valor">${c.emAberto.toLocaleString('pt-BR')}</span>
          </div>
          <div class="dep-linha">
            <span class="dep-dot baixado"></span>
            <span class="dep-linha-texto">Baixadas</span>
            <span class="dep-linha-valor">${c.baixadas.toLocaleString('pt-BR')}</span>
          </div>
        </div>
        <div class="dep-footer">Clique aqui para ver os detalhes</div>
      </div>`;
  }).join('');
}

// ── Relatórios ────────────────────────────────────────────
function inicioSemana() {
  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0);
  const diasDesdeSegunda = hoje.getDay() === 0 ? 6 : hoje.getDay() - 1;
  const seg = new Date(hoje);
  seg.setDate(hoje.getDate() - diasDesdeSegunda);
  return seg;
}

function normData(d) {
  const n = new Date(d);
  n.setHours(0, 0, 0, 0);
  return n;
}

function rankColaboradores(rows) {
  const contagem = {};
  for (const r of rows) {
    const nome = usuarioBaixaEfetivo(r);
    if (!nome) continue;
    contagem[nome] = (contagem[nome] || 0) + 1;
  }
  return Object.entries(contagem)
    .sort((a, b) => b[1] - a[1])
    .map(([nome, qtd]) => ({ nome, qtd }));
}

function renderRelatorios(rows) {
  const agora  = new Date();
  const hoje   = normData(agora);
  const semana = inicioSemana();
  const mes    = new Date(agora.getFullYear(), agora.getMonth(), 1);
  const nomeMes = agora.toLocaleDateString('pt-BR', { month: 'long' });

  const rowsHoje   = rows.filter(r => r.DataBaixa instanceof Date && normData(r.DataBaixa) >= hoje);
  const rowsSemana = rows.filter(r => r.DataBaixa instanceof Date && normData(r.DataBaixa) >= semana);
  const rowsMes    = rows.filter(r => r.DataBaixa instanceof Date && normData(r.DataBaixa) >= mes);

  const defs = [
    { classe: 'hoje',   label: 'Baixadas hoje',          rows: rowsHoje,   desc: agora.toLocaleDateString('pt-BR') },
    { classe: 'semana', label: 'Baixadas na semana',     rows: rowsSemana, desc: 'segunda-feira até hoje' },
    { classe: 'mes',    label: `Baixadas em ${nomeMes}`, rows: rowsMes,    desc: String(agora.getFullYear()) },
  ];

  document.getElementById('baixadas-grid').innerHTML = defs.map(d => {
    const ranking = rankColaboradores(d.rows);
    const max = ranking[0]?.qtd || 1;

    const rankHTML = ranking.length
      ? ranking.map(item => `
          <div class="rank-item">
            <span class="rank-nome" title="${esc(item.nome)}">${esc(item.nome)}</span>
            <div class="rank-barra-wrapper">
              <div class="rank-barra" style="width:${(item.qtd / max * 100).toFixed(1)}%"></div>
            </div>
            <span class="rank-qtd">${item.qtd}</span>
          </div>`).join('')
      : '<span class="rank-vazio">Nenhum registro no período</span>';

    return `
      <div class="baixada-card ${d.classe}">
        <div class="baixada-card-topo">
          <div class="baixada-card-label">${d.label}</div>
          <div class="baixada-card-valor">${d.rows.length.toLocaleString('pt-BR')}</div>
          <div class="baixada-card-desc">${d.desc}</div>
        </div>
        <div class="baixada-ranking">${rankHTML}</div>
      </div>`;
  }).join('');
}

// ── Geração de PDF ───────────────────────────────────────
async function gerarPDF() {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
  const PW = doc.internal.pageSize.getWidth();
  const PH = doc.internal.pageSize.getHeight();

  const agora  = new Date();
  const ano    = agora.getFullYear();
  const mes    = agora.getMonth();
  const diasMes = new Date(ano, mes + 1, 0).getDate();
  const hoje   = agora.getDate();

  const semanas = [];
  for (let ini = 1; ini <= diasMes && ini <= hoje; ini += 7) {
    const fim = Math.min(ini + 6, diasMes, hoje);
    const label = `Semana ${semanas.length + 1}\n${String(ini).padStart(2,'0')}–${String(fim).padStart(2,'0')}/${String(mes+1).padStart(2,'0')}`;
    semanas.push({ ini, fim, label });
  }

  const base = filtrarRows(
    estado.depRows.filter(r =>
      r.Status === 'Baixado' &&
      r.DataBaixa instanceof Date &&
      r.DataBaixa.getFullYear() === ano &&
      r.DataBaixa.getMonth() === mes
    ),
    filtrosRel
  );

  const empMap = {};
  for (const r of base) {
    const nome  = usuarioBaixaEfetivo(r);
    const coord = String(r.Coordenador || '').trim();
    if (!nome) continue;
    if (!empMap[nome]) empMap[nome] = { nome, coord, qtds: new Array(semanas.length).fill(0), total: 0 };
    const dia = r.DataBaixa.getDate();
    const idx = semanas.findIndex(s => dia >= s.ini && dia <= s.fim);
    if (idx >= 0) empMap[nome].qtds[idx]++;
    empMap[nome].total++;
  }

  const empRows = Object.values(empMap).sort((a, b) => b.total - a.total || a.nome.localeCompare(b.nome, 'pt-BR'));
  const totSem  = semanas.map((_, i) => empRows.reduce((s, e) => s + e.qtds[i], 0));
  const totGeral = empRows.reduce((s, e) => s + e.total, 0);

  let logoB64 = null;
  try {
    const r = await fetch('dados/imagens/logo.png');
    const blob = await r.blob();
    logoB64 = await new Promise(res => {
      const fr = new FileReader();
      fr.onloadend = () => res(fr.result);
      fr.readAsDataURL(blob);
    });
  } catch (_) {}

  const nomeMes = agora.toLocaleDateString('pt-BR', { month: 'long', year: 'numeric' });
  const dataGer = agora.toLocaleDateString('pt-BR');
  const subHead = `${estado.unidade.nome}  —  ${estado.departamento}`;
  const filtrosAtivos = [
    filtrosRel.coordenador && `Coordenação: ${filtrosRel.coordenador}`,
    filtrosRel.responsavel && `Responsável: ${filtrosRel.responsavel}`,
    filtrosRel.grupo       && `Grupo: ${filtrosRel.grupo}`,
    filtrosRel.busca       && `Busca: "${filtrosRel.busca}"`,
  ].filter(Boolean);
  const filtrosStr = filtrosAtivos.length ? filtrosAtivos.join('  |  ') : 'Sem filtros ativos';

  function desenharCabecalho() {
    doc.setFillColor(192, 0, 0);
    doc.rect(0, 0, PW, 18, 'F');
    if (logoB64) doc.addImage(logoB64, 'PNG', 6, 3.5, 14, 10);
    doc.setTextColor(255, 255, 255);
    doc.setFontSize(13);
    doc.setFont('helvetica', 'bold');
    doc.text('MG Contécnica — Relatório de Baixas', logoB64 ? 24 : 10, 11.5);
    doc.setTextColor(18, 18, 18);
    doc.setFontSize(10);
    doc.setFont('helvetica', 'bold');
    doc.text(subHead, 10, 23);
    doc.setDrawColor(192, 0, 0);
    doc.setLineWidth(0.35);
    doc.line(10, 25.8, PW - 10, 25.8);
    doc.setTextColor(110, 110, 110);
    doc.setFontSize(7.5);
    doc.setFont('helvetica', 'normal');
    doc.text(`${nomeMes}  ·  Gerado em ${dataGer}  ·  Filtros: ${filtrosStr}`, 10, 30.5);
  }

  function desenharRodape(pageNum, totalPages) {
    doc.setFontSize(7);
    doc.setTextColor(160, 160, 160);
    doc.setFont('helvetica', 'normal');
    doc.text('MG Contécnica — Relatório Confidencial', 10, PH - 5);
    doc.text(`Página ${pageNum} de ${totalPages}`, PW - 10, PH - 5, { align: 'right' });
  }

  const cols = [
    { header: 'Funcionário', dataKey: 'nome' },
    { header: 'Coordenador', dataKey: 'coord' },
    ...semanas.map((s, i) => ({ header: s.label, dataKey: `s${i}` })),
    { header: 'Total', dataKey: 'total' },
  ];

  const tableBody = empRows.map(e => {
    const row = { nome: e.nome, coord: e.coord, total: e.total };
    e.qtds.forEach((v, i) => { row[`s${i}`] = v > 0 ? v : '—'; });
    return row;
  });
  const totalRow = { nome: 'TOTAL', coord: '', total: totGeral };
  totSem.forEach((v, i) => { totalRow[`s${i}`] = v > 0 ? v : '—'; });
  tableBody.push(totalRow);

  doc.autoTable({
    startY: 34,
    columns: cols,
    body: tableBody,
    styles: { fontSize: 8, textColor: [18, 18, 18], cellPadding: { top: 3, right: 4, bottom: 3, left: 4 } },
    headStyles: { fillColor: [26, 26, 26], textColor: [255, 255, 255], fontStyle: 'bold', halign: 'center', valign: 'middle' },
    columnStyles: {
      nome:  { cellWidth: 52 },
      coord: { cellWidth: 48 },
      total: { cellWidth: 18, halign: 'center', fontStyle: 'bold' },
      ...Object.fromEntries(semanas.map((_, i) => [`s${i}`, { halign: 'center', cellWidth: 'auto' }])),
    },
    didParseCell(data) {
      if (data.row.index === tableBody.length - 1) {
        data.cell.styles.fontStyle = 'bold';
        data.cell.styles.fillColor = [235, 235, 235];
        data.cell.styles.textColor = [30, 30, 30];
      }
      if (data.section === 'body' && data.row.index < tableBody.length - 1) {
        data.cell.styles.fillColor = data.row.index % 2 === 0 ? [244, 244, 248] : [255, 255, 255];
      }
    },
    didDrawPage() { desenharCabecalho(); },
    margin: { top: 34, left: 10, right: 10, bottom: 12 },
  });

  const chartDados = [...empRows].filter(e => e.total > 0).sort((a, b) => b.total - a.total);
  if (chartDados.length > 0) {
    doc.addPage();
    desenharCabecalho();

    const mL = 22, mR = 12, cTop = 40, cLabH = 34;
    const cH = PH - cTop - cLabH - 14;
    const cW = PW - mL - mR;
    const n = chartDados.length;
    const slotW = cW / n;
    const barW = Math.min(18, Math.max(4, slotW * 0.58));
    const maxVal = Math.max(...chartDados.map(d => d.total));
    const yMax = Math.ceil(maxVal / 5) * 5 || 5;
    const yScale = cH / yMax;

    doc.setFont('helvetica', 'bold');
    doc.setFontSize(10);
    doc.setTextColor(30, 30, 30);
    doc.text(`Baixas por Colaborador — ${nomeMes}`, mL + cW / 2, cTop - 5, { align: 'center' });

    for (let i = 0; i <= 5; i++) {
      const val = Math.round(yMax * i / 5);
      const gy = cTop + cH - val * yScale;
      doc.setDrawColor(210, 210, 210); doc.setLineWidth(0.2);
      doc.line(mL, gy, mL + cW, gy);
      doc.setFont('helvetica', 'normal'); doc.setFontSize(6); doc.setTextColor(120, 120, 120);
      doc.text(String(val), mL - 2, gy + 1.5, { align: 'right' });
    }

    doc.setDrawColor(130, 130, 130); doc.setLineWidth(0.5);
    doc.line(mL, cTop, mL, cTop + cH);
    doc.line(mL, cTop + cH, mL + cW, cTop + cH);

    chartDados.forEach((d, i) => {
      const barH = Math.max(0.5, d.total * yScale);
      const bx = mL + i * slotW + (slotW - barW) / 2;
      const by = cTop + cH - barH;
      doc.setFillColor(192, 0, 0);
      doc.rect(bx, by, barW, barH, 'F');
      doc.setFont('helvetica', 'bold'); doc.setFontSize(6); doc.setTextColor(40, 40, 40);
      doc.text(String(d.total), bx + barW / 2, by - 1.5, { align: 'center' });
      const partes = d.nome.trim().split(' ');
      const label = partes.length > 1 ? `${partes[0]} ${partes[partes.length - 1]}` : partes[0];
      doc.setFont('helvetica', 'normal'); doc.setFontSize(7); doc.setTextColor(50, 50, 50);
      doc.text(label, bx + barW / 2, cTop + cH + 2, { angle: -45 });
    });
  }

  const totalPages = doc.internal.getNumberOfPages();
  for (let p = 1; p <= totalPages; p++) {
    doc.setPage(p);
    desenharRodape(p, totalPages);
  }

  const depSlug = estado.departamento.replace(/[^a-z0-9]/gi, '_');
  doc.save(`relatorio_baixas_${depSlug}_${ano}-${String(mes+1).padStart(2,'0')}.pdf`);
}

// ── Tela 3 ───────────────────────────────────────────────
function mostrarTela3(dep) {
  estado.departamento = dep;
  estado.depRows      = estado.unidadeRows.filter(r => r.Departamento === dep);
  estado.abaAtiva     = 'Em Atraso';

  Object.assign(filtrosRel, { coordenador: '', responsavel: '', grupo: '', busca: '' });
  filtrosTab.busca = '';
  TODAS_CHAVES_TAB.forEach(k => filtrosTab.excluidos[k] = new Set());
  document.getElementById('tab-busca').value = '';
  popularFiltrosPrefixo('rel', estado.depRows);
  renderFiltroColunas();

  setBreadcrumb(['Painel de controle', estado.unidade.nome, dep]);
  renderPlacares('placares-tela3', estado.depRows);
  renderRelatorios(estado.depRows);

  document.querySelectorAll('.tab').forEach(t => t.classList.remove('ativo'));
  document.querySelector('.tab[data-status="Em Atraso"]').classList.add('ativo');

  renderTabela('Em Atraso');
  trocarTela('tela-3');
}

function voltarTela2() {
  mostrarTela2(estado.unidade.cod, estado.unidade.nome);
}

function mudarAba(btn) {
  document.querySelectorAll('.tab').forEach(t => t.classList.remove('ativo'));
  btn.classList.add('ativo');
  estado.abaAtiva = btn.dataset.status;
  fecharPaineisFiltro();
  renderFiltroColunas();
  renderTabela(estado.abaAtiva);
}

// ── Filtros ───────────────────────────────────────────────
function popularSelect(id, valores) {
  const el = document.getElementById(id);
  const atual = el.value;
  el.innerHTML = '<option value="">Todos</option>' +
    valores.map(v => `<option value="${esc(v)}">${esc(v)}</option>`).join('');
  if (valores.includes(atual)) el.value = atual;
}

function popularFiltrosPrefixo(prefixo, rows) {
  const sort = arr => arr.sort((a, b) => String(a).localeCompare(String(b), 'pt-BR'));
  popularSelect(`${prefixo}-coordenador`, sort([...new Set(rows.map(r => r.Coordenador).filter(Boolean))]));
  popularSelect(`${prefixo}-responsavel`, sort([...new Set(rows.map(r => r.UsuarioResponsavel).filter(Boolean))]));
  popularSelect(`${prefixo}-grupo`,       sort([...new Set(rows.map(r => r.Grupo).filter(Boolean))]));
}

function lerFiltros(prefixo, obj) {
  obj.coordenador = document.getElementById(`${prefixo}-coordenador`).value;
  obj.responsavel = document.getElementById(`${prefixo}-responsavel`).value;
  obj.grupo       = document.getElementById(`${prefixo}-grupo`).value;
  obj.busca       = document.getElementById(`${prefixo}-busca`).value.trim().toLowerCase();
}

function filtrarRows(rows, f) {
  return rows.filter(r => {
    if (f.coordenador && String(r.Coordenador || '')        !== f.coordenador) return false;
    if (f.responsavel && String(r.UsuarioResponsavel || '') !== f.responsavel) return false;
    if (f.grupo       && String(r.Grupo || '')              !== f.grupo)       return false;
    if (f.busca) {
      const hay = [
        r.CodCliente, r.RazaoSocial, r.Grupo, r.Titulo,
        r.UsuarioResponsavel, r.Coordenador,
        fmtCompetencia(r.Competencia), fmtData(r.DataVencimento), fmtData(r.DataPrevisaoConclusao)
      ].map(v => String(v || '').toLowerCase()).join('\0');
      if (!hay.includes(f.busca)) return false;
    }
    return true;
  });
}

function aplicarFiltroRel() {
  lerFiltros('rel', filtrosRel);
  renderRelatorios(filtrarRows(estado.depRows, filtrosRel));
}

function agendarBuscaRel() {
  clearTimeout(buscaTimerRel);
  buscaTimerRel = setTimeout(aplicarFiltroRel, 200);
}

function coordenadorMudou(prefixo) {
  const coordSel = document.getElementById(`${prefixo}-coordenador`).value;
  const sort = arr => arr.sort((a, b) => String(a).localeCompare(String(b), 'pt-BR'));
  const rowsFiltradas = coordSel
    ? estado.depRows.filter(r => String(r.Coordenador || '') === coordSel)
    : estado.depRows;

  const respAtual = document.getElementById(`${prefixo}-responsavel`).value;
  popularSelect(`${prefixo}-responsavel`, sort([...new Set(rowsFiltradas.map(r => r.UsuarioResponsavel).filter(Boolean))]));

  const opcoes = [...document.getElementById(`${prefixo}-responsavel`).options].map(o => o.value);
  if (respAtual && !opcoes.includes(respAtual)) {
    document.getElementById(`${prefixo}-responsavel`).value = '';
  }

  aplicarFiltroRel();
}

function limparFiltros(prefixo) {
  const obj = filtrosRel;
  [`${prefixo}-coordenador`, `${prefixo}-responsavel`, `${prefixo}-grupo`]
    .forEach(id => document.getElementById(id).value = '');
  document.getElementById(`${prefixo}-busca`).value = '';
  obj.coordenador = ''; obj.responsavel = ''; obj.grupo = ''; obj.busca = '';
  const sort = arr => arr.sort((a, b) => String(a).localeCompare(String(b), 'pt-BR'));
  popularSelect(`${prefixo}-responsavel`, sort([...new Set(estado.depRows.map(r => r.UsuarioResponsavel).filter(Boolean))]));
  renderRelatorios(estado.depRows);
}

// ── Filtros da tabela de Tarefas (estilo Excel) ────────────
function renderFiltroColunas() {
  const cont = document.getElementById('filtro-colunas-tab');
  cont.innerHTML = colunasTab().map(col => `
    <div class="filtro-coluna" id="filtro-coluna-${col.key}">
      <button type="button" class="filtro-col-btn" onclick="toggleFiltroColuna('${col.key}')">
        <span class="filtro-col-label">${esc(col.label)}</span>
        <span class="filtro-col-seta">&#9662;</span>
      </button>
      <div class="filtro-col-painel hidden" id="painel-${col.key}"></div>
    </div>
  `).join('');
  atualizarIndicadoresFiltro();
}

function atualizarIndicadoresFiltro() {
  colunasTab().forEach(col => {
    const btn = document.querySelector(`#filtro-coluna-${col.key} .filtro-col-btn`);
    if (btn) btn.classList.toggle('filtro-ativo', filtrosTab.excluidos[col.key].size > 0);
  });
}

function linhasBaseTab() {
  return estado.depRows.filter(r => r.Status === estado.abaAtiva);
}

function aplicarExcluidos(rows, excluidos, ignorarCol) {
  return rows.filter(r => {
    for (const col of colunasTab()) {
      if (col.key === ignorarCol) continue;
      const excl = excluidos[col.key];
      if (excl && excl.size && excl.has(col.valor(r))) return false;
    }
    return true;
  });
}

function aplicarBuscaTab(rows) {
  if (!filtrosTab.busca) return rows;
  const b = filtrosTab.busca;
  return rows.filter(r => {
    const hay = [
      r.CodCliente, r.RazaoSocial, r.Grupo, r.Titulo,
      estado.abaAtiva === 'Baixado' ? usuarioBaixaEfetivo(r) : r.UsuarioResponsavel, r.Coordenador,
      fmtCompetencia(r.Competencia), fmtData(r.DataVencimento),
      estado.abaAtiva === 'Baixado' ? fmtData(r.DataBaixa) : fmtData(r.DataPrevisaoConclusao)
    ].map(v => String(v || '').toLowerCase()).join('\0');
    return hay.includes(b);
  });
}

function filtrarRowsTab(rows) {
  return aplicarBuscaTab(aplicarExcluidos(rows, filtrosTab.excluidos, null));
}

function toggleFiltroColuna(key) {
  const painel = document.getElementById(`painel-${key}`);
  const estavaAberto = !painel.classList.contains('hidden');
  fecharPaineisFiltro();
  if (estavaAberto) return;
  abrirPainelFiltro(key);
}

function fecharPaineisFiltro() {
  document.querySelectorAll('.filtro-col-painel').forEach(p => p.classList.add('hidden'));
  filtroColunaAberta = null;
  draftExcluidos = null;
}

function abrirPainelFiltro(key) {
  const col = colunasTab().find(c => c.key === key);
  const painel = document.getElementById(`painel-${key}`);

  const rowsBase = aplicarBuscaTab(aplicarExcluidos(linhasBaseTab(), filtrosTab.excluidos, key));
  const valoresOrdem = new Map();
  rowsBase.forEach(r => {
    const v = col.valor(r);
    if (!valoresOrdem.has(v)) valoresOrdem.set(v, col.ordem ? col.ordem(r) : v);
  });
  const opcoes = [...valoresOrdem.keys()].sort((a, b) => {
    if (a === '') return 1;
    if (b === '') return -1;
    return col.ordem
      ? valoresOrdem.get(a) - valoresOrdem.get(b)
      : String(a).localeCompare(String(b), 'pt-BR');
  });

  draftExcluidos = new Set(filtrosTab.excluidos[key]);
  filtroColunaAberta = key;

  painel.innerHTML = `
    <input type="text" class="filtro-col-busca" placeholder="Pesquisar..." oninput="filtrarOpcoesColuna(this.value)" />
    <label class="filtro-col-opcao filtro-col-marcartudo">
      <input type="checkbox" id="chk-marcar-tudo" ${draftExcluidos.size === 0 ? 'checked' : ''} onchange="marcarTudoColuna(this.checked)" />
      <span>Selecionar tudo</span>
    </label>
    <div class="filtro-col-lista" id="lista-${key}">
      ${opcoes.length ? opcoes.map(v => `
        <label class="filtro-col-opcao" data-valor="${esc(v)}">
          <input type="checkbox" value="${esc(v)}" ${draftExcluidos.has(v) ? '' : 'checked'} onchange="alternarOpcaoColuna(this)" />
          <span title="${v ? esc(v) : '(Vazias)'}">${v ? esc(v) : '<em>(Vazias)</em>'}</span>
        </label>
      `).join('') : '<div class="filtro-col-vazio">Nenhuma opção</div>'}
    </div>
    <div class="filtro-col-acoes">
      <button type="button" class="filtro-col-cancelar" onclick="fecharPaineisFiltro()">Cancelar</button>
      <button type="button" class="filtro-col-ok" onclick="aplicarFiltroColuna()">OK</button>
    </div>
  `;
  painel.classList.remove('hidden');
  painel.querySelector('.filtro-col-busca').focus();
}

function filtrarOpcoesColuna(texto) {
  const t = texto.trim().toLowerCase();
  document.querySelectorAll(`#lista-${filtroColunaAberta} .filtro-col-opcao`).forEach(el => {
    const valor = (el.dataset.valor || '').toLowerCase();
    el.classList.toggle('hidden', Boolean(t) && !valor.includes(t));
  });
}

function alternarOpcaoColuna(chk) {
  if (chk.checked) draftExcluidos.delete(chk.value);
  else draftExcluidos.add(chk.value);
  const todasMarcadas = ![...document.querySelectorAll(`#lista-${filtroColunaAberta} input[type=checkbox]`)]
    .some(c => !c.checked);
  document.getElementById('chk-marcar-tudo').checked = todasMarcadas;
}

function marcarTudoColuna(checked) {
  document.querySelectorAll(`#lista-${filtroColunaAberta} input[type=checkbox]`).forEach(c => {
    c.checked = checked;
    if (checked) draftExcluidos.delete(c.value);
    else draftExcluidos.add(c.value);
  });
}

function aplicarFiltroColuna() {
  if (!filtroColunaAberta) return;
  filtrosTab.excluidos[filtroColunaAberta] = draftExcluidos;
  fecharPaineisFiltro();
  atualizarIndicadoresFiltro();
  renderTabela(estado.abaAtiva);
}

function aplicarFiltroTab() {
  filtrosTab.busca = document.getElementById('tab-busca').value.trim().toLowerCase();
  renderTabela(estado.abaAtiva);
}

function agendarBuscaTab() {
  clearTimeout(buscaTimerTab);
  buscaTimerTab = setTimeout(aplicarFiltroTab, 200);
}

function limparFiltrosTab() {
  TODAS_CHAVES_TAB.forEach(k => filtrosTab.excluidos[k] = new Set());
  filtrosTab.busca = '';
  document.getElementById('tab-busca').value = '';
  fecharPaineisFiltro();
  atualizarIndicadoresFiltro();
  renderTabela(estado.abaAtiva);
}

document.addEventListener('click', (e) => {
  if (!filtroColunaAberta) return;
  const painelAberto = document.getElementById(`painel-${filtroColunaAberta}`);
  const btnAberto = document.querySelector(`#filtro-coluna-${filtroColunaAberta} .filtro-col-btn`);
  if (painelAberto && !painelAberto.contains(e.target) && btnAberto && !btnAberto.contains(e.target)) {
    fecharPaineisFiltro();
  }
});

document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape' && filtroColunaAberta) fecharPaineisFiltro();
});

function sortRows(rows) {
  return [...rows].sort((a, b) => {
    const tA = a.DataVencimento instanceof Date ? a.DataVencimento.getTime() : 0;
    const tB = b.DataVencimento instanceof Date ? b.DataVencimento.getTime() : 0;
    if (tA !== tB) return tA - tB;
    const cA = String(a.RazaoSocial || a.CodCliente || '').toLowerCase();
    const cB = String(b.RazaoSocial || b.CodCliente || '').toLowerCase();
    return cA.localeCompare(cB, 'pt-BR');
  });
}

function renderTabela(status) {
  const rows = sortRows(filtrarRowsTab(estado.depRows.filter(r => r.Status === status)));
  const corpo = document.getElementById('tabela-corpo');
  const contador = document.getElementById('tab-contador');
  const ehBaixado = status === 'Baixado';

  const thPrevisaoBaixa = document.getElementById('th-previsao-baixa');
  if (thPrevisaoBaixa) thPrevisaoBaixa.textContent = ehBaixado ? 'Data de baixa' : 'Previsão';

  contador.innerHTML = rows.length
    ? `<span>${rows.length.toLocaleString('pt-BR')}</span> Tarefa${rows.length !== 1 ? 's' : ''}`
    : '';

  if (!rows.length) {
    corpo.innerHTML = `<tr><td colspan="9" class="tabela-vazia">Nenhuma tarefa encontrada.</td></tr>`;
    return;
  }

  corpo.innerHTML = rows.map((r, idx) => {
    const codCliente = esc(r.CodCliente || '');
    const razao      = esc(r.RazaoSocial || '');
    const cliente    = codCliente && razao ? `${codCliente} - ${razao}` : codCliente || razao;

    const cmt = r.Comentario ? String(r.Comentario).trim() : '';
    const colCmt = cmt
      ? `<button class="btn-comentario" onclick="toggleComentario(this,'cmt-${idx}')">Ver comentário</button>
         <span class="comentario-texto hidden" id="cmt-${idx}">${esc(cmt)}</span>`
      : '-';

    return `<tr>
      <td>${cliente}</td>
      <td>${esc(r.Grupo)}</td>
      <td>${esc(r.Titulo)}</td>
      <td>${ehBaixado ? esc(usuarioBaixaEfetivo(r)) : esc(r.UsuarioResponsavel)}</td>
      <td>${fmtCompetencia(r.Competencia)}</td>
      <td>${fmtData(r.DataVencimento)}</td>
      <td>${ehBaixado ? fmtData(r.DataBaixa) : fmtData(r.DataPrevisaoConclusao)}</td>
      <td>${esc(r.Prioridade)}</td>
      <td>${colCmt}</td>
    </tr>`;
  }).join('');
}

function toggleComentario(btn, id) {
  const span = document.getElementById(id);
  const aberto = !span.classList.contains('hidden');
  span.classList.toggle('hidden');
  btn.textContent = aberto ? 'Ver comentário' : 'Ocultar';
}

// ── Navegação ─────────────────────────────────────────────
function trocarTela(id) {
  ['tela-1', 'tela-2', 'tela-3'].forEach(t =>
    document.getElementById(t).classList.toggle('hidden', t !== id)
  );
  if (id === 'tela-1') setBreadcrumb([]);
  window.scrollTo(0, 0);
}

function mostrarTela1() {
  trocarTela('tela-1');
}

// ── Tema ─────────────────────────────────────────────────
function alternarTema() {
  const escuro = document.body.classList.toggle('tema-escuro');
  document.querySelector('.btn-tema').textContent =
    escuro ? 'Alterar tema para Claro' : 'Alterar tema para Escuro';
}

// ── Init ─────────────────────────────────────────────────
iniciarCarregamento();
