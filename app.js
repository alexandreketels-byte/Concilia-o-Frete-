// ============ ESTADO GLOBAL ============
const state = {
  files: { relatorio: null, cidades: null, tarifa: null, seccat: null },
  parsed: { relatorio: null, cidades: null, tarifa: null, seccat: null },
  cityLookup: {},
  secCatLookup: {},
  conciliacaoRows: [],
  cidadesList: [],
};

// ============ HELPERS ============
function brNum(v) {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return v;
  let s = String(v).trim();
  if (s.includes(',') ) s = s.replace(/\./g, '').replace(',', '.');
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}
function fmtBRL(v) {
  if (v === null || v === undefined || isNaN(v)) return '—';
  return v.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
}
function fmtNum(v, dec = 2) {
  if (v === null || v === undefined || isNaN(v)) return '—';
  return v.toLocaleString('pt-BR', { minimumFractionDigits: dec, maximumFractionDigits: dec });
}
function splitCidadeUF(str) {
  if (!str) return { cidade: '', uf: '' };
  const parts = str.split('/');
  if (parts.length === 2) return { cidade: parts[0].trim(), uf: parts[1].trim() };
  // fallback: últimos 2 caracteres como UF
  const trimmed = str.trim();
  return { cidade: trimmed.slice(0, -2).trim(), uf: trimmed.slice(-2).trim() };
}
function readFileAsArrayBuffer(file) {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = () => res(r.result);
    r.onerror = rej;
    r.readAsArrayBuffer(file);
  });
}
function readFileAsText(file, encoding = 'ISO-8859-1') {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = () => res(r.result);
    r.onerror = rej;
    r.readAsText(file, encoding);
  });
}

// ============ PARSERS ============

// Relatório de frete: HTML disfarçado de .xls, OU xlsx real
async function parseRelatorioFrete(file) {
  const name = file.name.toLowerCase();
  let rows = [];

  if (name.endsWith('.xlsx')) {
    const buf = await readFileAsArrayBuffer(file);
    const wb = XLSX.read(buf, { type: 'array' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
  } else {
    // tenta como HTML primeiro
    const text = await readFileAsText(file, 'ISO-8859-1');
    if (text.includes('<table') || text.includes('<TABLE')) {
      const parser = new DOMParser();
      const doc = parser.parseFromString(text, 'text/html');
      const table = doc.querySelector('table');
      if (table) {
        rows = Array.from(table.querySelectorAll('tr')).map(tr =>
          Array.from(tr.querySelectorAll('td,th')).map(td => td.textContent.trim())
        );
      }
    } else {
      const buf = await readFileAsArrayBuffer(file);
      const wb = XLSX.read(buf, { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
    }
  }

  const pedidos = [];
  let transportadoraAtual = null;
  let lastPedidoNF = { pedido: null, nf: null, cidadeUF: null, cnpj: null };

  for (const row of rows) {
    const vals = (row || []).map(v => (v === null || v === undefined) ? '' : String(v).trim());
    const nonEmpty = vals.filter(v => v !== '');
    if (nonEmpty.length === 0) continue;

    if (nonEmpty.length === 1 && nonEmpty[0].toLowerCase().startsWith('transportadora')) {
      transportadoraAtual = nonEmpty[0].replace(/transportadora:\s*/i, '').trim();
      continue;
    }
    if (vals[0] === 'Pedido') continue; // header repetido por bloco
    if (nonEmpty.length === 1 && /total/i.test(nonEmpty[0])) continue; // linha de total textual

    // Linha completa (15 colunas) com Pedido preenchido
    if (vals.length >= 14 && vals[4] && /entrega normal|devolu|reentrega/i.test(vals[4])) {
      const { cidade, uf } = splitCidadeUF(vals[2]);
      const pedido = {
        transportadora: transportadoraAtual,
        pedido: vals[0] || lastPedidoNF.pedido,
        nf: vals[1] || lastPedidoNF.nf,
        cidade, uf,
        cidadeRaw: vals[2],
        cnpj: vals[3] || lastPedidoNF.cnpj,
        tipoFrete: vals[4],
        freteCobrado: brNum(vals[5]),
        preco: brNum(vals[6]),
        cubagem: brNum(vals[8]),
        qtdeVolumes: brNum(vals[9]),
        peso: brNum(vals[10]),
        venda: brNum(vals[11]),
        dataLancto: vals[14] || '',
      };
      if (vals[0]) lastPedidoNF = { pedido: vals[0], nf: vals[1], cidadeUF: vals[2], cnpj: vals[3] };
      pedidos.push(pedido);
    }
    // Linha de devolução/reentrega sem Pedido preenchido (shift de coluna): NF na posição 0, tipo na posição 3
    else if (vals.length >= 13 && vals[3] && /devolu|reentrega/i.test(vals[3]) && !vals[4]) {
      const { cidade, uf } = splitCidadeUF(vals[1]);
      const pedido = {
        transportadora: transportadoraAtual,
        pedido: lastPedidoNF.pedido,
        nf: vals[0] || lastPedidoNF.nf,
        cidade, uf,
        cidadeRaw: vals[1],
        cnpj: vals[2] || lastPedidoNF.cnpj,
        tipoFrete: vals[3],
        freteCobrado: brNum(vals[4]),
        preco: brNum(vals[5]),
        cubagem: brNum(vals[7]),
        qtdeVolumes: brNum(vals[8]),
        peso: brNum(vals[9]),
        venda: brNum(vals[10]),
        dataLancto: vals[13] || '',
      };
      pedidos.push(pedido);
    }
  }
  return pedidos;
}

// Relação de cidades: colunas fixas (D=CIDADE DESTINO idx5, E=UF idx6, H=CHAVE idx7, Q=POLO idx17)
async function parseRelacaoCidades(file) {
  const buf = await readFileAsArrayBuffer(file);
  const wb = XLSX.read(buf, { type: 'array' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });

  const lookup = {};
  const list = [];
  // descobre dinamicamente os índices pelo cabeçalho (linha que contém 'CHAVE' e 'POLO')
  let headerIdx = -1, colCidade = -1, colUF = -1, colChave = -1, colPolo = -1, colRegiao = -1;
  for (let i = 0; i < Math.min(rows.length, 5); i++) {
    const r = (rows[i] || []).map(v => (v || '').toString().toUpperCase().trim());
    if (r.includes('CHAVE') && r.includes('POLO')) {
      headerIdx = i;
      colCidade = r.indexOf('CIDADE DESTINO');
      colUF = r.lastIndexOf('UF');
      colChave = r.indexOf('CHAVE');
      colPolo = r.indexOf('POLO');
      colRegiao = r.indexOf('REGIÃO SUPER') >= 0 ? r.indexOf('REGIÃO SUPER') : r.findIndex(x => x.includes('REGI'));
      break;
    }
  }
  if (headerIdx === -1) throw new Error('Não encontrei as colunas CHAVE/POLO na relação de cidades. Verifique o arquivo.');

  for (let i = headerIdx + 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r || !r[colChave]) continue;
    const chave = String(r[colChave]).trim().toUpperCase()
      .normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    const cidade = colCidade >= 0 ? String(r[colCidade] || '').trim() : '';
    const uf = colUF >= 0 ? String(r[colUF] || '').trim() : '';
    const polo = colPolo >= 0 ? String(r[colPolo] || '').trim().toUpperCase() : '';
    const regiao = colRegiao >= 0 ? String(r[colRegiao] || '').trim() : '';
    lookup[chave] = { cidade, uf, polo, regiao };
    if (cidade && uf) list.push(`${cidade} / ${uf}`);
  }
  return { lookup, list };
}

// Tabela de tarifa frete-peso: extrai bloco de faixas (linha 'Faixa de KM'/'0 - 10 kg' até linha em branco)
async function parseTabelaTarifa(file) {
  const buf = await readFileAsArrayBuffer(file);
  const wb = XLSX.read(buf, { type: 'array' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });

  let headerIdx = -1, labelCol = -1, valueCols = [];
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i] || [];
    const faixaIdx = r.findIndex(v => /faixa de km/i.test(v || ''));
    if (faixaIdx >= 0) {
      const kgIdx = [];
      for (let c = faixaIdx + 1; c < r.length; c++) {
        if (/kg/i.test(r[c] || '')) kgIdx.push(c);
      }
      if (kgIdx.length >= 5) {
        headerIdx = i;
        labelCol = faixaIdx;
        valueCols = kgIdx;
        break;
      }
    }
  }
  if (headerIdx === -1) throw new Error('Não encontrei o cabeçalho de faixas de peso na tabela de tarifa.');

  const table = {};
  let i = headerIdx + 1;
  for (; i < rows.length; i++) {
    const r = rows[i];
    if (!r) continue;
    // primeira coluna não-nula a partir do início é o rótulo (UF/classe), pode estar em qualquer coluna antes de labelCol+1
    let label = null;
    for (let c = 0; c <= labelCol; c++) {
      if (r[c] !== null && r[c] !== undefined && String(r[c]).trim() !== '') { label = String(r[c]).trim().toUpperCase(); break; }
    }
    if (!label) {
      // linha em branco: se já temos pelo menos uma faixa carregada, encerra o bloco
      if (Object.keys(table).length > 0) break;
      continue;
    }
    const values = valueCols.map(c => brNum(r[c]));
    if (values.every(v => v === 0)) {
      if (Object.keys(table).length > 0) break;
      continue;
    }
    table[label] = values;
  }

  // Generalidades: percorre o resto buscando por rótulos conhecidos
  const general = {
    cubagemFatorKgM3: 200, freteValorPct: 0.003, freteValorMinimo: 3.549,
    freteValorPctNorte: 0.004, freteValorMinimoNorte: 7.848,
    taxaDespacho: 23.44, taxaDespachoUFs: [],
    tasValor: 5.427467, tasUFs: [],
    tasSuframaValor: 10.705651999999999, tasSuframaUFs: [],
    grisMinimo: 1.36, grisPct: 0.001, grisMinimoRegiao2: 1.15, grisPctRegiao2: 0.001, grisRegiao2UFs: [],
    grisMinimoRegiao3: 2.47, grisPctRegiao3: 0.003, grisRegiao3UFs: [],
    pedagioPorFracao100kg: 5.427467, pedagioRegiaoNorteValor: 8.786312, pedagioRegiaoNorteUFs: [],
    reentregaPct: 0.5, devolucaoPct: 1.0,
    icmsAliquota: 0.12,
  };

  function extractUFs(text) {
    const matches = text.toUpperCase().match(/\b[A-Z]{2}\b/g) || [];
    const knownUFs = ['AC','AL','AP','AM','BA','CE','DF','ES','GO','MA','MT','MS','MG','PA','PB','PR','PE','PI','RJ','RN','RS','RO','RR','SC','SP','SE','TO'];
    return [...new Set(matches.filter(m => knownUFs.includes(m)))];
  }
  let grisBlock = null; // 'geral' | 'regiao2' | 'regiao3'
  for (; i < rows.length; i++) {
    const r = rows[i];
    if (!r) continue;
    const nonNull = r.map((v, idx) => (v !== null && v !== undefined && String(v).trim() !== '') ? { idx, v } : null).filter(Boolean);
    if (nonNull.length === 0) continue;
    const label = String(nonNull[0].v).toLowerCase();
    const lastVal = nonNull[nonNull.length - 1].v;
    const val = brNum(lastVal);

    if (label.includes('cubagem')) general.cubagemFatorKgM3 = val || general.cubagemFatorKgM3;
    if (label.includes('frete valor norte')) general.freteValorPctNorte = val;
    if (label.includes('minimo norte') || label.includes('mínimo norte')) general.freteValorMinimoNorte = val;
    if (label.includes('frete valor') && !label.includes('norte') && !label.includes('mínimo') && !label.includes('minimo')) general.freteValorPct = val;
    if (label.includes('mínimo do frete valor') || label.includes('minimo do frete valor')) general.freteValorMinimo = val;
    if (label.includes('taxa de despacho')) { general.taxaDespacho = val; general.taxaDespachoUFs = extractUFs(label); }
    const isTAS = /tas\s*\(taxa administrativa sefaz/i.test(label);
    if (isTAS && label.includes('suframa')) { general.tasSuframaValor = val; general.tasSuframaUFs = extractUFs(label); }
    else if (isTAS) { general.tasValor = val; general.tasUFs = extractUFs(label); }

    // Blocos GRIS: a linha "GRIS (...)" define o contexto para as 2 linhas seguintes
    if (label.includes('gris (gerenciamento de risco)')) {
      if (/rj,\s*es,\s*mt,\s*ms,\s*ac,\s*ro/i.test(label)) { grisBlock = 'regiao2'; general.grisRegiao2UFs = extractUFs(label); }
      else if (/pa,\s*rr,\s*ap,\s*to/i.test(label)) { grisBlock = 'regiao3'; general.grisRegiao3UFs = extractUFs(label); }
      else grisBlock = 'geral';
      continue;
    }
    if (label.includes('mínimo') && label.includes('gris') || label.includes('minimo') && label.includes('gris')) {
      if (grisBlock === 'regiao2') general.grisMinimoRegiao2 = val;
      else if (grisBlock === 'regiao3') general.grisMinimoRegiao3 = val;
      else general.grisMinimo = val;
      continue;
    }
    if ((label.includes('% sobre valor da mercadoria') || (label.includes('notas fiscais') && label.includes('%'))) && grisBlock) {
      if (grisBlock === 'regiao2') general.grisPctRegiao2 = val;
      else if (grisBlock === 'regiao3') general.grisPctRegiao3 = val;
      else if (grisBlock === 'geral') general.grisPct = val;
      grisBlock = null; // bloco GRIS consumido, evita capturar % de outras seções (ex: TFD) depois
      continue;
    }

    if (label.includes('pedágio') || label.includes('pedagio')) {
      if (/pa,\s*rr,\s*ap,\s*to/i.test(label)) { general.pedagioRegiaoNorteValor = val; general.pedagioRegiaoNorteUFs = extractUFs(label); }
      else general.pedagioPorFracao100kg = val;
    }
    if (label.trim().startsWith('- devolução') || label.trim().startsWith('- devolucao')) general.devolucaoPct = val || 1.0;
    if (label.trim().startsWith('- reentrega')) general.reentregaPct = 0.5;
  }

  return { table, general };
}

// Taxas SEC-CAT por cidade
async function parseSecCat(file) {
  const buf = await readFileAsArrayBuffer(file);
  const wb = XLSX.read(buf, { type: 'array' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });

  let headerIdx = -1, colNome = -1, colValor = -1, colChave = -1;
  for (let i = 0; i < Math.min(rows.length, 5); i++) {
    const r = (rows[i] || []).map(v => (v || '').toString().toUpperCase().trim());
    if (r.includes('CHAVE')) {
      headerIdx = i;
      colNome = r.indexOf('NOME TARIFA');
      colValor = r.indexOf('VALOR');
      colChave = r.indexOf('CHAVE');
      break;
    }
  }
  if (headerIdx === -1) throw new Error('Não encontrei a coluna CHAVE na planilha de taxas SEC-CAT.');

  const lookup = {};
  const seen = {}; // chave -> { nomeTarifa: valor } para deduplicar
  for (let i = headerIdx + 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r || !r[colChave]) continue;
    const chave = String(r[colChave]).trim().toUpperCase()
      .normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    const nomeTarifa = colNome >= 0 ? String(r[colNome] || '').trim() : '';
    const valor = colValor >= 0 ? brNum(r[colValor]) : 0;
    if (!seen[chave]) seen[chave] = {};
    // Algumas combinações cidade+tarifa aparecem duplicadas com valores diferentes
    // (provável atualização de tabela). Mantém o maior valor por segurança.
    if (!(nomeTarifa in seen[chave]) || valor > seen[chave][nomeTarifa]) {
      seen[chave][nomeTarifa] = valor;
    }
  }
  for (const chave in seen) {
    lookup[chave] = Object.entries(seen[chave]).map(([nomeTarifa, valor]) => ({ nomeTarifa, valor }));
  }
  return lookup;
}

// ============ MOTOR (usa freight-engine.browser.js + dados carregados) ============
function buildEngineContext() {
  const tarifa = state.parsed.tarifa;
  const general = tarifa ? { ...tarifa.general } : {};
  const icmsInput = document.getElementById('icms-aliquota');
  if (icmsInput) {
    const pct = parseFloat(icmsInput.value.replace(',', '.'));
    general.icmsAliquota = isNaN(pct) ? 0.12 : pct / 100;
  }
  return {
    table: tarifa ? tarifa.table : {},
    general,
  };
}

function calcularComContexto(pedido) {
  const ctx = buildEngineContext();
  return calcularFreteTeorico(pedido, state.cityLookup, state.secCatLookup, ctx);
}

// ============ UI: UPLOAD ============
const fileBoxes = {
  relatorio: { input: 'file-relatorio', box: 'box-relatorio', status: 'status-relatorio' },
  cidades: { input: 'file-cidades', box: 'box-cidades', status: 'status-cidades' },
  tarifa: { input: 'file-tarifa', box: 'box-tarifa', status: 'status-tarifa' },
  seccat: { input: 'file-seccat', box: 'box-seccat', status: 'status-seccat' },
};

Object.entries(fileBoxes).forEach(([key, ids]) => {
  const input = document.getElementById(ids.input);
  input.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (!file) return;
    state.files[key] = file;
    document.getElementById(ids.box).classList.add('filled');
    document.getElementById(ids.status).textContent = file.name;
    updateProcessButton();
  });
});

function updateProcessButton() {
  const allFilled = Object.values(state.files).every(f => f !== null);
  document.getElementById('btn-process').disabled = !allFilled;
}

document.getElementById('btn-process').addEventListener('click', processarArquivos);

async function processarArquivos() {
  const errEl = document.getElementById('error-list');
  errEl.style.display = 'none';
  errEl.innerHTML = '';
  const statusEl = document.getElementById('process-status');
  statusEl.textContent = 'Processando...';
  const errors = [];

  try {
    state.parsed.relatorio = await parseRelatorioFrete(state.files.relatorio);
  } catch (e) { errors.push('Relatório de Frete: ' + e.message); }

  try {
    const r = await parseRelacaoCidades(state.files.cidades);
    state.parsed.cidades = r;
    state.cityLookup = r.lookup;
    state.cidadesList = r.list;
  } catch (e) { errors.push('Relação de Cidades: ' + e.message); }

  try {
    state.parsed.tarifa = await parseTabelaTarifa(state.files.tarifa);
  } catch (e) { errors.push('Tabela de Tarifa: ' + e.message); }

  try {
    state.secCatLookup = await parseSecCat(state.files.seccat);
  } catch (e) { errors.push('Taxas SEC-CAT: ' + e.message); }

  if (errors.length) {
    errEl.style.display = 'block';
    errEl.innerHTML = errors.map(e => `<div>⚠ ${e}</div>`).join('');
    statusEl.textContent = 'Concluído com pendências.';
  } else {
    statusEl.textContent = '✓ Arquivos processados com sucesso.';
  }

  renderDataSummary();
  if (state.parsed.relatorio) {
    rodarConciliacao();
  }
  populateCidadesDatalist();
  if (state.cityLookup && Object.keys(state.cityLookup).length && state.parsed.tarifa) {
    document.getElementById('cotacao-empty').style.display = 'none';
    document.getElementById('cotacao-content').style.display = 'block';
  }
}

function renderDataSummary() {
  const card = document.getElementById('preview-card');
  const el = document.getElementById('data-summary');
  card.style.display = 'block';
  const nPedidos = state.parsed.relatorio ? state.parsed.relatorio.length : 0;
  const nCidades = state.cidadesList.length;
  const nFaixas = state.parsed.tarifa ? Object.keys(state.parsed.tarifa.table).length : 0;
  const nSecCat = Object.keys(state.secCatLookup).length;
  el.innerHTML = `
    <div class="stat"><div class="v">${nPedidos}</div><div class="l">Pedidos no relatório</div></div>
    <div class="stat"><div class="v">${nCidades}</div><div class="l">Cidades cadastradas</div></div>
    <div class="stat"><div class="v">${nFaixas}</div><div class="l">Combinações UF/classe na tarifa</div></div>
    <div class="stat"><div class="v">${nSecCat}</div><div class="l">Cidades com taxa local cadastrada</div></div>
  `;
}

function populateCidadesDatalist() {
  const dl = document.getElementById('cidades-list');
  dl.innerHTML = '';
  const seen = new Set();
  state.cidadesList.slice(0, 2000).forEach(c => {
    if (seen.has(c)) return;
    seen.add(c);
    const opt = document.createElement('option');
    opt.value = c.split(' / ')[0];
    dl.appendChild(opt);
  });
}

// ============ TABS ============
document.querySelectorAll('.tab').forEach(tab => {
  tab.addEventListener('click', () => {
    document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
    document.querySelectorAll('.panel').forEach(p => p.classList.remove('active'));
    tab.classList.add('active');
    document.getElementById('panel-' + tab.dataset.tab).classList.add('active');
  });
});

// ============ CONCILIAÇÃO ============
const TRANSPORTADORAS_SUPORTADAS = ['RODONAVES'];

function rodarConciliacao() {
  const pedidos = state.parsed.relatorio;
  if (!pedidos) return;

  const rows = pedidos.map(p => {
    const transpNorm = (p.transportadora || '').toUpperCase().trim();
    if (!TRANSPORTADORAS_SUPORTADAS.includes(transpNorm)) {
      return {
        ...p,
        calc: { erro: `Transportadora "${p.transportadora}" ainda não tem tabela cadastrada no sistema`, valorTotal: null },
        diff: null, diffPct: null, status: 'sem-tabela',
      };
    }
    const calc = calcularComContexto({
      cidade: p.cidade, uf: p.uf, peso: p.peso, cubagem: p.cubagem,
      valorMercadoria: p.preco, tipoFrete: p.tipoFrete, freteOriginal: p.preco ? null : null,
    });
    let status = 'ok', diff = null, diffPct = null;
    if (calc.erro) {
      status = 'erro';
    } else {
      diff = p.freteCobrado - calc.valorTotal;
      diffPct = calc.valorTotal ? (diff / calc.valorTotal) * 100 : 0;
      status = Math.abs(diffPct) <= 5 ? 'ok' : 'divergente';
    }
    return { ...p, calc, diff, diffPct, status };
  });

  state.conciliacaoRows = rows;
  document.getElementById('conciliacao-empty').style.display = 'none';
  document.getElementById('conciliacao-content').style.display = 'block';

  const transportadoras = [...new Set(rows.map(r => r.transportadora))].sort();
  const sel = document.getElementById('filtro-transportadora');
  sel.innerHTML = '<option value="">Todas as transportadoras</option>' +
    transportadoras.map(t => `<option value="${t}">${t}</option>`).join('');

  renderConciliacaoSummary(rows);
  renderConciliacaoTable(rows);
}

function renderConciliacaoSummary(rows) {
  const ok = rows.filter(r => r.status === 'ok').length;
  const div = rows.filter(r => r.status === 'divergente').length;
  const err = rows.filter(r => r.status === 'erro').length;
  const semTabela = rows.filter(r => r.status === 'sem-tabela').length;
  const totalCobrado = rows.reduce((s, r) => s + (r.freteCobrado || 0), 0);

  document.getElementById('conciliacao-summary').innerHTML = `
    <div class="stat ok"><div class="v">${ok}</div><div class="l">OK (dentro de 5%)</div></div>
    <div class="stat warn"><div class="v">${div}</div><div class="l">Divergentes</div></div>
    <div class="stat err"><div class="v">${err + semTabela}</div><div class="l">Erro / sem tabela cadastrada</div></div>
    <div class="stat"><div class="v">${fmtBRL(totalCobrado)}</div><div class="l">Total cobrado (${rows.length} pedidos)</div></div>
  `;
}

function statusBadge(status) {
  if (status === 'ok') return '<span class="badge ok">OK</span>';
  if (status === 'divergente') return '<span class="badge warn">Divergente</span>';
  if (status === 'sem-tabela') return '<span class="badge err">Sem tabela</span>';
  return '<span class="badge err">Erro</span>';
}

function renderConciliacaoTable(rows) {
  const tbody = document.getElementById('tbody-conciliacao');
  tbody.innerHTML = '';
  rows.forEach((r, idx) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><span class="toggle-btn" data-idx="${idx}">▸ ver</span></td>
      <td>${r.pedido || ''}</td>
      <td>${r.nf || ''}</td>
      <td>${r.cidadeRaw || ''}</td>
      <td>${r.transportadora || ''}</td>
      <td>${r.tipoFrete || ''}</td>
      <td class="num">${fmtNum(r.peso)}</td>
      <td class="num">${fmtBRL(r.freteCobrado)}</td>
      <td class="num">${r.calc.erro ? '—' : fmtBRL(r.calc.valorTotal)}</td>
      <td class="num">${r.diff === null ? '—' : fmtBRL(r.diff)}</td>
      <td class="num">${r.diffPct === null ? '—' : fmtNum(r.diffPct, 1) + '%'}</td>
      <td>${statusBadge(r.status)}</td>
    `;
    tbody.appendChild(tr);

    const detailTr = document.createElement('tr');
    detailTr.className = 'detail-row';
    detailTr.id = `detail-${idx}`;
    let detailHtml = '';
    if (r.calc.erro) {
      detailHtml = `<span style="color:var(--err)">${r.calc.erro}</span>`;
    } else {
      detailHtml = r.calc.detalhe.map(d => `<div class="detail-line"><span>${d.nome}</span><b>${fmtBRL(d.valor)}</b></div>`).join('');
      detailHtml += `<div class="detail-line" style="margin-top:6px;border-top:1px solid var(--border);padding-top:6px"><span>Classe cidade / Peso de cálculo</span><b>${r.calc.classe} / ${fmtNum(r.calc.pesoCalculo)} kg</b></div>`;
    }
    detailTr.innerHTML = `<td colspan="12">${detailHtml}</td>`;
    tbody.appendChild(detailTr);
  });

  tbody.querySelectorAll('.toggle-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const idx = btn.dataset.idx;
      const row = document.getElementById(`detail-${idx}`);
      const isOpen = row.classList.toggle('open');
      btn.textContent = isOpen ? '▾ ocultar' : '▸ ver';
    });
  });
}

function applyConciliacaoFilters() {
  const transp = document.getElementById('filtro-transportadora').value;
  const status = document.getElementById('filtro-status').value;
  const busca = document.getElementById('filtro-busca').value.toLowerCase();

  let rows = state.conciliacaoRows;
  if (transp) rows = rows.filter(r => r.transportadora === transp);
  if (status) rows = rows.filter(r => r.status === status);
  if (busca) rows = rows.filter(r =>
    (r.pedido || '').toLowerCase().includes(busca) ||
    (r.nf || '').toLowerCase().includes(busca) ||
    (r.cidadeRaw || '').toLowerCase().includes(busca)
  );
  renderConciliacaoSummary(rows);
  renderConciliacaoTable(rows);
}

['filtro-transportadora', 'filtro-status'].forEach(id =>
  document.getElementById(id).addEventListener('change', applyConciliacaoFilters)
);
document.getElementById('filtro-busca').addEventListener('input', applyConciliacaoFilters);

document.getElementById('icms-aliquota').addEventListener('change', () => {
  if (state.parsed.relatorio) rodarConciliacao();
});

document.getElementById('btn-export-conciliacao').addEventListener('click', () => {
  const rows = state.conciliacaoRows;
  const data = rows.map(r => ({
    Pedido: r.pedido, NF: r.nf, Cidade: r.cidadeRaw, Transportadora: r.transportadora,
    'Tipo Frete': r.tipoFrete, 'Peso (kg)': r.peso, 'Cubagem (m³)': r.cubagem,
    'Frete Cobrado': r.freteCobrado,
    'Frete Calculado (sem ICMS)': r.calc.erro ? null : r.calc.valorSemICMS,
    'ICMS': r.calc.erro ? null : r.calc.valorICMS,
    'Frete Calculado (com ICMS)': r.calc.erro ? null : r.calc.valorTotal,
    'Diferença': r.diff, 'Diferença %': r.diffPct, Status: r.status,
    Observação: r.calc.erro || '',
  }));
  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Conciliação');
  XLSX.writeFile(wb, `conciliacao_frete_${new Date().toISOString().slice(0,10)}.xlsx`);
});

// ============ COTAÇÃO MANUAL ============
document.getElementById('btn-cotar').addEventListener('click', () => {
  const cidade = document.getElementById('q-cidade').value;
  const uf = document.getElementById('q-uf').value;
  const peso = brNum(document.getElementById('q-peso').value);
  const cubagem = brNum(document.getElementById('q-cubagem').value);
  const valorMercadoria = brNum(document.getElementById('q-valor').value);
  const tipoFrete = document.getElementById('q-tipo').value;
  const freteOriginal = brNum(document.getElementById('q-frete-original').value) || null;

  const calc = calcularComContexto({ cidade, uf, peso, cubagem, valorMercadoria, tipoFrete, freteOriginal });
  const box = document.getElementById('cotacao-resultado');
  box.style.display = 'block';

  if (calc.erro) {
    box.innerHTML = `<div style="color:var(--err)">⚠ ${calc.erro}</div>`;
    return;
  }
  box.innerHTML = `
    <div class="total">${fmtBRL(calc.valorTotal)}</div>
    <div class="note">Classe: ${calc.classe} · Peso de cálculo: ${fmtNum(calc.pesoCalculo)} kg (cubado: ${fmtNum(calc.pesoCubado)} kg)</div>
    <div style="margin-top:10px">
      ${calc.detalhe.map(d => `<div class="detail-line"><span>${d.nome}</span><b>${fmtBRL(d.valor)}</b></div>`).join('')}
    </div>
  `;
});

// ============ COTAÇÃO EM LOTE ============
let loteData = null;
let loteResultRows = [];

document.getElementById('file-cotacao-lote').addEventListener('change', async (e) => {
  const file = e.target.files[0];
  if (!file) return;
  const statusEl = document.getElementById('lote-status');
  try {
    let rows;
    if (file.name.toLowerCase().endsWith('.csv')) {
      const text = await readFileAsText(file, 'UTF-8');
      rows = text.split('\n').filter(l => l.trim()).map(l => l.split(/[;,]/).map(c => c.trim()));
    } else {
      const buf = await readFileAsArrayBuffer(file);
      const wb = XLSX.read(buf, { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    }
    const header = rows[0].map(h => String(h).toLowerCase().trim());
    const idx = {
      cidade: header.findIndex(h => h.includes('cidade')),
      uf: header.findIndex(h => h === 'uf'),
      peso: header.findIndex(h => h.includes('peso')),
      cubagem: header.findIndex(h => h.includes('cubagem')),
      valor: header.findIndex(h => h.includes('valor')),
      tipo: header.findIndex(h => h.includes('tipo')),
    };
    loteData = rows.slice(1).filter(r => r.length && r[idx.cidade]).map(r => ({
      cidade: r[idx.cidade], uf: r[idx.uf] || '', peso: brNum(r[idx.peso]),
      cubagem: brNum(r[idx.cubagem]), valorMercadoria: brNum(r[idx.valor]),
      tipoFrete: r[idx.tipo] || 'Entrega Normal',
    }));
    statusEl.textContent = `${loteData.length} linhas carregadas.`;
    document.getElementById('btn-cotar-lote').disabled = false;
  } catch (err) {
    statusEl.textContent = 'Erro ao ler arquivo: ' + err.message;
  }
});

document.getElementById('btn-cotar-lote').addEventListener('click', () => {
  if (!loteData) return;
  loteResultRows = loteData.map(p => {
    const calc = calcularComContexto(p);
    return { ...p, calc };
  });
  const tbody = document.getElementById('tbody-cotacao-lote');
  tbody.innerHTML = loteResultRows.map(r => `
    <tr>
      <td>${r.cidade} / ${r.uf}</td>
      <td class="num">${fmtNum(r.peso)}</td>
      <td class="num">${fmtNum(r.cubagem)}</td>
      <td class="num">${r.calc.erro ? '—' : fmtBRL(r.calc.valorTotal)}</td>
      <td>${r.calc.erro ? `<span class="badge err">${r.calc.erro}</span>` : '<span class="badge ok">OK</span>'}</td>
    </tr>
  `).join('');
  document.getElementById('lote-resultado-wrap').style.display = 'block';
  document.getElementById('btn-export-cotacao').style.display = 'inline-block';
});

document.getElementById('btn-export-cotacao').addEventListener('click', () => {
  const data = loteResultRows.map(r => ({
    Cidade: r.cidade, UF: r.uf, 'Peso (kg)': r.peso, 'Cubagem (m³)': r.cubagem,
    'Valor Mercadoria': r.valorMercadoria, 'Tipo Frete': r.tipoFrete,
    'Frete Calculado': r.calc.erro ? null : r.calc.valorTotal, Observação: r.calc.erro || '',
  }));
  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Cotação');
  XLSX.writeFile(wb, `cotacao_frete_${new Date().toISOString().slice(0,10)}.xlsx`);
});

// Exposto para depuração no console do navegador (window.state)
if (typeof window !== 'undefined') window.state = state;
