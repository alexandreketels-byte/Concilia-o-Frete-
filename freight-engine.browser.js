// Motor de cálculo de frete - genérico, alimentado pelos dados carregados via upload
// (table e general vêm da planilha de tarifa real; constantes abaixo são apenas fallback)

function normalizeKey(str) {
  return (str || '')
    .toString()
    .trim()
    .toUpperCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}

function parseBRNumber(val) {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') return val;
  let s = val.toString().trim();
  s = s.replace(/\./g, '').replace(',', '.');
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

// ---- Fallbacks (usados só se a tabela de tarifa não puder ser lida do arquivo) ----
const RODONAVES_TABLE_FALLBACK = {
  'SP CAPITAL': [19.37861104306209, 24.556889, 29.69438623000033, 36.57409, 52.81341248, 0.593918437],
  'SP INTERIOR': [27.49923722, 34.57744325, 42.22516011, 49.76827294, 71.260829, 0.794990628],
};
const GENERAL_RULES_FALLBACK = {
  cubagemFatorKgM3: 200,
  freteValorPct: 0.003, freteValorMinimo: 3.549,
  freteValorPctNorte: 0.004, freteValorMinimoNorte: 7.848,
  taxaDespacho: 23.44,
  taxaDespachoUFs: ['RJ', 'ES', 'MT', 'MS', 'AM', 'PA', 'RO', 'AC'],
  tasValor: 5.427467,
  tasUFs: ['MG', 'GO', 'DF', 'MT', 'MS', 'RO', 'PA', 'AC'],
  tasSuframaValor: 10.705651999999999,
  tasSuframaUFs: ['PA', 'RR', 'RO', 'AP', 'TO', 'AM'],
  grisMinimo: 1.36, grisPct: 0.001,
  grisMinimoRegiao2: 1.15, grisPctRegiao2: 0.001,
  grisRegiao2UFs: ['RJ', 'ES', 'MT', 'MS', 'AC', 'RO'],
  grisMinimoRegiao3: 2.47, grisPctRegiao3: 0.003,
  grisRegiao3UFs: ['PA', 'RR', 'AP', 'TO', 'AM'],
  pedagioPorFracao100kg: 5.427467,
  pedagioRegiaoNorteValor: 8.786312,
  pedagioRegiaoNorteUFs: ['PA', 'RR', 'AP', 'TO', 'AM'],
  reentregaPct: 0.5,
  devolucaoPct: 1.0,
  icmsAliquota: 0.12,
};

function getTableKey(uf, classe) {
  const ufN = normalizeKey(uf);
  const single = ['AP', 'RR', 'AM'];
  if (single.includes(ufN)) return `${ufN} ESTADO`;
  if (ufN === 'PA' && normalizeKey(classe).includes('FLUVIAL')) return 'NORTE FLUVIAL';
  if (ufN === 'GO' || ufN === 'DF') {
    return normalizeKey(classe).includes('INTERIOR') ? 'GO INTERIOR' : 'GO E DF CAPITAL';
  }
  const isCapital = normalizeKey(classe).includes('CAPITAL');
  return `${ufN} ${isCapital ? 'CAPITAL' : 'INTERIOR'}`;
}

function calcFretePeso(pesoCalculo, uf, classe, table) {
  const key = getTableKey(uf, classe);
  const values = table[key];
  if (!values) {
    return { valor: null, erro: `Faixa de tarifa não encontrada para "${key}" (UF=${uf}, classe=${classe}). Verifique se essa combinação existe na tabela de tarifa.`, faixa: null };
  }
  const [v0_10, v11_20, v21_40, v41_60, v61_100, vAcima100kg] = values;

  if (pesoCalculo <= 10) return { valor: v0_10, faixa: '0 - 10 kg' };
  if (pesoCalculo <= 20) return { valor: v11_20, faixa: '11 - 20 kg' };
  if (pesoCalculo <= 40) return { valor: v21_40, faixa: '21 - 40 kg' };
  if (pesoCalculo <= 60) return { valor: v41_60, faixa: '41 - 60 kg' };
  if (pesoCalculo <= 100) return { valor: v61_100, faixa: '61 - 100 kg' };

  const excedente = pesoCalculo - 100;
  const valor = v61_100 + excedente * vAcima100kg;
  return { valor, faixa: `acima de 100 kg (61-100kg + ${excedente.toFixed(1)}kg × R$${vAcima100kg.toFixed(4)}/kg)` };
}

function calcAcessorios(params, general) {
  const { uf, valorMercadoria, pesoCalculo, isSPCapital, secCatRows, tipoFrete, freteOriginal } = params;
  const ufN = normalizeKey(uf);
  const acessorios = [];

  const isNorte = (general.pedagioRegiaoNorteUFs || []).includes(ufN);
  const pctFV = isNorte ? general.freteValorPctNorte : general.freteValorPct;
  const minFV = isNorte ? general.freteValorMinimoNorte : general.freteValorMinimo;
  const valorFV = Math.max((valorMercadoria || 0) * pctFV, minFV);
  acessorios.push({ nome: 'Frete Valor', valor: valorFV });

  if ((general.taxaDespachoUFs || []).includes(ufN)) {
    acessorios.push({ nome: 'Taxa de Despacho', valor: general.taxaDespacho });
  }

  if ((general.tasUFs || []).includes(ufN)) {
    acessorios.push({ nome: 'TAS (Taxa Administrativa Sefaz)', valor: general.tasValor });
  }
  if ((general.tasSuframaUFs || []).includes(ufN)) {
    acessorios.push({ nome: 'TAS Suframa', valor: general.tasSuframaValor });
  }

  let grisPct = general.grisPct, grisMin = general.grisMinimo;
  if ((general.grisRegiao3UFs || []).includes(ufN)) {
    grisPct = general.grisPctRegiao3; grisMin = general.grisMinimoRegiao3;
  } else if ((general.grisRegiao2UFs || []).includes(ufN)) {
    grisPct = general.grisPctRegiao2; grisMin = general.grisMinimoRegiao2;
  }
  const valorGris = Math.max((valorMercadoria || 0) * grisPct, grisMin);
  acessorios.push({ nome: 'GRIS', valor: valorGris });

  if (!isSPCapital) {
    const fracoes = Math.ceil(pesoCalculo / 100);
    const valorPorFracao = (general.pedagioRegiaoNorteUFs || []).includes(ufN)
      ? general.pedagioRegiaoNorteValor
      : general.pedagioPorFracao100kg;
    acessorios.push({ nome: 'Pedágio', valor: fracoes * valorPorFracao });
  }

  if (secCatRows && secCatRows.length > 0) {
    for (const row of secCatRows) {
      acessorios.push({ nome: `Taxa local: ${row.nomeTarifa}`, valor: row.valor });
    }
  }

  if (tipoFrete && normalizeKey(tipoFrete) === 'DEVOLUCAO' && freteOriginal) {
    acessorios.push({ nome: 'Devolução (100% frete original)', valor: freteOriginal * general.devolucaoPct, substitui: true });
  }
  if (tipoFrete && normalizeKey(tipoFrete) === 'REENTREGA' && freteOriginal) {
    acessorios.push({ nome: 'Reentrega (50% frete original)', valor: Math.max(freteOriginal * general.reentregaPct, 15), substitui: true });
  }

  return acessorios;
}

// pedido: {cidade, uf, peso, cubagem, valorMercadoria, tipoFrete, freteOriginal}
// cityLookup: { CHAVE: {polo, ...} }
// secCatLookup: { CHAVE: [{nomeTarifa, valor}] }
// ctx: { table, general } — opcional, usa fallback se ausente
function calcularFreteTeorico(pedido, cityLookup, secCatLookup, ctx) {
  const table = (ctx && ctx.table && Object.keys(ctx.table).length) ? ctx.table : RODONAVES_TABLE_FALLBACK;
  const general = (ctx && ctx.general) ? ctx.general : GENERAL_RULES_FALLBACK;

  const { cidade, uf, peso, cubagem, valorMercadoria, tipoFrete, freteOriginal } = pedido;
  const cityKey = normalizeKey(cidade) + normalizeKey(uf);
  const cityInfo = cityLookup ? cityLookup[cityKey] : null;
  const classe = cityInfo ? cityInfo.polo : null;

  if (!classe) {
    return {
      erro: `Cidade "${cidade}/${uf}" não encontrada na relação de cidades`,
      valorTotal: null,
    };
  }

  const pesoCubado = (cubagem || 0) * (general.cubagemFatorKgM3 || 200);
  const pesoCalculo = Math.max(peso || 0, pesoCubado);

  const baseResult = calcFretePeso(pesoCalculo, uf, classe, table);
  if (baseResult.erro) {
    return { erro: baseResult.erro, valorTotal: null };
  }

  const ufN = normalizeKey(uf);
  const isSPCapital = ufN === 'SP' && classe === 'CAPITAL';
  const secCatRows = (secCatLookup && secCatLookup[cityKey]) || [];

  const acessorios = calcAcessorios({
    uf, valorMercadoria, pesoCalculo, isSPCapital, secCatRows, tipoFrete, freteOriginal,
  }, general);

  let valorTotal = baseResult.valor;
  const detalhe = [{ nome: `Frete-peso (${baseResult.faixa})`, valor: valorTotal }];

  for (const a of acessorios) {
    if (a.substitui) {
      valorTotal = a.valor;
      detalhe.length = 0;
      detalhe.push({ nome: a.nome, valor: a.valor });
    } else {
      valorTotal += a.valor;
      detalhe.push(a);
    }
  }

  // ICMS por fora (gross-up): o frete cobrado pela transportadora já inclui ICMS,
  // então para comparar com o mesmo critério aplicamos: valorComICMS = valorSemICMS / (1 - alíquota)
  const aliquotaICMS = (general.icmsAliquota !== undefined) ? general.icmsAliquota : 0.12;
  const valorSemICMS = valorTotal;
  const valorICMS = aliquotaICMS > 0 ? (valorSemICMS / (1 - aliquotaICMS)) - valorSemICMS : 0;
  valorTotal = valorSemICMS + valorICMS;
  if (valorICMS > 0) {
    detalhe.push({ nome: `ICMS (${(aliquotaICMS * 100).toFixed(0)}% por fora)`, valor: valorICMS });
  }

  return { valorTotal, valorSemICMS, valorICMS, detalhe, classe, pesoCalculo, pesoCubado, erro: null };
}
