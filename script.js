/* ---------------------------------------------------------------------
   MODO CLARO / ESCURO
--------------------------------------------------------------------- */
(function initTheme(){
  const root = document.documentElement;
  const btn = document.getElementById('theme-toggle');
  const icon = document.getElementById('theme-icon');
  const label = document.getElementById('theme-label');

  function apply(theme){
    root.setAttribute('data-theme', theme);
    icon.textContent = theme === 'dark' ? '☀' : '☾';
    label.textContent = theme === 'dark' ? 'Modo claro' : 'Modo escuro';
    try{ localStorage.setItem('composicao-preco-theme', theme); }catch(e){}
  }

  let saved = 'light';
  try{ saved = localStorage.getItem('composicao-preco-theme') || 'light'; }catch(e){}
  apply(saved);

  btn.addEventListener('click', () => {
    apply(root.getAttribute('data-theme') === 'dark' ? 'light' : 'dark');
  });
})();

/* ---------------------------------------------------------------------
   HELPERS
--------------------------------------------------------------------- */
function fmtBRL(n){ return (n || 0).toLocaleString('pt-BR', {style:'currency', currency:'BRL'}); }
function fmtPct(n){ return (n || 0).toLocaleString('pt-BR', {maximumFractionDigits: 2}) + '%'; }
function escapeHtml(str){
  return String(str == null ? '' : str).replace(/[&<>"']/g, s => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[s]));
}

/* ---------------------------------------------------------------------
   LEITURA DA PLANILHA ("Ver. 3.0 - Planilha de Cálculo de Revenda ou
   Distribuição"). O layout é fixo; as colunas abaixo seguem esse modelo.
   O site NÃO recalcula: lê os valores já calculados (em cache) pelo Excel.
--------------------------------------------------------------------- */
const SHEET_NAME = 'Revenda';
const HEADER_ROW = 13;
const FIRST_ROW = 14;
const LAST_ROW = 113;
const COL = {
  item: 'A',      // código do item
  desc: 'E',      // descrição
  ncm: 'G',       // NCM
  qtd: 'P',       // quantidade
  icms: 'R',      // ICMS de saída (%)
  pisCofins: 'S', // PIS/COFINS de saída (%)
  ipi: 'T',       // IPI de saída (%)
  unitDireta: 'W',  // preço unitário - revenda direta (receita bruta s/ IPI)
  totalDireta: 'X', // preço total - revenda direta
  unitPortal: 'Z',  // preço unitário - portal (líquido de ICMS e IPI)
  totalPortal: 'AA' // preço total - portal
};
const CELL_SERVICOS = 'E9';

let items = [];
let servicos = null;

/* células com erro (#REF!, #VALUE!...) são tratadas como vazias */
function cellAt(ws, addr){
  const c = ws[addr];
  return (c && c.t !== 'e') ? c : null;
}
function numAt(ws, addr){
  const c = cellAt(ws, addr);
  return (c && typeof c.v === 'number' && isFinite(c.v)) ? c.v : null;
}
function strAt(ws, addr){
  const c = cellAt(ws, addr);
  return (c && c.v != null) ? String(c.v).trim() : '';
}

/* procura uma coluna de PN nas linhas de cabeçalho (12 e 13) */
function findPnColumn(ws){
  const re = /^(pn|p\/n|part\s*number|partnumber|n[ºo°]?\s*da\s*pe[çc]a|n[úu]mero\s*da\s*pe[çc]a)$/i;
  for(const row of [HEADER_ROW - 1, HEADER_ROW]){
    for(let c = 0; c < 60; c++){
      const addr = XLSX.utils.encode_cell({r: row - 1, c: c});
      if(re.test(strAt(ws, addr))) return XLSX.utils.encode_col(c);
    }
  }
  return null;
}

/* acha a aba "Revenda" pelo nome (ignora maiúsculas, acentos e espaços nas pontas) */
function normalizeName(n){
  return String(n).normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toLowerCase();
}
function findSheetName(names){
  const alvo = normalizeName(SHEET_NAME);
  return names.find(n => normalizeName(n) === alvo) || null;
}

function parseWorkbook(wb, sheetName){
  const warnings = [];
  const ws = wb.Sheets[sheetName];

  const okLayout = strAt(ws, COL.item + HEADER_ROW).toUpperCase() === 'ITEM'
    && /TOTAL/i.test(strAt(ws, COL.totalDireta + HEADER_ROW));
  if(!okLayout) warnings.push('O cabeçalho da linha ' + HEADER_ROW + ' não parece o da planilha Ver. 3.0. Confira se as colunas leem os valores certos.');

  const pnCol = findPnColumn(ws);
  const list = [];
  const skipped = [];

  for(let r = FIRST_ROW; r <= LAST_ROW; r++){
    const code = strAt(ws, COL.item + r);
    if(code === '') continue;

    const qtd = numAt(ws, COL.qtd + r);
    const prices = {
      unitDireta: numAt(ws, COL.unitDireta + r),
      totalDireta: numAt(ws, COL.totalDireta + r),
      unitPortal: numAt(ws, COL.unitPortal + r),
      totalPortal: numAt(ws, COL.totalPortal + r)
    };
    if(!(qtd > 0)){ skipped.push('Item ' + code + ' (linha ' + r + '): quantidade vazia.'); continue; }
    if(Object.values(prices).some(v => v === null)){
      skipped.push('Item ' + code + ' (linha ' + r + '): preço sem valor calculado (confira se a planilha tem erro, como #REF!).');
      continue;
    }

    list.push({
      code: code,
      pn: pnCol ? strAt(ws, pnCol + r) : '',
      desc: strAt(ws, COL.desc + r),
      ncm: strAt(ws, COL.ncm + r),
      qtd: qtd,
      icms: (numAt(ws, COL.icms + r) || 0) * 100,
      pisCofins: (numAt(ws, COL.pisCofins + r) || 0) * 100,
      ipi: (numAt(ws, COL.ipi + r) || 0) * 100,
      unitDireta: prices.unitDireta,
      totalDireta: prices.totalDireta,
      unitPortal: prices.unitPortal,
      totalPortal: prices.totalPortal
    });
  }

  if(!pnCol) warnings.push('Não encontrei uma coluna "PN" no cabeçalho; a descrição será usada como identificação do item.');
  skipped.forEach(s => warnings.push(s));

  /* conferência: soma lida x subtotal da própria planilha (X11) */
  const subtotal = numAt(ws, COL.totalDireta + '11');
  if(skipped.length === 0 && subtotal !== null){
    const soma = list.reduce((s, it) => s + it.totalDireta, 0);
    if(Math.abs(soma - subtotal) > 0.01) warnings.push('A soma dos itens (' + fmtBRL(soma) + ') difere do subtotal da planilha (' + fmtBRL(subtotal) + ').');
  }

  return {
    list: list,
    warnings: warnings,
    servicos: numAt(ws, CELL_SERVICOS),
    proposta: strAt(ws, 'E2'),
    cliente: strAt(ws, 'E3'),
    pnCol: pnCol
  };
}

document.getElementById('xlsx-input').addEventListener('change', async (e) => {
  const file = e.target.files[0];
  const statusEl = document.getElementById('upload-status');
  const warnEl = document.getElementById('upload-warnings');
  if(!file) return;
  statusEl.className = '';
  statusEl.textContent = 'Lendo planilha...';
  warnEl.innerHTML = '';

  if(typeof XLSX === 'undefined'){
    statusEl.className = 'err';
    statusEl.textContent = 'A biblioteca de leitura de Excel não carregou (verifique a conexão ou se algo bloqueia cdnjs.cloudflare.com).';
    return;
  }

  try{
    const buf = await file.arrayBuffer();

    /* 1ª leitura: só os nomes das abas (rápida, mesmo com abas grandes) */
    const nomes = XLSX.read(buf, {type: 'array', bookSheets: true}).SheetNames;
    const alvo = findSheetName(nomes);
    if(!alvo){
      items = [];
      servicos = null;
      statusEl.className = 'err';
      statusEl.textContent = 'Não encontrei a aba "' + SHEET_NAME + '" neste arquivo. Abas encontradas: ' + nomes.join(', ') + '.';
      renderOutput();
      return;
    }

    /* 2ª leitura: só a aba Revenda */
    const wb = XLSX.read(buf, {type: 'array', sheets: alvo});
    const res = parseWorkbook(wb, alvo);
    items = res.list;
    servicos = res.servicos;

    const partes = ['aba "' + alvo + '"', items.length + (items.length === 1 ? ' item lido' : ' itens lidos')];
    if(res.proposta) partes.push('proposta ' + res.proposta);
    if(res.cliente) partes.push('cliente ' + res.cliente);
    statusEl.className = items.length ? 'ok' : 'err';
    const resumo = partes.join(' · ');
    statusEl.textContent = items.length ? 'Planilha lida: ' + resumo + (resumo.endsWith('.') ? '' : '.') : 'Nenhum item válido encontrado na planilha.';
    warnEl.innerHTML = res.warnings.map(w => '<li>' + escapeHtml(w) + '</li>').join('');
    renderOutput();
  } catch(err){
    statusEl.className = 'err';
    statusEl.textContent = 'Erro ao ler a planilha: ' + err.message;
  }
});

/* ---------------------------------------------------------------------
   TABELA DE SAÍDA — Item | Modelo | Qtd | NCM | Preço Unitário | Preço Total
--------------------------------------------------------------------- */
function groupByRate(list, key){
  const map = new Map();
  list.forEach((it, idx) => {
    const v = Math.round(it[key] * 10000) / 10000;
    if(!map.has(v)) map.set(v, []);
    map.get(v).push(idx + 1);
  });
  return Array.from(map.entries()).map(([rate, nums]) => ({rate, nums}));
}
function describeRates(list, key){
  const groups = groupByRate(list, key);
  if(groups.length === 1) return fmtPct(groups[0].rate);
  return groups.map(g => fmtPct(g.rate) + ' (' + (g.nums.length > 1 ? 'itens ' : 'item ') + g.nums.join(', ') + ')').join(' e ');
}

function renderOutput(){
  const portal = document.getElementById('opt-preco').value === 'portal';
  const tbody = document.getElementById('output-tbody');
  let total = 0;

  const rowsHtml = items.map((it, idx) => {
    const unit = portal ? it.unitPortal : it.unitDireta;
    const tot = portal ? it.totalPortal : it.totalDireta;
    total += tot;
    const nome = it.pn || it.desc || ('Item ' + it.code);
    const sub = it.pn ? it.desc : '';
    return '<tr>' +
      '<td>' + (idx + 1) + '</td>' +
      '<td class="left"><strong>' + escapeHtml(nome) + '</strong>' + (sub ? '<br>' + escapeHtml(sub) : '') + '</td>' +
      '<td>' + it.qtd + '</td>' +
      '<td>' + escapeHtml(it.ncm || '—') + '</td>' +
      '<td>' + fmtBRL(unit) + '</td>' +
      '<td>' + fmtBRL(tot) + '</td>' +
    '</tr>';
  }).join('');

  tbody.innerHTML = rowsHtml || '<tr><td colspan="6">Escolha a planilha para carregar os itens.</td></tr>';
  document.getElementById('output-total').textContent = fmtBRL(total);

  const freteRow = document.getElementById('output-frete-row');
  if(items.length && servicos !== null && servicos > 0){
    freteRow.classList.remove('hidden');
    document.getElementById('output-frete-valor').textContent = fmtBRL(servicos);
  } else {
    freteRow.classList.add('hidden');
  }

  renderOutputNotes(portal);
}

/* bullets de impostos (alíquotas de saída lidas da planilha), estado e Zona Franca */
function renderOutputNotes(portal){
  const ncms = [];
  items.forEach(it => {
    const n = (it.ncm || '').trim();
    if(n && ncms.indexOf(n) === -1) ncms.push(n);
  });

  const ipiMode = portal ? 'excluso' : document.getElementById('opt-ipi').value;
  const taxLines = [];
  if(items.length){
    taxLines.push('PIS/COFINS: ' + describeRates(items, 'pisCofins') + ' incluso');
    taxLines.push('ICMS: ' + describeRates(items, 'icms') + (portal ? ' excluso (preços líquidos de ICMS)' : ' incluso'));
    taxLines.push('IPI: ' + describeRates(items, 'ipi') + ' ' + ipiMode);
  }
  const taxHtml = taxLines.map(l => '<li>' + escapeHtml(l) + '</li>').join('');

  document.getElementById('output-tax-list').innerHTML = ncms.length
    ? '<li>NCM: ' + escapeHtml(ncms.join(', ')) + '<ul>' + taxHtml + '</ul></li>'
    : taxHtml;

  const estado = document.getElementById('opt-estado').value.trim();
  const estadoItem = document.getElementById('output-estado-item');
  if(estado){
    estadoItem.textContent = 'Estamos considerando que o cliente é contribuinte de ICMS no estado do ' + estado + ';';
    estadoItem.classList.remove('hidden');
  } else {
    estadoItem.classList.add('hidden');
  }

  document.getElementById('output-zf').classList.toggle('hidden', !document.getElementById('opt-zf').checked);
}

['opt-preco', 'opt-estado', 'opt-ipi', 'opt-zf'].forEach(id => {
  const el = document.getElementById(id);
  el.addEventListener('input', renderOutput);
  el.addEventListener('change', renderOutput);
});

/* ---------------------------------------------------------------------
   COPIAR TABELA — seleciona e copia mantendo a formatação (cola no Word)
--------------------------------------------------------------------- */
document.getElementById('copy-table').addEventListener('click', () => {
  const statusEl = document.getElementById('copy-status');
  const wrap = document.getElementById('output-wrap');
  try{
    const range = document.createRange();
    range.selectNodeContents(wrap);
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
    const ok = document.execCommand('copy');
    sel.removeAllRanges();
    statusEl.textContent = ok ? 'Copiado! Cole com Ctrl+V no Word.' : 'Não copiou automaticamente — selecione o conteúdo manualmente e use Ctrl+C.';
  } catch(err){
    statusEl.textContent = 'Não copiou automaticamente — selecione o conteúdo manualmente e use Ctrl+C.';
  }
  setTimeout(() => { statusEl.textContent = ''; }, 4000);
});

renderOutput();
