if(typeof pdfjsLib !== 'undefined'){
  pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
}

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
    if(theme === 'dark'){
      icon.textContent = '☀';
      label.textContent = 'Modo claro';
    } else {
      icon.textContent = '☾';
      label.textContent = 'Modo escuro';
    }
    try{ localStorage.setItem('composicao-preco-theme', theme); }catch(e){}
  }

  let saved = 'light';
  try{ saved = localStorage.getItem('composicao-preco-theme') || 'light'; }catch(e){}
  apply(saved);

  btn.addEventListener('click', () => {
    const current = root.getAttribute('data-theme') === 'dark' ? 'light' : 'dark';
    apply(current);
  });
})();

/* ---------------------------------------------------------------------
   HELPERS
--------------------------------------------------------------------- */
function num(id){ return parseFloat(document.getElementById(id).value) || 0; }
function fmtBRL(n){ return (n || 0).toLocaleString('pt-BR', {style:'currency', currency:'BRL'}); }
function escapeHtml(str){
  return String(str == null ? '' : str).replace(/[&<>"']/g, s => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[s]));
}
function freteAtivo(){
  return !document.getElementById('frete-box').classList.contains('hidden');
}

/* ---------------------------------------------------------------------
   ITENS — todos usam as mesmas alíquotas, definidas na seção 2
--------------------------------------------------------------------- */
let items = [];
let itemCounter = 0;

function calcItemValues(item, totalQtd){
  const custo = item.custo;
  const qtd = item.qtd;
  const ipiInF = num('ipi-in') / 100;
  const icmsInF = num('icms-in') / 100;
  const pisCofinsInF = num('pis-cofins-in') / 100;
  const ipiOutF = num('ipi-out') / 100;
  const icmsOutF = num('icms-out') / 100;
  const pisCofinsOutF = num('pis-cofins-out') / 100;
  const margemF = num('margem-out') / 100;

  const freteTotal = freteAtivo() ? num('frete-valor') : 0;
  const servicoUnit = (freteTotal > 0 && totalQtd > 0) ? freteTotal / totalQtd : 0;

  /* impostos de entrada — em cascata, "por dentro" para o IPI de entrada */
  const ipiEntradaValor = custo / (1 + ipiInF) * ipiInF;
  const icmsEntradaValor = (custo - ipiEntradaValor) * icmsInF;
  const pisCofinsEntradaValor = (custo - ipiEntradaValor - icmsEntradaValor) * pisCofinsInF;

  const cmv = servicoUnit + (custo - ipiEntradaValor - icmsEntradaValor - pisCofinsEntradaValor);

  /* impostos de saída — um único divisor que já embute ICMS, o efeito do
     IPI sobre o ICMS e PIS/COFINS de saída */
  const denomSaida = 1 - icmsOutF - (icmsOutF * ipiOutF) - pisCofinsOutF;
  const denomMargem = 1 - margemF;

  let precoFinal = 0, warn = false;
  if(denomMargem <= 0 || denomSaida <= 0){
    warn = true;
  } else {
    const recLiq = cmv / denomMargem;
    precoFinal = recLiq / denomSaida;
  }
  const total = precoFinal * (parseFloat(qtd) || 0);
  return {cmv, precoFinal, total, warn};
}

function recalcAll(){
  const totalQtd = items.reduce((s, it) => s + (parseFloat(it.qtd) || 0), 0);
  let grandTotal = 0;
  let anyWarn = false;

  items.forEach(it => {
    const r = calcItemValues(it, totalQtd);
    if(r.warn) anyWarn = true;
    grandTotal += r.total;
    it._calc = r;

    const row = document.querySelector('tr[data-row-id="' + it.id + '"]');
    if(row){
      row.querySelector('.c-cmv').textContent = fmtBRL(r.cmv);
      row.querySelector('.c-preco-final').textContent = fmtBRL(r.precoFinal);
      row.querySelector('.c-total').textContent = fmtBRL(r.total);
    }
  });

  document.getElementById('grand-total').textContent = fmtBRL(grandTotal);

  const warnEl = document.getElementById('pricing-warning');
  if(anyWarn){
    warnEl.style.display = 'block';
    warnEl.textContent = 'Os percentuais informados resultam em divisor zero ou negativo (ICMS saída + IPI, PIS/COFINS saída ou margem estão altos demais). Ajuste os percentuais — o preço final não pode ser calculado assim.';
  } else {
    warnEl.style.display = 'none';
  }

  renderOutputTable(grandTotal);
}

function addItem(pn, descricao, ncm, custo, qtd){
  pn = pn || '';
  descricao = descricao || '';
  ncm = ncm || '';
  custo = custo || 0;
  qtd = (qtd === undefined || qtd === null) ? 1 : qtd;

  itemCounter++;
  const id = 'i' + itemCounter;
  const item = {id: id, pn: pn, descricao: descricao, ncm: ncm, custo: custo, qtd: qtd};
  items.push(item);

  const tr = document.createElement('tr');
  tr.setAttribute('data-row-id', id);
  tr.innerHTML =
    '<td><input type="text" class="cell-pn" value="' + escapeHtml(pn) + '" placeholder="PN"></td>' +
    '<td><input type="text" class="cell-descricao" value="' + escapeHtml(descricao) + '" placeholder="Descrição"></td>' +
    '<td><input type="text" class="cell-ncm" value="' + escapeHtml(ncm) + '" placeholder="0000.00.00"></td>' +
    '<td><input type="number" class="cell-custo" step="0.01" value="' + custo + '"></td>' +
    '<td><input type="number" class="cell-qtd" step="1" min="0" value="' + qtd + '"></td>' +
    '<td class="calc c-cmv">R$ 0,00</td>' +
    '<td class="calc c-preco-final strong">R$ 0,00</td>' +
    '<td class="calc c-total strong">R$ 0,00</td>' +
    '<td><button type="button" class="row-remove" aria-label="Remover item">✕</button></td>';

  document.getElementById('items-tbody').appendChild(tr);

  tr.querySelector('.cell-pn').addEventListener('input', (e) => { item.pn = e.target.value; renderOutputTable(); });
  tr.querySelector('.cell-descricao').addEventListener('input', (e) => { item.descricao = e.target.value; renderOutputTable(); });
  tr.querySelector('.cell-ncm').addEventListener('input', (e) => { item.ncm = e.target.value; renderOutputTable(); });
  tr.querySelector('.cell-custo').addEventListener('input', (e) => {
    item.custo = parseFloat(e.target.value) || 0;
    recalcAll();
  });
  tr.querySelector('.cell-qtd').addEventListener('input', (e) => {
    item.qtd = parseFloat(e.target.value) || 0;
    recalcAll();
  });
  tr.querySelector('.row-remove').addEventListener('click', () => {
    items = items.filter(i => i.id !== id);
    tr.remove();
    recalcAll();
  });

  recalcAll();
  return id;
}

document.getElementById('add-item').addEventListener('click', () => addItem('', '', '', 0, 1));

const rateIds = ['icms-in','pis-cofins-in','ipi-in','ipi-out','icms-out','pis-cofins-out','iss-out','margem-out','frete-valor'];
rateIds.forEach(id => document.getElementById(id).addEventListener('input', recalcAll));

document.getElementById('frete-toggle').addEventListener('click', () => {
  const box = document.getElementById('frete-box');
  const btn = document.getElementById('frete-toggle');
  const willShow = box.classList.contains('hidden');
  box.classList.toggle('hidden', !willShow);
  btn.classList.toggle('active', willShow);
  btn.textContent = willShow ? '− Remover frete do preço' : '+ Considerar frete no preço';
  recalcAll();
});

/* linha inicial */
addItem('', '', '', 0, 1);

/* ---------------------------------------------------------------------
   TABELA DE SAÍDA — mesma estrutura do modelo de referência:
   Item | Modelo (PN + descrição) | Qtd | NCM | Preço Unitário | Preço Total
--------------------------------------------------------------------- */
function renderOutputTable(grandTotalArg){
  const tbody = document.getElementById('output-tbody');
  const totalQtd = items.reduce((s, it) => s + (parseFloat(it.qtd) || 0), 0);
  let grandTotal = 0;

  const rowsHtml = items.map((it, idx) => {
    const r = it._calc || calcItemValues(it, totalQtd);
    grandTotal += r.total;
    return '<tr>' +
      '<td>' + (idx + 1) + '</td>' +
      '<td><span class="out-pn">' + escapeHtml(it.pn || '—') + '</span>' +
        (it.descricao ? '<span class="out-desc">' + escapeHtml(it.descricao) + '</span>' : '') +
      '</td>' +
      '<td class="center">' + (it.qtd || 0) + '</td>' +
      '<td class="center">' + escapeHtml(it.ncm || '—') + '</td>' +
      '<td class="num">' + fmtBRL(r.precoFinal) + '</td>' +
      '<td class="num">' + fmtBRL(r.total) + '</td>' +
    '</tr>';
  }).join('');

  tbody.innerHTML = rowsHtml || '<tr><td colspan="6">Nenhum item adicionado.</td></tr>';
  document.getElementById('output-total').innerHTML = '<strong>' + fmtBRL(grandTotalArg !== undefined ? grandTotalArg : grandTotal) + '</strong>';
}

/* ---------------------------------------------------------------------
   COPIAR TABELA — seleciona e copia mantendo a formatação (cola no Word)
--------------------------------------------------------------------- */
document.getElementById('copy-table').addEventListener('click', () => {
  const statusEl = document.getElementById('copy-status');
  const table = document.getElementById('output-table');
  try{
    const range = document.createRange();
    range.selectNode(table);
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
    const ok = document.execCommand('copy');
    sel.removeAllRanges();
    statusEl.textContent = ok ? 'Copiado! Cole com Ctrl+V no Word.' : 'Não copiou automaticamente — selecione a tabela manualmente e use Ctrl+C.';
  } catch(err){
    statusEl.textContent = 'Não copiou automaticamente — selecione a tabela manualmente e use Ctrl+C.';
  }
  setTimeout(() => { statusEl.textContent = ''; }, 4000);
});

/* ---------------------------------------------------------------------
   EXTRAÇÃO DO PDF (opcional)
   O layout do fornecedor pode variar, então isso é só um ponto de
   partida — confira e ajuste os campos manualmente. Ajuste os regex em
   PATTERNS conforme o modelo real do seu fornecedor para aumentar a
   precisão.
--------------------------------------------------------------------- */
const PATTERNS = {
  precoCompra: /(?:valor unit[aá]rio|pre[cç]o unit[aá]rio|valor total|pre[cç]o)\s*[:\-]?\s*R?\$?\s*([\d.,]+)/i,
  ncm: /NCM\s*[:\-]?\s*(\d{4}\.?\d{2}\.?\d{2})/i
};

function parseCurrencyToNumber(str){
  if(!str) return null;
  const cleaned = str.replace(/[^\d,.-]/g,'').replace(/\.(?=\d{3},)/g,'').replace(',', '.');
  const n = parseFloat(cleaned);
  return isNaN(n) ? null : n;
}

document.getElementById('pdf-input').addEventListener('change', async (e) => {
  const file = e.target.files[0];
  const statusEl = document.getElementById('upload-status');
  if(!file) return;
  statusEl.className = '';
  statusEl.textContent = 'Lendo PDF...';

  if(typeof pdfjsLib === 'undefined'){
    statusEl.className = 'err';
    statusEl.textContent = 'A biblioteca de leitura de PDF não carregou (verifique sua conexão ou se algo está bloqueando cdnjs.cloudflare.com). Preencha os campos manualmente.';
    return;
  }

  try{
    const buf = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({data: buf}).promise;
    let fullText = '';
    for(let i=1; i<=pdf.numPages; i++){
      const page = await pdf.getPage(i);
      const content = await page.getTextContent();
      fullText += content.items.map(it => it.str).join(' ') + '\n';
    }

    const cMatch = fullText.match(PATTERNS.precoCompra);
    const nMatch = fullText.match(PATTERNS.ncm);
    let found = [];

    if(cMatch && items.length){
      const val = parseCurrencyToNumber(cMatch[1]);
      if(val !== null){
        items[0].custo = val;
        const row = document.querySelector('tr[data-row-id="' + items[0].id + '"]');
        if(row) row.querySelector('.cell-custo').value = val.toFixed(2);
        found.push('preço do 1º item');
      }
    }
    if(nMatch && items.length){
      items[0].ncm = nMatch[1];
      const row = document.querySelector('tr[data-row-id="' + items[0].id + '"]');
      if(row) row.querySelector('.cell-ncm').value = nMatch[1];
      found.push('NCM do 1º item');
    }
    recalcAll();

    if(found.length){
      statusEl.className = 'ok';
      statusEl.textContent = 'Extraído automaticamente: ' + found.join(', ') + '. Confira os campos antes de prosseguir.';
    } else {
      statusEl.className = 'err';
      statusEl.textContent = 'Não foi possível reconhecer os campos automaticamente neste layout. Preencha manualmente.';
    }
  } catch(err){
    statusEl.className = 'err';
    statusEl.textContent = 'Erro ao ler o PDF: ' + err.message;
  }
});
