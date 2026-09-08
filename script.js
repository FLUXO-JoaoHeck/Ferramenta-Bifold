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

/* logo opcional: guarda como dataURL para usar na capa */
let logoDataUrl = null;
document.getElementById('f-logo').addEventListener('change', (e) => {
  const file = e.target.files[0];
  if(!file) { logoDataUrl = null; return; }
  const reader = new FileReader();
  reader.onload = () => { logoDataUrl = reader.result; };
  reader.readAsDataURL(file);
});

/* ---------------------------------------------------------------------
   ITENS DA PROPOSTA
--------------------------------------------------------------------- */
let items = [];
let itemCounter = 0;

function calcItemValues(custo, qtd, totalQtd){
  const ipiF = num('ipi-in') / 100;
  const icmsInF = num('icms-in') / 100;
  const pisCofinsInF = (num('pis-in') + num('cofins-in')) / 100;
  const icmsOutF = num('icms-out') / 100;
  const pisCofinsOutF = (num('pis-out') + num('cofins-out')) / 100;
  const margemF = num('margem-out') / 100;

  const freteTotal = freteAtivo() ? num('frete-valor') : 0;
  const servicoUnit = (freteTotal > 0 && totalQtd > 0) ? freteTotal / totalQtd : 0;

  /* impostos de entrada — em cascata, "por dentro" para o IPI */
  const ipiValor = custo / (1 + ipiF) * ipiF;
  const icmsValor = (custo - ipiValor) * icmsInF;
  const pisCofinsValor = (custo - ipiValor - icmsValor) * pisCofinsInF;
  const impEntradaValor = ipiValor + icmsValor + pisCofinsValor;

  const cmv = servicoUnit + (custo - ipiValor - icmsValor - pisCofinsValor);

  const denomIcms = 1 - icmsOutF - (icmsOutF * ipiF);
  const denomPisCofins = 1 - pisCofinsOutF;
  const denomMargem = 1 - margemF;

  let precoFinal = 0, impSaidaValor = 0, warn = false;
  if(denomMargem <= 0 || denomIcms <= 0 || denomPisCofins <= 0){
    warn = true;
  } else {
    const recLiq = cmv / denomMargem;
    precoFinal = (recLiq / denomIcms / denomPisCofins) * (1 + ipiF);
    impSaidaValor = precoFinal - recLiq;
  }
  const total = precoFinal * (parseFloat(qtd) || 0);
  return {impEntradaValor, cmv, impSaidaValor, precoFinal, total, warn};
}

function recalcAll(){
  const totalQtd = items.reduce((s, it) => s + (parseFloat(it.qtd) || 0), 0);
  let grandTotal = 0;
  let anyWarn = false;

  items.forEach(it => {
    const r = calcItemValues(it.custo, it.qtd, totalQtd);
    if(r.warn) anyWarn = true;
    grandTotal += r.total;
    const row = document.querySelector('tr[data-row-id="' + it.id + '"]');
    if(row){
      row.querySelector('.c-imp-entrada').textContent = fmtBRL(r.impEntradaValor);
      row.querySelector('.c-cmv').textContent = fmtBRL(r.cmv);
      row.querySelector('.c-imp-saida').textContent = fmtBRL(r.impSaidaValor);
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
}

function addItem(pn, custo, qtd){
  pn = pn || '';
  custo = custo || 0;
  qtd = (qtd === undefined || qtd === null) ? 1 : qtd;

  itemCounter++;
  const id = 'i' + itemCounter;
  items.push({id: id, pn: pn, custo: custo, qtd: qtd});

  const tr = document.createElement('tr');
  tr.setAttribute('data-row-id', id);
  tr.innerHTML =
    '<td><input type="text" class="cell-pn" value="' + escapeHtml(pn) + '" placeholder="PN"></td>' +
    '<td><input type="number" class="cell-custo" step="0.01" value="' + custo + '"></td>' +
    '<td><input type="number" class="cell-qtd" step="1" min="0" value="' + qtd + '"></td>' +
    '<td class="calc c-imp-entrada">R$ 0,00</td>' +
    '<td class="calc c-cmv">R$ 0,00</td>' +
    '<td class="calc c-imp-saida">R$ 0,00</td>' +
    '<td class="calc c-preco-final strong">R$ 0,00</td>' +
    '<td class="calc c-total strong">R$ 0,00</td>' +
    '<td><button type="button" class="row-remove" aria-label="Remover item">✕</button></td>';

  document.getElementById('items-tbody').appendChild(tr);

  tr.querySelector('.cell-pn').addEventListener('input', (e) => {
    const it = items.find(i => i.id === id);
    if(it) it.pn = e.target.value;
  });
  tr.querySelector('.cell-custo').addEventListener('input', (e) => {
    const it = items.find(i => i.id === id);
    if(it){ it.custo = parseFloat(e.target.value) || 0; recalcAll(); }
  });
  tr.querySelector('.cell-qtd').addEventListener('input', (e) => {
    const it = items.find(i => i.id === id);
    if(it){ it.qtd = parseFloat(e.target.value) || 0; recalcAll(); }
  });
  tr.querySelector('.row-remove').addEventListener('click', () => {
    items = items.filter(i => i.id !== id);
    tr.remove();
    recalcAll();
  });

  recalcAll();
  return id;
}

document.getElementById('add-item').addEventListener('click', () => addItem('', 0, 1));

const rateIds = ['icms-in','pis-in','cofins-in','ipi-in','icms-out','pis-out','cofins-out','iss-out','margem-out','frete-valor'];
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
addItem('', 0, 1);

/* ---------------------------------------------------------------------
   EXTRAÇÃO DO PDF
   O layout do fornecedor é fixo, mas os padrões abaixo são genéricos
   (busca por palavras-chave comuns + valores em R$). Ajuste os regex em
   PATTERNS conforme o modelo real do seu fornecedor para aumentar a
   precisão — a extração é sempre um ponto de partida editável, nunca
   a fonte final dos dados.
--------------------------------------------------------------------- */
const PATTERNS = {
  fornecedor: /(?:raz[aã]o social|fornecedor)\s*[:\-]?\s*(.+)/i,
  produto: /(?:descri[cç][aã]o|produto|item)\s*[:\-]?\s*(.+)/i,
  precoCompra: /(?:valor unit[aá]rio|pre[cç]o unit[aá]rio|valor total|pre[cç]o)\s*[:\-]?\s*R?\$?\s*([\d.,]+)/i
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

    const fMatch = fullText.match(PATTERNS.fornecedor);
    const pMatch = fullText.match(PATTERNS.produto);
    const cMatch = fullText.match(PATTERNS.precoCompra);

    if(fMatch) document.getElementById('f-fornecedor').value = fMatch[1].trim().slice(0,80);
    if(pMatch) document.getElementById('f-produto').value = pMatch[1].trim().slice(0,80);
    if(cMatch && items.length){
      const val = parseCurrencyToNumber(cMatch[1]);
      if(val !== null){
        items[0].custo = val;
        const row = document.querySelector('tr[data-row-id="' + items[0].id + '"]');
        if(row) row.querySelector('.cell-custo').value = val.toFixed(2);
        recalcAll();
      }
    }

    const found = [fMatch && 'fornecedor', pMatch && 'produto', cMatch && 'preço do 1º item'].filter(Boolean);
    if(found.length){
      statusEl.className = 'ok';
      statusEl.textContent = 'Extraído automaticamente: ' + found.join(', ') + '. Confira os campos abaixo antes de prosseguir.';
    } else {
      statusEl.className = 'err';
      statusEl.textContent = 'Não foi possível reconhecer os campos automaticamente neste layout. Preencha manualmente — ajuste os padrões em PATTERNS no código para este modelo de PDF.';
    }
  } catch(err){
    statusEl.className = 'err';
    statusEl.textContent = 'Erro ao ler o PDF: ' + err.message;
  }
});

/* ---------------------------------------------------------------------
   GERAÇÃO DA PROPOSTA — via impressão nativa do navegador
--------------------------------------------------------------------- */
document.getElementById('gerar-proposta').addEventListener('click', () => {
  const empresa = document.getElementById('f-empresa').value || 'Sua Empresa';
  const responsavel = document.getElementById('f-responsavel').value || '—';
  const cargo = document.getElementById('f-cargo').value || '';
  const email = document.getElementById('f-email').value || '';
  const telefone = document.getElementById('f-telefone').value || '';
  const endereco = document.getElementById('f-endereco').value || '';
  const cliente = document.getElementById('f-cliente').value || '—';
  const codigoProposta = document.getElementById('f-codigo-proposta').value || '';
  const oportunidade = document.getElementById('f-oportunidade').value || '—';
  const produto = document.getElementById('f-produto').value || 'Fornecimento';
  const sobreEmpresa = document.getElementById('f-sobre-empresa').value || '';
  const obsBase = document.getElementById('f-obs').value || '';
  const dataInput = document.getElementById('f-data').value;
  const dataFmt = dataInput ? new Date(dataInput + 'T00:00:00').toLocaleDateString('pt-BR') : new Date().toLocaleDateString('pt-BR');

  /* capa */
  document.getElementById('p-codigo-sub').textContent = codigoProposta;
  document.getElementById('p-titulo-projeto').textContent = produto;
  document.getElementById('p-data-cover').textContent = dataFmt;
  document.getElementById('p-responsavel').textContent = responsavel;
  document.getElementById('p-cargo').textContent = cargo;
  document.getElementById('p-email').textContent = email;
  document.getElementById('p-telefone').textContent = telefone;
  document.getElementById('p-endereco').textContent = endereco;
  document.getElementById('p-cliente-cover').textContent = cliente;
  document.getElementById('p-oportunidade-cover').textContent = oportunidade;

  const logoWrap = document.getElementById('p-logo-wrap');
  if(logoDataUrl){
    logoWrap.innerHTML = '<img class="logo-img" src="' + logoDataUrl + '" alt="Logo">';
  } else {
    logoWrap.innerHTML = '<div class="logo-placeholder">' + escapeHtml(empresa) + '</div>';
  }

  /* cabeçalhos das páginas internas */
  ['p-header-empresa-1','p-header-empresa-2'].forEach(id => document.getElementById(id).textContent = empresa);
  ['p-header-endereco-1','p-header-endereco-2'].forEach(id => document.getElementById(id).textContent = endereco);
  const footerLabel = (codigoProposta ? codigoProposta + ' — ' : '') + produto;
  document.getElementById('p-footer-titulo-1').textContent = footerLabel;
  document.getElementById('p-footer-titulo-2').textContent = footerLabel;

  /* revisão / apresentação / objetivo */
  document.getElementById('p-rev-data').textContent = dataFmt;
  const apresentacaoBlock = document.getElementById('p-apresentacao-block');
  if(sobreEmpresa.trim()){
    apresentacaoBlock.classList.remove('hidden');
    document.getElementById('p-apresentacao-text').textContent = sobreEmpresa;
  } else {
    apresentacaoBlock.classList.add('hidden');
  }
  document.getElementById('p-objetivo-text').textContent =
    'Esta proposta tem por objetivo apresentar as condições comerciais para o ' +
    produto.charAt(0).toLowerCase() + produto.slice(1) + ' para ' + cliente + '.';

  /* escopo de fornecimento */
  const totalQtd = items.reduce((s, it) => s + (parseFloat(it.qtd) || 0), 0);
  let grandTotal = 0;
  const rowsHtml = items.map((it, idx) => {
    const r = calcItemValues(it.custo, it.qtd, totalQtd);
    grandTotal += r.total;
    return '<tr>' +
      '<td>' + (idx + 1) + '</td>' +
      '<td>' + escapeHtml(it.pn || '—') + '</td>' +
      '<td style="text-align:center">' + (it.qtd || 0) + '</td>' +
      '<td class="num">' + fmtBRL(r.precoFinal) + '</td>' +
      '<td class="num">' + fmtBRL(r.total) + '</td>' +
    '</tr>';
  }).join('');
  document.getElementById('p-items-tbody').innerHTML = rowsHtml || '<tr><td colspan="5">Nenhum item adicionado.</td></tr>';
  document.getElementById('p-total-geral').textContent = fmtBRL(grandTotal);

  /* notas */
  let obsFinal = obsBase;
  if(freteAtivo() && num('frete-valor') > 0){
    const freteNote = 'Frete de ' + fmtBRL(num('frete-valor')) + ' considerado na composição do preço.';
    obsFinal = obsFinal ? (obsFinal + '\n\n' + freteNote) : freteNote;
  }
  const notasBlock = document.getElementById('p-notas-block');
  if(obsFinal.trim()){
    notasBlock.classList.remove('hidden');
    document.getElementById('p-obs').textContent = obsFinal;
  } else {
    notasBlock.classList.add('hidden');
  }

  /* contatos */
  const contatoLinhas = [
    empresa + ' agradece a oportunidade e coloca-se à disposição para quaisquer esclarecimentos adicionais necessários.',
    '',
    'Atenciosamente,',
    '',
    responsavel,
    [cargo, empresa].filter(Boolean).join(' — '),
    [telefone, email].filter(Boolean).join(' · ')
  ].filter(line => line !== undefined);
  document.getElementById('p-contatos-text').textContent = contatoLinhas.join('\n');

  window.print();
});
