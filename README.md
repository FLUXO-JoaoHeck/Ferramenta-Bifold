# Tabela de Preços e Impostos

Site estático (`index.html`, `style.css`, `script.js`), sem backend, que lê a
planilha de revenda (Ver. 3.0) e gera a seção "Preços e Impostos" da proposta,
pronta para copiar e colar no Word.

## Como usar
1. **Planilha:** escolha o arquivo Excel (aba "Revenda") e confira os itens lidos e os avisos.
2. **PDF do fornecedor:** escolha o PDF da cotação; o site lê o PN, a descrição técnica e a quantidade.
3. **Resultado final:** tudo em uma só seção.
   - **Opções** (não são copiadas): preço exibido (Revenda direta ou Portal), estado do cliente,
     como o IPI aparece e se o destino é Zona Franca de Manaus.
   - **Vínculo dos itens**, quando planilha e PDF estão carregados: cada item da planilha recebe o
     PN e as especificações de um item do PDF (por padrão, na mesma ordem; dá para trocar).
   - **Folha:** uma página só, com o Escopo de Fornecimento (proposta técnica) em cima de Preços e
     Impostos (proposta comercial). Os campos de margem (topo, rodapé, esquerda, direita, em mm)
     reservam o espaço do cabeçalho e do rodapé do modelo.
   - **Imprimir / Salvar PDF:** imprime só a folha. Na janela de impressão, escolha "Salvar como
     PDF", deixe as margens em "Padrão" e desmarque "Cabeçalhos e rodapés".
   - **Copiar:** "Copiar folha inteira" (títulos + as duas tabelas) ou só uma das tabelas, para colar
     (Ctrl+V) no Word.

## Folha e modelo (papel timbrado)
- As margens padrão (topo 42, rodapé 24, esquerda 35, direita 25 mm) foram estimadas a partir da
  proposta de exemplo e são lembradas no navegador. Ajuste-as até o conteúdo cair na área livre do
  modelo.
- Com poucos itens a folha cabe em uma página; com muitos, continua na seguinte, repetindo as
  mesmas margens, e as linhas das tabelas não são cortadas ao meio.
- Próximo passo previsto: usar o arquivo em branco do modelo (cabeçalho e rodapé) como fundo da
  página impressa.

## O que o site lê da planilha
O site NÃO recalcula preços: lê os valores já calculados pelo Excel
(o arquivo precisa ter sido salvo pelo Excel, com os resultados em cache).

O arquivo pode ter várias abas (Entradas, Lista de Preços, Resumo, NCM...). O
site procura **somente** a aba chamada "Revenda" (sem diferenciar maiúsculas,
acentos ou espaços nas pontas) e lê apenas ela. Se não a encontrar, para e
mostra a lista de abas do arquivo, sem tentar outra aba no lugar.

| Dado | Coluna |
|---|---|
| Item (código) | A |
| Descrição (nome do item, coluna "DESCRIÇÃO") | E |
| NCM | G |
| Quantidade | P |
| ICMS / PIS-COFINS / IPI de saída (%) | R / S / T |
| Preço unitário e total, revenda direta | W / X |
| Preço unitário e total, portal (líquido de ICMS e IPI) | Z / AA |

- Cabeçalho: a linha é detectada automaticamente (coluna A = "ITEM" e coluna X =
  "P. TOTAL", nas primeiras 60 linhas), então a tabela pode começar na linha 13,
  14, 15 etc. Os itens são lidos nas até 100 linhas abaixo dela.
- Proposta e cliente são achados pelo rótulo, acima do cabeçalho; o subtotal é o
  primeiro número da coluna X acima do cabeçalho.
- Itens sem quantidade ou com erro (`#REF!`) nos preços são ignorados, com aviso
  na tela. Se o cabeçalho não for encontrado, a leitura para com uma mensagem.
- Nome do item: o site procura a coluna **"DESCRIÇÃO"** no cabeçalho (qualquer
  posição) e a usa como identificação na coluna "Modelo" da tabela. Se não achar,
  usa a coluna E e avisa na tela.
- A soma dos itens é conferida com o subtotal da planilha quando todos os itens
  são lidos.

## PDF do fornecedor e vínculo com a planilha
- Cada item do PDF é reconhecido pela linha `P/N: <código> – (<descrição>)`; a quantidade é o
  primeiro número do trio "quantidade, preço unitário, valor" que vem depois da descrição.
- A descrição é separada em tópicos "Rótulo: valor" (uma vírgula seguida de um rótulo com
  dois-pontos inicia um novo tópico), mantendo o texto do fornecedor, inclusive "N/A".
- Quebras de linha no meio do código (`XW-00-` / `01580-...`) são juntadas; cabeçalho e rodapé
  repetidos em cada página são descartados (lista `PDF_NOISE` em `script.js`).
- O pdf.js só é baixado quando um PDF é escolhido. Os preços do PDF não são lidos.
- **Vínculo:** o nome do item vem da planilha (coluna "DESCRIÇÃO") e o PN e as especificações
  vêm do PDF. O item N da planilha usa o item N do PDF; o vínculo pode ser trocado na seção 3 (Resultado final). O
  site avisa quando as quantidades de itens diferem, quando a quantidade não bate e quando
  dois itens apontam para o mesmo PN. As tabelas usam a quantidade da planilha.
- **Proposta técnica:** com planilha carregada, segue a ordem dela (os mesmos itens da tabela de
  preços); itens sem PN vinculado aparecem só com a descrição. Sem planilha, lista os itens do PDF.
- **Proposta comercial:** coluna "Modelo" = PN (negrito) + descrição da planilha; sem PDF, só a
  descrição.
- Layout testado: cotação da Rotork. Outros fornecedores precisam usar a convenção "P/N:".

## Publicar no GitHub Pages
Suba os três arquivos juntos para a raiz do repositório e ative o Pages em
Settings → Pages (branch `main`, pasta `/ (root)`).

## Observações
- A biblioteca de leitura de Excel (SheetJS) é carregada por CDN: é preciso
  estar online.
- A formatação da tabela copiada segue o documento original (Open Sans, 8 pt,
  linhas horizontais finas, alinhamento centralizado, linhas de 2 cm).
