# Tabela de Preços e Impostos

Site estático (`index.html`, `style.css`, `script.js`), sem backend, que lê a
planilha de revenda (Ver. 3.0) e gera a seção "Preços e Impostos" da proposta,
pronta para copiar e colar no Word.

## Como usar
1. Abra o site e escolha o arquivo Excel (aba "Revenda").
2. Confira os itens lidos e os avisos exibidos.
3. Em "Opções da tabela", escolha o preço (Revenda direta ou Portal), o estado
   do cliente, como o IPI aparece e se o destino é Zona Franca de Manaus.
4. Clique em "Copiar tabela" e cole (Ctrl+V) no Word.

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
| Descrição | E |
| NCM | G |
| Quantidade | P |
| ICMS / PIS-COFINS / IPI de saída (%) | R / S / T |
| Preço unitário e total, revenda direta | W / X |
| Preço unitário e total, portal (líquido de ICMS e IPI) | Z / AA |
| Frete rateado ("Serviços") | E9 |

- Linhas lidas: 14 a 113. Itens sem quantidade ou com erro (`#REF!`) nos preços
  são ignorados, com aviso na tela.
- PN: o site procura uma coluna chamada "PN" (ou "Part Number") nas linhas 12 e
  13. Se não achar, usa a descrição como identificação do item.
- A soma dos itens é conferida com o subtotal da planilha (X11) quando todos os
  itens são lidos.

## Publicar no GitHub Pages
Suba os três arquivos juntos para a raiz do repositório e ative o Pages em
Settings → Pages (branch `main`, pasta `/ (root)`).

## Observações
- A biblioteca de leitura de Excel (SheetJS) é carregada por CDN: é preciso
  estar online.
- A formatação da tabela copiada segue o documento original (Open Sans, 8 pt,
  linhas horizontais finas, alinhamento centralizado, linhas de 2 cm).
