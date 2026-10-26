# Tabela de Preços e Impostos + extração de PN de PDF

Arquivos do projeto:
- `index.html`: interface da planilha e seção de extração de PDF.
- `script.js`: leitura da planilha e extração de PN/especificações do PDF.
- `style.css`: estilos da ferramenta e da tabela de PN.

## Como usar
1. Mantenha `index.html`, `script.js` e `style.css` na mesma pasta.
2. Abra `index.html` no navegador.
3. Para extração, selecione um PDF com texto selecionável e use **Extrair texto**.
4. A tabela gerada contém ITEM, DESCRIÇÃO (PN em negrito e especificações em bullets) e QTD. A quantidade aparece como `01` por padrão, pois esta versão extrai somente PN e descrição técnica.
5. Clique em **Copiar tabela** para colar no Word.

As bibliotecas XLSX e PDF.js são carregadas de CDNs, portanto a primeira abertura exige conexão com a internet. PDFs digitalizados como imagens podem exigir OCR e não são processados por esta versão.
