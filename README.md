# Gerador de Estruturas PEP da AGCGNO

Aplicação 100% no navegador para gerar estruturas PEP com subníveis a partir do CAPEX.

## Link

`https://agcgno.github.io/Gerador-PEP-/index.html`

## Como usar

1. Abra o arquivo `index.html` no navegador.
2. Cole as colunas **M:AH** do CAPEX no campo "Dados CAPEX".
3. (Opcional) Use o formulário manual para inserir um projeto específico.
4. Clique em `Interpretar dados`.
5. Clique em `Gerar PEPs`.
6. Use `Copiar resultado` ou `Baixar Excel`.

## Colunas utilizadas do CAPEX

- **M**: Coletor de Custo  
- **Q**: ID real  
- **R**: Descrição do projeto  
- **AG**: Objeto  
- **AH**: Local/Usina  

## Saída (tabela final)

- Usina  
- Nível  
- Elemento PEP  
- Denominação  
- ID real  
- Tipo  
- Pri  
- Centro de lucro  

## Regras principais

- Gera 5 linhas por projeto (níveis 2, 3, 4.001, 4.002, 4.003).
- Nível 2 e 3: descrição abreviada e limitada a 40 caracteres.
- Nível 4.001: `ID real - CUSTO COMUM`
- Nível 4.002: `ID real - SERVIÇO`
- Nível 4.003: `ID real - OBJETO` (limitado a 40 caracteres)
- Sequencial respeita o número inicial mínimo por usina.

## Mapeamento de usinas

- BALBINA → prefixo `0161`, centro de lucro `NO10101003`, início `01`
- COARACY NUNES → prefixo `0159`, centro de lucro `N010102001`, início `03`
- CURUA-UNA → prefixo `0160`, centro de lucro `N010103001`, início `04`
- SAMUEL → prefixo `0160`, centro de lucro `N010104001`, início `02`
- TUCURI → prefixo `0162`, centro de lucro `N010105001`, início `14`

## Recalcular sequencial

No bloco "Pré-visualização", selecione a usina e informe o novo início.
O sistema ajusta apenas aquela usina e mantém as demais inalteradas.
Ao selecionar a usina, as linhas dela sobem para facilitar a análise.

## Arquivos

- `index.html`
- `style.css`
- `app.js`
# Gerador-de-Estruturas-PEP-da-AGCGNO
