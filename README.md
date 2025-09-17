# 📄 README – Script de Formatação de Planilhas Excel

## 📌 Descrição
Este script em Python lê uma planilha Excel existente (`tmp001.xlsx`) com múltiplas abas, aplica algumas transformações em cada aba e gera uma nova planilha formatada (`v1_relatorio_rafas.xlsx`).  
Ele foi feito utilizando a biblioteca **pandas** para manipulação de dados e o **XlsxWriter** (via `ExcelWriter`) para aplicar estilos e formatações avançadas no Excel.

---

## 📥 Imports Utilizados

```python
import pandas as pd
```

- **pandas**  
  - Motivo: É a principal biblioteca para manipulação de dados em Python.  
  - Uso no script:
    - `pd.read_excel()` → para ler todas as abas do arquivo Excel de entrada.  
    - `DataFrame` → usado para armazenar e manipular os dados de cada aba.  
    - `pd.ExcelWriter(..., engine="xlsxwriter")` → permite escrever os DataFrames em um novo Excel, usando o motor **XlsxWriter** para aplicar estilos.  
    - `pd.isna()` → usado para detectar nomes de colunas vazios.  

- **xlsxwriter** (não precisa importar manualmente, pois é usado internamente pelo `pandas.ExcelWriter`)  
  - Motivo: É o motor escolhido para escrita do Excel porque oferece suporte a **formatações avançadas**, como tabelas com estilos, ajuste automático de colunas e `wrap_text`.  
  - Uso no script:  
    - `writer.book` → acesso ao objeto do workbook (arquivo Excel).  
    - `writer.sheets` → acesso direto a cada aba escrita.  
    - `worksheet.add_table()` → cria uma tabela formatada no Excel.  
    - `worksheet.set_column()` e `worksheet.set_row()` → ajustam largura de colunas e estilos de células.  
    - `workbook.add_format({'text_wrap': True})` → cria um estilo de quebra de linha automática.  

---

## 🚀 Funcionalidades do Script
1. **Leitura de todas as abas da planilha original**  
   - Usa `pd.read_excel(..., sheet_name=None)` para carregar todas as abas em um dicionário `{nome_da_aba: DataFrame}`.

2. **Processamento de cada aba**  
   Para cada aba da planilha original:
   - Remove a primeira coluna (caso exista).  
   - Padroniza os nomes das colunas:
     - Garante que todos os nomes sejam `string`.  
     - Substitui colunas sem nome por `Coluna_X`.  
     - Resolve duplicatas adicionando `_1` ao final.  

3. **Escrita no novo Excel**  
   - Cada aba processada é escrita na planilha de saída (`v1_relatorio_rafas.xlsx`) preservando seu nome original.

4. **Estilização da planilha**  
   - Cria uma **tabela formatada** com estilo `"Table Style Medium 9"`.  
   - Ajusta automaticamente a **largura das colunas** com base no conteúdo.  
   - Aplica `wrap_text` (quebra de linha automática) em todas as células de dados (exceto cabeçalhos).

---

## 🛠️ Requisitos
Antes de rodar o script, instale as dependências necessárias:

```bash
pip install pandas xlsxwriter
```

---

## ▶️ Como Executar
1. Coloque a planilha de entrada no mesmo diretório do script e nomeie como **`tmp001.xlsx`** (ou ajuste no código).  
2. Execute o script com Python:

```bash
python nome_do_script.py
```

3. O resultado será salvo como **`v1_relatorio_rafas.xlsx`** no mesmo diretório.

---

## 📊 Exemplo de Transformações
### Antes (aba original):
| (vazio) | Nome   | Idade |
|---------|--------|-------|
| 1       | Ana    | 25    |
| 2       | Bruno  | 30    |

### Depois (aba processada):
| Nome   | Idade |
|--------|-------|
| Ana    | 25    |
| Bruno  | 30    |

- A primeira coluna foi removida.  
- Os nomes das colunas foram mantidos, mas caso estivessem em branco seriam renomeados.  
- Foi criada uma tabela formatada com estilo do Excel.  
- As colunas foram ajustadas automaticamente.  

---

## 📌 Observações
- O script é **genérico**: funciona para qualquer planilha com múltiplas abas.  
- Caso existam colunas sem título ou com nomes repetidos, elas serão renomeadas automaticamente para evitar erros.  
- O estilo `"Table Style Medium 9"` pode ser alterado para qualquer estilo disponível no Excel.  
