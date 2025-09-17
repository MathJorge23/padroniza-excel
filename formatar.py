import pandas as pd

# Nome do arquivo de entrada (existente)
arquivo_entrada = "arquivo_entrada.xlsx"

# Nome do arquivo de saída (novo arquivo)
arquivo_saida = "arquivo_saida.xlsx"

# Lê todas as abas da planilha existente
# Retorna um dicionário: {nome_aba: DataFrame}
planilhas = pd.read_excel(arquivo_entrada, sheet_name=None)

# Cria um ExcelWriter usando xlsxwriter
with pd.ExcelWriter(arquivo_saida, engine="xlsxwriter") as writer:
    for aba_nome, df in planilhas.items():
        # 1) Apagar a primeira coluna
        if df.shape[1] > 0:
            df = df.iloc[:, 1:]  # remove a primeira coluna

        # 2) Garantir que todos os nomes de colunas sejam strings e únicos
        novas_colunas = []
        for i, col in enumerate(df.columns, start=1):
            if pd.isna(col) or str(col).strip() == "":
                col_nome = f"Coluna_{i}"
            else:
                col_nome = str(col).strip()
            # evitar duplicatas
            while col_nome in novas_colunas:
                col_nome += "_1"
            novas_colunas.append(col_nome)
        df.columns = novas_colunas

        # 3) Escrever a aba no Excel
        df.to_excel(writer, sheet_name=aba_nome, startrow=0, startcol=0, index=False)

        # 4) Criar tabela com estilo usando XlsxWriter
        workbook = writer.book
        worksheet = writer.sheets[aba_nome]

        max_row, max_col = df.shape
        # xlsxwriter usa índice 0 para linha/coluna
        worksheet.add_table(0, 0, max_row, max_col - 1, {
            'columns': [{'header': h} for h in df.columns],
            'style': 'Table Style Medium 9'  # estilo de tabela do Excel
        })

        # 5) Ajustar largura das colunas automaticamente
        for i, col in enumerate(df.columns):
            max_lenght = max(
                df[col].astype(str).map(len).max(),  # maior conteúdo da coluna
                len(col)  # largura do cabeçalho
            )
            worksheet.set_column(i, i, max_lenght + 2)  # adicionar margem

        # 6) Aplicar wrap_text em todas as células (menos cabeçalho)
        cell_format = workbook.add_format({'text_wrap': True})
        for row in range(1, max_row + 1):
            worksheet.set_row(row, None, cell_format)
