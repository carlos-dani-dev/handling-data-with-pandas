import pandas as pd
import numpy as np
import os

PATH = "C:\\Users\\carlo\\OneDrive\\Área de Trabalho\\SEAD\\pratica_manipulando_tabelas\\exemplo_usado_tabela.xlsx"

def importar_plan(path):
    file = pd.ExcelFile(path)
    plans = file.sheet_names
    
    # Carregar e pré-filtrar os DataFrames
    df_plans_pre_filtrado = []
    for plan in plans:
        df = pd.read_excel(file, sheet_name=plan)
        df['null count'] = df.isnull().sum(axis=1)
        df = df[df['null count'] < 4]
        
        ultimo_orgao = None
        ultimo_cargo = None
        ultima_atividade = None
        ultimo_total = None

        for i in range(len(df)):
            if pd.isna(df.iat[i, 0]) and ultimo_orgao is not None: df.iat[i, 0] = ultimo_orgao
            else: ultimo_orgao = df.iat[i, 0]
            if pd.isna(df.iat[i, 1]) and ultimo_cargo is not None: df.iat[i, 1] = ultimo_cargo
            else: ultimo_cargo = df.iat[i, 1]
            if pd.isna(df.iat[i, 3]) and ultimo_total is not None: df.iat[i, 3] = ultimo_total
            else: ultimo_total = df.iat[i, 3]
            if pd.isna(df.iat[i, 2]) and ultima_atividade is not None: df.iat[i, 2] = ultima_atividade
            else: ultima_atividade = df.iat[i, 2]
        df_plans_pre_filtrado.append(df)


    return df_plans_pre_filtrado


def exportar_planilhas(dfs_planilhas_filtrados):
    cont = 0
    novo_path = "filtragens/"
    nome_file = "pratica_pandas_filtrado.xlsx"
    if not os.path.exists(novo_path): os.makedirs(novo_path+nome_file)
    
    with pd.ExcelWriter(novo_path + nome_file, engine="xlsxwriter") as writer:
        for df_filtrado in dfs_planilhas_filtrados:
            cont += 1
            sheet_name = "Planilha " + str(cont)
            df_filtrado.to_excel(writer, sheet_name=sheet_name, index=False)
            worksheet = writer.sheets[sheet_name]

            # Mesclar células com valores repetidos
            for col_num, col_name in enumerate(df_filtrado.columns):
                merge_ranges(df_filtrado, worksheet, col_num, col_name)
            print(df_filtrado.columns)


def merge_ranges(df, worksheet, col_num, col_name):
    start_row = 1
    end_row = 1
    value = df.iat[0, col_num]

    for row in range(1, len(df) + 1):
        if row < len(df) and df.iat[row, col_num] == value:
            end_row += 1
        else:
            if start_row != end_row:
                worksheet.merge_range(start_row, col_num, end_row, col_num, value)
            start_row = row + 1
            end_row = start_row
            if row < len(df):
                value = df.iat[row, col_num]


df_plans_pre_filtrado = importar_plan(PATH)
exportar_planilhas(df_plans_pre_filtrado)
print(df_plans_pre_filtrado[0])
