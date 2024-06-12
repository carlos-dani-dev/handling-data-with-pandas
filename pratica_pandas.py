import pandas as pd
import os


PATH = "C:\\Users\\carlo\\OneDrive\\Área de Trabalho\\SEAD\\pratica_manipulando_tabelas\\pratica_pandas.xlsx"


# pd.set_option('display.max_rows', None)
# pd.set_option('display.max_columns', None)
# pd.set_option('display.width', None)

def importar_planilhas(path):
    """
    Retorna: Lista com DataFrame de cada uma das planilhas do arquivo de excel
    """
    
    arquivo_excel = pd.ExcelFile(path)
    planilhas = arquivo_excel.sheet_names
    dfs_planilhas = [pd.read_excel(arquivo_excel, sheet_name=planilha) for planilha in planilhas]
    dfs_planilhas_filtrados = []

    for df in dfs_planilhas:
        df['NULL COUNT'] = df.isnull().sum(axis=1) # conta itens nulos em cada linha do df

    for df in dfs_planilhas:
        # filtra o df por linhas que possuam a coluna 'NULL COUNT' < 2
        dfs_planilhas_filtrados.append(df[df['NULL COUNT'] < 2])


    return (dfs_planilhas, dfs_planilhas_filtrados)


def exportar_planilhas(dfs_planilhas_filtrados):
    """
    Exporta planilhas após manipulação para um arquivo de excel em específico
    """

    cont = 0

    novo_path = "filtragens/"
    nome_file = "pratica_pandas_filtrado.xlsx"
    if not os.path.exists(novo_path): os.makedirs(novo_path)
    with pd.ExcelWriter(novo_path+nome_file, engine="xlsxwriter") as f:
        for df_filtrado in dfs_planilhas_filtrados:
            cont+=1
            df_filtrado.to_excel(f, sheet_name="Planilha "+str(cont), index=False)


dfs_planilhas, dfs_planilhas_filtrados = importar_planilhas(PATH)
print(dfs_planilhas[0])
print(dfs_planilhas_filtrados[0])
#exportar_planilhas(dfs_planilhas_filtrados=dfs_planilhas_filtrados)