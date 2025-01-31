import xlrd
import pandas as pd
from datetime import datetime

data_atual = datetime.now().strftime('%d_%m_%y') # Data atual para renomear a nova planilha

sheetwork_necessidade_compra_tiny = 'necessidade-compra_31-01-2025-09-15-41.xls'
sheetwork_estoque_geral_full = 'stock_general_full_31081638_615e6c642c50da1ac8e5b8abdd671ef9.xlsx'

df_estoque_tiny_geral = xlrd.open_workbook(sheetwork_necessidade_compra_tiny, ignore_workbook_corruption=True)
df_full = pd.read_excel(sheetwork_estoque_geral_full)
workbook = df_estoque_tiny_geral.sheet_by_index(0)

ColunaSKUFull = df_full.iloc[14:, 3]
ColunaQTDEstoque = df_full.iloc[14:, 19]
EntradaPendente = df_full.iloc[14:, 20]

Rows = workbook.nrows
Estoque_Full = {}

WorkNewCompras = {
    'SKU_Seller': [],
    'Estoque Geral': [],
    'Estoque Full': [],
    'Estoque Comprado': [],  # A Definir
    'Saidas Periodo': [],
    'Sugestão de Compras': [],
}

for x, y, z in zip(ColunaSKUFull, ColunaQTDEstoque, EntradaPendente):
    Estoque_Full[x] = y - z

for row in range(1, Rows):
    estoque_full = 0

    sku_tiny = workbook.cell(row, 1).value
    estoque_virtual = int(workbook.cell(row, 8).value)
    saidas_periodo = int(workbook.cell(row, 9).value)

    # Verifica se tem estoque no FULL com base na planilha/ Se não = 0
    try:
        estoque_full = Estoque_Full[sku_tiny]
        quantidade_disponivel = estoque_virtual + estoque_full
        sugestao_compra = saidas_periodo - quantidade_disponivel  # Calcula a sugestão EstoqueS - Vendas Periodo
    except KeyError:
        sugestao_compra = saidas_periodo - estoque_virtual

    if sugestao_compra == 0:
        continue

    WorkNewCompras['Estoque Full'].append(estoque_full)
    WorkNewCompras['SKU_Seller'].append(sku_tiny)
    WorkNewCompras['Estoque Geral'].append(estoque_virtual)
    WorkNewCompras['Saidas Periodo'].append(saidas_periodo)

    # Aqui verifica SE já foi comprada A Definir
    WorkNewCompras['Estoque Comprado'].append(0)
    WorkNewCompras['Sugestão de Compras'].append(sugestao_compra)

df_new = pd.DataFrame(WorkNewCompras)
df_new.to_excel(f'SugestãoCompras{data_atual}.xlsx', index=False)
