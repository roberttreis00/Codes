import xlrd
import pandas as pd
from datetime import datetime

data_atual = datetime.now().strftime('%d_%m_%y')

sheetwork_necessidade_compra_tiny = 'necessidade-compra_28-01-2025-14-01-05.xls'
sheetwork_estoque_geral_full = 'stock_general_full_28130141_aa5d87e067b9f193b217c27a29e7bbc1.xlsx'

df_estoque_tiny_geral = xlrd.open_workbook(sheetwork_necessidade_compra_tiny, ignore_workbook_corruption=True)
df_full = pd.read_excel(sheetwork_estoque_geral_full)
workbook = df_estoque_tiny_geral.sheet_by_index(0)

ColunaSKUFull = df_full.iloc[14:, 3]
ColunaQTDEstoque = df_full.iloc[14:, 19]

Rows = workbook.nrows
Estoque_Full = {}

WorkNewCompras = {
    'SKU_Seller': [],
    'Estoque Geral': [],
    'Estoque Full': [],
    'Estoque Comprado': [],
    'Saidas Periodo': [],
    'Sugestão de Compras': [],
}

for x, y in zip(ColunaSKUFull, ColunaQTDEstoque):
    Estoque_Full[x] = y

for row in range(1, Rows):
    estoque_full = 0

    sku_tiny = workbook.cell(row, 1).value
    estoque_virtual = int(workbook.cell(row, 8).value)
    saidas_periodo = int(workbook.cell(row, 9).value)

    # Verifica se tem estoque no FULL
    try:
        estoque_full = Estoque_Full[sku_tiny]
        quantidade_disponivel = estoque_virtual + estoque_full
        sugestao_compra = saidas_periodo - quantidade_disponivel
    except KeyError:
        sugestao_compra = saidas_periodo - estoque_virtual

    if sugestao_compra == 0:
        continue

    WorkNewCompras['Estoque Full'].append(estoque_full)
    WorkNewCompras['SKU_Seller'].append(sku_tiny)
    WorkNewCompras['Estoque Geral'].append(estoque_virtual)
    WorkNewCompras['Saidas Periodo'].append(saidas_periodo)

    # Aqui verifica SE foi comprado e ainda não faturou
    WorkNewCompras['Estoque Comprado'].append(0)

    # Calculo de sugestão do que comprar

    WorkNewCompras['Sugestão de Compras'].append(sugestao_compra)

df_new = pd.DataFrame(WorkNewCompras)
df_new.to_excel(f'SugestãoCompras{data_atual}.xlsx', index=False)
