import xlrd
import pandas as pd
from datetime import datetime

data_atual = datetime.now().strftime('%d_%m_%y')  # Data atual para renomear a nova planilha

# GET Dados saldo estoque disponível do Tiny
EstoqueTiny = 'saldos-em-estoque18-02-2025-09-34-54.xls'
DF_Estoque_Tiny = xlrd.open_workbook(EstoqueTiny, ignore_workbook_corruption=True)
DF_Estoque_Tiny_PD = pd.read_excel(DF_Estoque_Tiny)

ColunaSKU1 = DF_Estoque_Tiny_PD.iloc[0:, 0]
ColunaEstoqueDisponivel = DF_Estoque_Tiny_PD.iloc[0:, 4]

SaldoEstoqueDisponivel = dict()

for x, y in zip(ColunaSKU1, ColunaEstoqueDisponivel):
    SaldoEstoqueDisponivel[x] = y

# GET Dados saldo estoque Full
DF_Estoque_Full = pd.read_excel('stock_general_full_18083826_c808fdc32d6f309a3176c4a1a5bb9fe1.xlsx')

ColunaSKU2 = DF_Estoque_Full.iloc[14:, 3]
ColunaSaldoEstoqueFull = DF_Estoque_Full.iloc[14:, 20]
SaldoEstoqueFull = dict()

for x, y in zip(ColunaSKU2, ColunaSaldoEstoqueFull):
    SaldoEstoqueFull[x] = y

# GET Dados necessidade compra
Saidas_periodo = 'necessidade-compra_18-02-2025-09-08-03.xls'
DF_SaidasPeriodo = xlrd.open_workbook(Saidas_periodo, ignore_workbook_corruption=True)
DF_SaidasPeriodo_PD = pd.read_excel(DF_SaidasPeriodo)

ColunaSKU3 = DF_SaidasPeriodo_PD.iloc[0:, 1]
Saidas = DF_SaidasPeriodo_PD.iloc[0:, 9]

# Criar nova planilha
WorkNewCompras = {
    'SKU_Seller': [],
    'Estoque Geral': [],
    'Estoque Full': [],
    'Estoque Comprado': [],  # A Definir
    'Saidas Periodo': [],
    'Sugestão de Compras': [],
}

# Com base na necessidade extrair os dados e juntar
for x, y in zip(ColunaSKU3, Saidas):
    try:
        Estoque_geral = SaldoEstoqueDisponivel[x]
    except KeyError:
        Estoque_geral = 0

    try:
        EstoqueFull = SaldoEstoqueFull[x]
    except KeyError:
        EstoqueFull = 0

    EstoqueComprado = 0
    SugestaoCompras = y - (Estoque_geral + EstoqueFull + EstoqueComprado)

    if SugestaoCompras <= 0:
        continue

    WorkNewCompras['SKU_Seller'].append(x)
    WorkNewCompras['Estoque Geral'].append(Estoque_geral)
    WorkNewCompras['Estoque Full'].append(EstoqueFull)
    WorkNewCompras['Estoque Comprado'].append(EstoqueComprado)
    WorkNewCompras['Saidas Periodo'].append(y)
    WorkNewCompras['Sugestão de Compras'].append(SugestaoCompras)

pd.DataFrame(WorkNewCompras).to_excel(f'SugestãoCompras{data_atual}.xlsx', index=False)
