from flask import Flask, request, render_template, send_file
import pandas as pd
import xlrd
from datetime import datetime
import os

app = Flask(__name__)

# Pasta para salvar arquivos temporários
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        # Obter arquivos
        file1 = request.files['file1']
        file2 = request.files['file2']

        # Salvar arquivos
        file1_path = os.path.join(UPLOAD_FOLDER, file1.filename)
        file2_path = os.path.join(UPLOAD_FOLDER, file2.filename)
        file1.save(file1_path)
        file2.save(file2_path)

        # Processar arquivos
        output_file = process_files(file1_path, file2_path)

        # Remover arquivos temporários
        os.remove(file1_path)
        os.remove(file2_path)

        # Enviar arquivo gerado para download
        return send_file(output_file, as_attachment=True)

    return render_template('index.html')


def process_files(file1, file2):
    data_atual = datetime.now().strftime('%d_%m_%y')
    df_estoque_tiny_geral = xlrd.open_workbook(file1, ignore_workbook_corruption=True)
    df_full = pd.read_excel(file2)
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
        WorkNewCompras['Estoque Comprado'].append(0)
        WorkNewCompras['Sugestão de Compras'].append(sugestao_compra)

    df_new = pd.DataFrame(WorkNewCompras)
    output_file = f'SugestãoCompras_{data_atual}.xlsx'
    df_new.to_excel(output_file, index=False)

    return output_file


if __name__ == '__main__':
    app.run(debug=True)
