from flask import Flask, render_template_string, request, send_file
import pandas as pd
import numpy as np
import io

app = Flask(__name__)

HTML_PAGE = '''
<!doctype html>
<title>Dock & Matera Upload</title>
<h2>Upload Dock (Excel) and Matera (CSV) Files</h2>
<form method=post enctype=multipart/form-data>
  <label>Dock (Excel):</label><br>
  <input type=file name=dock_file required><br><br>
  <label>Matera (CSV):</label><br>
  <input type=file name=matera_file required><br><br>
  <label>Depara (xlsm):</label><br>
  <input type=file name=depara_file required><br><br>
  <input type=submit value="Process Files">
</form>
'''

@app.route('/', methods=['GET', 'POST'])
def upload_files():
    if request.method == 'POST':
        dock_file = request.files['dock_file']
        matera_file = request.files['matera_file']
        depara_file = request.files['depara_file']

        ### 1. Leitura do Dock com tratamento
        try:
            dock_df = pd.read_excel(dock_file, dtype=str)
            start_index = dock_df[dock_df['Unnamed: 2'].notna()].index[0]
            dock_df = dock_df.iloc[start_index:].reset_index(drop=True)
            dock_df.columns = dock_df.iloc[0]
            dock_df = dock_df.iloc[1:].reset_index(drop=True)
            dock_df = dock_df.drop(columns=dock_df.columns[dock_df.columns.isna()])
            dock_df.columns = [str(col).strip() for col in dock_df.columns]
            dock_df = dock_df[~dock_df.apply(lambda row: row.astype(str).str.contains('^Unnamed', na=False)).any(axis=1)]
            dock_df['lcto'] = ['dock_{:02d}'.format(i) for i in range(1, len(dock_df)+1)]

            if 'Id Tipo Transacao' in dock_df.columns and 'Valor' in dock_df.columns:
                dock_df['Valor'] = pd.to_numeric(dock_df['Valor'], errors='coerce')
                dock_df['Valor'] = np.where(
                    dock_df['Id Tipo Transacao'].isin([30224, 30350]),
                    -abs(dock_df['Valor']),
                    abs(dock_df['Valor'])
                )
        except Exception as e:
            return f"<h3>Erro ao processar o Dock: {e}</h3>"
        

        ### 2. Leitura do Matera com tratamento
        try:
            matera_df = pd.read_csv(matera_file, delimiter=';')
            matera_df = matera_df[~matera_df.apply(lambda row: row.astype(str).str.contains('^Unnamed', na=False)).any(axis=1)]  
            matera_df = matera_df.loc[:, ~matera_df.columns.str.contains('^Unnamed', na=False)]
            matera_df['lcto'] = ['matera_{:02d}'.format(i) for i in range(1, len(matera_df)+1)]
            matera_df['nVlrLanc'] = matera_df['nVlrLanc'].str.replace(',', '.', regex=False).astype('float64')
            matera_df['sCpf_Cnpj'] = matera_df['sCpf_Cnpj'].astype(str).str.replace(r'[.\-]', '', regex=True)
            matera_df['nVlrLanc'] = np.where(
                matera_df['nHistorico'] == 9001,
                -abs(matera_df['nVlrLanc']),
                abs(matera_df['nVlrLanc'])
            )
            matera_df.rename(columns={'sCpf_Cnpj': 'CPF'}, inplace=True)
        except Exception as e:
            return f"<h3>Erro ao ler o CSV Matera: {e}</h3>"

        ### 3. Leitura do Depara com lógica flexível (Unnamed: 2)
        try:
            depara = pd.read_excel(depara_file, dtype=str)
            start_index = depara[depara['Unnamed: 2'].notna()].index[0]
            depara = depara.iloc[start_index:].reset_index(drop=True)
            depara.columns = depara.iloc[0]
            depara = depara.iloc[1:].reset_index(drop=True)
            depara = depara.drop(columns=depara.columns[depara.columns.isna()])
            depara.columns = [str(col).strip() for col in depara.columns]
            depara = depara[~depara.apply(lambda row: row.astype(str).str.contains('^Unnamed', na=False)).any(axis=1)]
        except Exception as e:
            return f"<h3>Erro ao processar o Depara: {e}</h3>"

        ### Montando as lógicas dos merges aqui
        dock_df = dock_df.merge(
            depara[['Id Conta', 'CPF', 'Nome', 'Status Conta', 'Data Cadastramento']],
            on='Id Conta',
            how='left'
        )

        ### 4. Geração do Excel final
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            dock_df.to_excel(writer, sheet_name='Dock', index=False)
            matera_df.to_excel(writer, sheet_name='Matera', index=False)
            depara.to_excel(writer, sheet_name='Depara', index=False)

            # Ajusta a largura das colunas nas três abas
            def auto_adjust(ws):
                for col_cells in ws.columns:
                    max_length = max(
                        (len(str(cell.value)) if cell.value is not None else 0)
                        for cell in col_cells
                    )
                    col_letter = col_cells[0].column_letter
                    ws.column_dimensions[col_letter].width = max(max_length + 2, 15)

            auto_adjust(writer.sheets['Dock'])
            auto_adjust(writer.sheets['Matera'])
            auto_adjust(writer.sheets['Depara'])

        output.seek(0)

        return send_file(
            output,
            download_name="dock_matera_depara.xlsx",
            as_attachment=True,
        )

    return render_template_string(HTML_PAGE)

if __name__ == '__main__':
    app.run(debug=True)
