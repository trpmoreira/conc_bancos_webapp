from flask import Flask, render_template, request, jsonify, send_file
from ficheiros_bancos import create_and_update_files, meses_pt
from base_dados import get_bank_transactions
import pandas as pd
import os
from io import BytesIO
import io

app = Flask(__name__)

# Ativar modo debug e logging
app.debug = True

# Lista com os nomes dos meses em português
meses_pt = [
    "", "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
]

@app.route('/')
def index():
    return render_template('index.html', meses=meses_pt[1:])

@app.route('/processar', methods=['POST'])
def processar():
    print("Requisição recebida") # Log para debug
    try:
        data = request.get_json()
        print(f"Dados recebidos: {data}") # Log para debug

        mes = int(data['mes'])
        print(f"Processando mês: {mes}") # Log para debug

        # Chama sua função existente
        create_and_update_files(mes)

        # Lê o arquivo de resumo gerado
        mes_nome = meses_pt[mes]
        arquivo_resumo = f"Resumo Conciliação - {mes} - {mes_nome}.xlsx"

        # Lê o arquivo Excel para mostrar na webapp
        df_resumo = pd.read_excel(arquivo_resumo, sheet_name='Resumo')
        df_docs_invalidos = pd.read_excel(arquivo_resumo, sheet_name='Docs Inválidos')

        return jsonify({
            "success": True,
            "message": "Análise concluída com sucesso",
            "arquivo": arquivo_resumo,
            "resumo": df_resumo.to_dict('records'),
            "docs_invalidos": df_docs_invalidos.to_dict('records')
        })

    except Exception as e:
        print(f"Erro durante o processamento: {str(e)}") # Log para debug
        return jsonify({
            "success": False,
            "message": f"Erro: {str(e)}"
        }), 500

@app.route('/download/<filename>')
def download(filename):
    try:
        return send_file(filename, as_attachment=True)
    except Exception as e:
        return str(e)

@app.route('/transacoes')
def transacoes_page():
    return render_template('transacoes.html', meses=meses_pt[1:])

@app.route('/buscar-transacoes', methods=['POST'])
def buscar_transacoes():
    try:
        data = request.get_json()
        conta = data['conta']
        mes = int(data['mes'])
        ano = int(data.get('ano', 2024))

        # Busca as transações
        df = get_bank_transactions(conta, mes, ano)

        if df is None:
            return jsonify({
                "success": False,
                "message": "Erro ao buscar transações"
            }), 500

        return jsonify({
            "success": True,
            "data": df.to_dict('records'),
            "total": len(df)
        })

    except Exception as e:
        print(f"Erro: {str(e)}")
        return jsonify({
            "success": False,
            "message": f"Erro: {str(e)}"
        }), 500

@app.route('/download-transacoes', methods=['POST'])
def download_transacoes():
    try:
        data = request.get_json()
        conta = data.get('conta')
        mes = int(data.get('mes'))
        ano = int(data.get('ano'))

        # Buscar os dados usando a função existente
        df = get_bank_transactions(conta, mes, ano)

        if df is None:
            return jsonify({'success': False, 'message': 'Nenhum dado encontrado'}), 404

        # Criar um buffer na memória para salvar o Excel
        output = io.BytesIO()

        # Salvar o DataFrame como Excel no buffer
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='Transações', index=False)

        # Preparar o buffer para leitura
        output.seek(0)

        # Retornar o arquivo
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f'transacoes_{conta}_{mes}_{ano}.xlsx'
        )

    except Exception as e:
        print(f"Erro ao gerar arquivo: {str(e)}")
        return jsonify({'success': False, 'message': str(e)}), 500

@app.context_processor
def utility_processor():
    return dict(request=request)

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)