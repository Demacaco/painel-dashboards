import pandas as pd
import os
from datetime import datetime

# Configuração dos arquivos
ARQUIVO_DADOS = 'Relatorio_NFe.xlsx'
ARQUIVO_SAIDA = 'painel.dash.html'

print("--- INICIANDO GERADOR DE DASHBOARD ---")

# 1. Verificação do Arquivo
if not os.path.exists(ARQUIVO_DADOS):
    print(f"ERRO CRÍTICO: O arquivo '{ARQUIVO_DADOS}' não foi encontrado na pasta.")
    print("Confira se o nome está exato (letras maiúsculas/minúsculas importam).")
    input("Pressione ENTER para sair...")
    exit()

print(f"1. Arquivo '{ARQUIVO_DADOS}' encontrado! Carregando dados...")

try:
    # 2. Carregar Dados (Tenta ler Excel)
    df = pd.read_excel(ARQUIVO_DADOS)
    print(f"2. Dados carregados com sucesso! Total de linhas: {len(df)}")
    
    # 3. Processamento Básico (KPIs)
    # Tenta adivinhar colunas comuns se não achar as exatas
    cols = [c.lower() for c in df.columns]
    
    # Valor Total do Estoque (procura coluna de valor/preço)
    valor_total = 0
    col_valor = next((c for c in df.columns if 'valor' in c.lower() or 'total' in c.lower() or 'custo' in c.lower()), None)
    if col_valor:
        valor_total = df[col_valor].sum()
    
    # Quantidade Total
    qtd_total = 0
    col_qtd = next((c for c in df.columns if 'qtd' in c.lower() or 'quant' in c.lower() or 'saldo' in c.lower()), None)
    if col_qtd:
        qtd_total = df[col_qtd].sum()
    else:
        qtd_total = len(df) # Se não tiver coluna qtd, conta as linhas

    # Data de Atualização
    data_hoje = datetime.now().strftime("%d/%m/%Y %H:%M")

    # 4. Gerar HTML
    html_content = f"""
    <!DOCTYPE html>
    <html lang="pt-br">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Relatório Fechamento 2025</title>
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
        <style>
            body {{ font-family: 'Segoe UI', sans-serif; background-color: #f4f6f9; margin: 0; padding: 20px; }}
            .container {{ max-width: 1200px; margin: 0 auto; }}
            .header {{ background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%); color: white; padding: 30px; border-radius: 15px; margin-bottom: 30px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }}
            .header h1 {{ margin: 0; font-size: 2.5rem; }}
            .header p {{ opacity: 0.8; margin-top: 5px; }}
            
            .kpi-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 20px; margin-bottom: 30px; }}
            .kpi-card {{ background: white; padding: 25px; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); border-left: 5px solid #3498db; }}
            .kpi-card h3 {{ margin: 0 0 10px 0; color: #7f8c8d; font-size: 0.9rem; text-transform: uppercase; }}
            .kpi-card .value {{ font-size: 2rem; font-weight: bold; color: #2c3e50; }}
            .kpi-card i {{ float: right; font-size: 2.5rem; opacity: 0.2; color: #2c3e50; }}

            .table-container {{ background: white; padding: 20px; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); overflow-x: auto; }}
            table {{ width: 100%; border-collapse: collapse; }}
            th {{ background-color: #f8f9fa; padding: 15px; text-align: left; font-weight: 600; color: #2c3e50; border-bottom: 2px solid #e9ecef; }}
            td {{ padding: 12px 15px; border-bottom: 1px solid #e9ecef; color: #495057; }}
            tr:hover {{ background-color: #f8f9fa; }}

            .btn-back {{ display: inline-block; margin-bottom: 20px; color: #3498db; text-decoration: none; font-weight: bold; }}
            .btn-back:hover {{ text-decoration: underline; }}
        </style>
    </head>
    <body>
        <div class="container">
            <a href="index.html" class="btn-back"><i class="fas fa-arrow-left"></i> Voltar ao Painel Principal</a>
            
            <div class="header">
                <i class="fas fa-chart-line" style="float: right; font-size: 4rem; opacity: 0.2;"></i>
                <h1>Fechamento Anual 2025</h1>
                <p>Relatório gerado em: {data_hoje}</p>
            </div>

            <div class="kpi-grid">
                <div class="kpi-card" style="border-color: #3498db;">
                    <i class="fas fa-boxes"></i>
                    <h3>Itens em Estoque</h3>
                    <div class="value">{qtd_total:,.0f}</div>
                </div>
                <div class="kpi-card" style="border-color: #2ecc71;">
                    <i class="fas fa-money-bill-wave"></i>
                    <h3>Valor Total Estimado</h3>
                    <div class="value">R$ {valor_total:,.2f}</div>
                </div>
                <div class="kpi-card" style="border-color: #e74c3c;">
                    <i class="fas fa-file-invoice"></i>
                    <h3>Total de Linhas</h3>
                    <div class="value">{len(df)}</div>
                </div>
            </div>

            <div class="table-container">
                <h3><i class="fas fa-table"></i> Prévia dos Dados (Top 50 itens)</h3>
                {df.head(50).to_html(index=False, border=0, classes='table')}
            </div>
            
            <div style="text-align: center; margin-top: 40px; color: #95a5a6; font-size: 0.8rem;">
                Painel gerado automaticamente via Python
            </div>
        </div>
    </body>
    </html>
    """

    # 5. Salvar Arquivo
    with open(ARQUIVO_SAIDA, "w", encoding="utf-8") as f:
        f.write(html_content)

    print(f"3. SUCESSO! Relatório '{ARQUIVO_SAIDA}' gerado na pasta.")

except Exception as e:
    print(f"\nERRO AO PROCESSAR O EXCEL: {e}")
    print("Verifique se você tem as bibliotecas instaladas (pip install pandas openpyxl)")

print("--- FIM DO PROCESSO ---")