import pandas as pd
import json
import glob

def carregar_dados():
    # Tenta ler os arquivos. Ajuste o nome se necessário.
    try:
        # Se você estiver usando os CSVs exportados:
        df_notas = pd.read_csv('Relatorio_NFe.xlsx - Lista_Notas.csv')
        df_itens = pd.read_csv('Relatorio_NFe.xlsx - Itens_Detalhados.csv')
        
        # Se for ler direto do Excel (precisa instalar openpyxl: pip install openpyxl)
        # df_notas = pd.read_excel('Relatorio_NFe.xlsx', sheet_name='Lista_Notas')
        # df_itens = pd.read_excel('Relatorio_NFe.xlsx', sheet_name='Itens_Detalhados')
    except FileNotFoundError:
        print("Erro: Arquivos de dados não encontrados.")
        return None

    return df_notas, df_itens

def gerar_json_fechamento():
    dados = carregar_dados()
    if not dados: return

    df_notas, df_itens = dados

    # --- Tratamento de Dados ---
    # Converter valor para numérico (caso venha como string com virgula)
    # df_notas['Valor_Total'] = df_notas['Valor_Total'].astype(str).str.replace(',', '.').astype(float)
    
    # Converter Data
    if 'Data' in df_notas.columns:
        df_notas['Data'] = pd.to_datetime(df_notas['Data'])
    elif 'Data Emissao' in df_notas.columns:
         df_notas['Data'] = pd.to_datetime(df_notas['Data Emissao'])

    # --- 1. KPIs ---
    faturamento = df_notas['Valor_Total'].sum()
    qtd_notas = df_notas['Chave'].nunique() if 'Chave' in df_notas.columns else len(df_notas)
    qtd_pecas = df_itens['Quantidade'].sum()
    ticket_medio = faturamento / qtd_notas if qtd_notas > 0 else 0

    # --- 2. Gráfico Mensal (Janeiro a Dezembro) ---
    # Cria um array de 12 posições zeradas
    vendas_mes = [0] * 12
    
    # Agrupa por mês (1=Jan, 12=Dez) e preenche o array (índice 0=Jan)
    grupo_mes = df_notas.groupby(df_notas['Data'].dt.month)['Valor_Total'].sum()
    for mes, valor in grupo_mes.items():
        if 1 <= mes <= 12:
            vendas_mes[int(mes)-1] = round(valor, 2)

    # --- 3. Top 5 Produtos (Curva A) ---
    top_prods = df_itens.groupby(['Codigo', 'Descricao'])['Quantidade'].sum().reset_index()
    top_prods = top_prods.sort_values(by='Quantidade', ascending=False).head(5)
    
    lista_produtos = []
    for _, row in top_prods.iterrows():
        lista_produtos.append({
            "Codigo": str(row['Codigo']),
            "Descricao": row['Descricao'],
            "Quantidade": int(row['Quantidade'])
        })

    # --- JSON Final ---
    payload = {
        "kpis": {
            "faturamento": round(faturamento, 2),
            "notas": int(qtd_notas),
            "pecas": int(qtd_pecas),
            "ticket": round(ticket_medio, 2)
        },
        "graficoMensal": vendas_mes,
        "topProdutos": lista_produtos
    }

    print("\n=== COPIE DAQUI PARA BAIXO E COLE NO SEU HTML ===\n")
    print(f"const dadosReais = {json.dumps(payload, indent=4, ensure_ascii=False)};")
    print("\n=================================================\n")

if __name__ == "__main__":
    gerar_json_fechamento()