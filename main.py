import os
import json
import time
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from criarHTML import processa_aba_gera_html
from atualizador_WP import atualizar_pagina_wp
from pitchs import gerar_html_pitchs_via_api
from criaHTMLPais import gerar_html_pais
from criarHTML_3col import gerar_html_3COL

# =========================================
# Conexão com Google Sheets
# =========================================
google_json = os.environ.get("GOOGLE_JSON")
if not google_json:
    raise ValueError("O secret GSHEETS_CREDENTIALS_JSON não está definido!")

creds_dict = json.loads(google_json)
scope = ["https://spreadsheets.google.com/feeds",
         "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
client = gspread.authorize(creds)

# Planilha e aba
GSHEET_KEY = os.environ.get("GSHEETS_KEY")  # secret com ID da planilha
sheet = client.open_by_key(GSHEET_KEY).worksheet("CHECAR ABAS")

# Pega valores da coluna A a partir da linha 2
abas_selecionadas = sheet.col_values(1)[1:]  # ignora a primeira linha (cabeçalho)
abas_selecionadas = [aba for aba in abas_selecionadas if aba.strip()]
print("Abas que serão atualizadas:", abas_selecionadas)

# =========================================
# Dicionário com links de todas as abas
# =========================================
# Você pode ter outra aba no Google ou outro arquivo com links
# Aqui assumimos que você tem todos os links já mapeados em um dicionário
# exemplo:
abas_links = {
    "TESTE": "https://inova.ufpr.br/teste",
    "ASSOCIAÇÕES EMPRESARIAIS": "https://inova.ufpr.br/associacoes",
    # ... adicione todas as abas que precisar
}

abas_pais = [
    "ASSOCIAÇÕES EMPRESARIAIS",
    "FINANCIAMENTO A INOVAÇÃO",
    "HUBS E ECOSSISTEMAS",
    "INSTITUTOS E GRUPOS DE PESQUISA",
    "POLÍTICAS DE INOVAÇÃO",
    "PROPRIEDADE INTELECTUAL",
    "TESTE"
]

# =========================================
# Atualização em lotes de 5 abas
# =========================================
tamanho_lote = 5
erros = []

for i in range(0, len(abas_selecionadas), tamanho_lote):
    lote = abas_selecionadas[i:i+tamanho_lote]
    print(f"\n➡️ Processando lote {i//tamanho_lote + 1}: {lote}")

    for aba in lote:
        try:
            print(f"\nProcessando aba: {aba}")

            if aba.strip().upper() == "PITCHS DE STARTUPS":
                html = gerar_html_pitchs_via_api()
            elif aba.strip().upper() == "VÍDEOS E PODCASTS":
                html = gerar_html_3COL(aba)
            elif aba.upper() in abas_pais:
                html = gerar_html_pais(aba)
            else:
                html = processa_aba_gera_html(aba)

            if html is None:
                print(f"❌ HTML retornado como None para aba: {aba}")
                erros.append(f"{aba}: Erro ao gerar HTML.")
                continue

            resposta = atualizar_pagina_wp(abas_links[aba], html)

            if not resposta:
                print(f"❌ Falha ao atualizar página: {abas_links[aba]}")
                erros.append(f"{aba}: Falha ao atualizar a página.")
            else:
                print(f"✅ Página atualizada com sucesso: {abas_links[aba]}")

        except Exception as e:
            erros.append(f"{aba}: {str(e)}")

    if i + tamanho_lote < len(abas_selecionadas):
        print("\n⏱️ Pausa de 1 minuto antes do próximo lote...")
        time.sleep(60)

if not erros:
    print("\nTodas as abas selecionadas foram atualizadas com sucesso.")
else:
    print("\nAlguns erros ocorreram:")
    for e in erros:
        print("-", e)
