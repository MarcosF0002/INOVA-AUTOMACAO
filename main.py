import os
import json
import time
import gspread
from google.oauth2.service_account import Credentials
from criarHTML import processa_aba_gera_html
from atualizador_WP import atualizar_pagina_wp
from pitchs import gerar_html_pitchs_via_api
from criaHTMLPais import gerar_html_pais
from criarHTML_3col import gerar_html_3COL

# =========================================
# Configuração e Autenticação
# =========================================
google_json = os.environ.get("GOOGLE_JSON")
if not google_json:
    raise ValueError("❌ O secret GOOGLE_JSON não está definido!")

GSHEET_KEY = os.environ.get("GSHEETS_KEY")
if not GSHEET_KEY:
    raise ValueError("❌ O secret GSHEETS_KEY não está definido!")

print(f"🔑 ID da planilha: {GSHEET_KEY}")

try:
    # Carrega credenciais
    creds_dict = json.loads(google_json)
    
    # Define escopos
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive",
        "https://www.googleapis.com/auth/spreadsheets"
    ]
    
    # Autenticação correta
    creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
    client = gspread.authorize(creds)
    
    print("✅ Autenticação com Google Sheets realizada com sucesso!")
    
except Exception as e:
    raise Exception(f"❌ Erro na autenticação: {e}")

# =========================================
# Acesso à Planilha (com verificações)
# =========================================
try:
    # Abre a planilha
    spreadsheet = client.open_by_key(GSHEET_KEY)
    print(f"✅ Planilha encontrada: {spreadsheet.title}")
    
    # Lista todas as worksheets disponíveis
    todas_worksheets = [ws.title for ws in spreadsheet.worksheets()]
    print(f"📋 Worksheets disponíveis: {todas_worksheets}")
    
    # Verifica se a worksheet "CHECAR ABAS" existe
    if "CHECAR ABAS" not in todas_worksheets:
        raise Exception(f"Worksheet 'CHECAR ABAS' não encontrada. Worksheets disponíveis: {todas_worksheets}")
    
    # Acessa a worksheet
    sheet = spreadsheet.worksheet("CHECAR ABAS")
    print("✅ Worksheet 'CHECAR ABAS' acessada com sucesso!")
    
except gspread.exceptions.SpreadsheetNotFound:
    raise Exception("❌ Planilha não encontrada! Verifique:")
    print("1. 🔗 A chave GSHEETS_KEY está correta?")
    print("2. 👥 A planilha foi compartilhada com o service account?")
    print(f"3. 📧 E-mail do service account: {creds.service_account_email}")
except Exception as e:
    raise Exception(f"❌ Erro ao acessar planilha: {e}")

# =========================================
# Leitura das Abas Selecionadas
# =========================================
try:
    # Pega valores da coluna A a partir da linha 2
    abas_selecionadas = sheet.col_values(1)[1:]  # ignora a primeira linha (cabeçalho)
    abas_selecionadas = [aba.strip() for aba in abas_selecionadas if aba.strip()]
    
    print(f"✅ Abas que serão atualizadas: {abas_selecionadas}")
    
    if not abas_selecionadas:
        print("⚠️  Nenhuma aba selecionada para atualização!")
        exit(0)
        
except Exception as e:
    raise Exception(f"❌ Erro ao ler abas selecionadas: {e}")

# =========================================
# Mapeamento de Links (mantido igual)
# =========================================
abas_links = {
    "DEEPTECHS": "https://inova.ufpr.br/biotechs/",
    "CONSTRUTECHS E PROPTECHS": "https://inova.ufpr.br/construtechs-e-proptechs/",
    "EDTECHS": "https://inova.ufpr.br/edtechs/",
    "ENERGYTECHS": "https://inova.ufpr.br/energytechs/",
    "FINTECHS": "https://inova.ufpr.br/fintechs/",
    "FOODTECHS": "https://inova.ufpr.br/foodtechs/",
    "GOVTECHS": "https://inova.ufpr.br/govtechs/",
    "GREENTECHS": "https://inova.ufpr.br/greentechs/",
    "HEALTHTECHS": "https://inova.ufpr.br/health-tech/",
    "INDTECHS": "https://inova.ufpr.br/indtechs/",
    "LOGTECHS": "https://inova.ufpr.br/logtechs/",
    "MARTECHS": "https://inova.ufpr.br/martechs/",
    "MOBITECHS": "https://inova.ufpr.br/mobitechs/",
    "RETAILTECHS": "https://inova.ufpr.br/retailtechs-2/",
    "SOCIALTECHS": "https://inova.ufpr.br/socialtechs/",
    "TECHS": "https://inova.ufpr.br/techs/",
    "WATERTECHS": "https://inova.ufpr.br/watertechs/",
    "LAWTECHS E LEGALTECHS": "https://inova.ufpr.br/lawtechs-e-legaltechs/",
    "PETTECHS": "https://inova.ufpr.br/pet-techs/",
    "TESTE": "https://inova.ufpr.br/teste/",
    "ACELERADORAS E INCUBADORAS": "https://inova.ufpr.br/aceleradoras-incubadoras/",
    "ASSOCIAÇÕES EMPRESARIAIS": "https://inova.ufpr.br/associacao-empresarial/",
    "FINANCIAMENTO A INOVAÇÃO": "https://inova.ufpr.br/financiamento-inovacao/",
    "HUBS E ECOSSISTEMAS": "https://inova.ufpr.br/hubs-e-ecossistemas/",
    "INOVAÇÃO NAS UNIVERSIDADES": "https://inova.ufpr.br/inovacao-nas-universidades/",
    "INSTITUTOS E CENTROS DE PESQUISA": "https://inova.ufpr.br/institutos-de-pesquisa/",
    "PARQUES CIENTÍFICOS": "https://inova.ufpr.br/parques-tecnologicos/",
    "PERÍODICOS CIENTÍFICOS": "https://inova.ufpr.br/periodicos-cientificos/",
    "POLÍTICAS DE INOVAÇÃO": "https://inova.ufpr.br/politicas-de-inovacao/",
    "PROPRIEDADE INTELECTUAL": "https://inova.ufpr.br/1234-2/",
    "VÍDEOS E PODCASTS": "https://inova.ufpr.br/cursos-e-podcasts-de-empreendedorismo/",
    "PITCHS DE STARTUPS": "https://inova.ufpr.br/pitchs-de-startups-incubadoras-e-aceleradoras/"
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
# Processamento em Lotes (com melhor tratamento de erros)
# =========================================
tamanho_lote = 5
erros = []
sucessos = []

for i in range(0, len(abas_selecionadas), tamanho_lote):
    lote = abas_selecionadas[i:i+tamanho_lote]
    print(f"\n➡️ Processando lote {i//tamanho_lote + 1}: {lote}")

    for aba in lote:
        try:
            print(f"\n🔄 Processando aba: {aba}")
            
            # Verifica se a aba existe no mapeamento
            if aba not in abas_links:
                print(f"❌ Aba '{aba}' não encontrada no mapeamento de links!")
                erros.append(f"{aba}: Link não mapeado")
                continue

            # Gera HTML baseado no tipo de aba
            if aba.upper() == "PITCHS DE STARTUPS":
                html = gerar_html_pitchs_via_api()
            elif aba.upper() == "VÍDEOS E PODCASTS":
                html = gerar_html_3COL(aba)
            elif aba.upper() in [a.upper() for a in abas_pais]:
                html = gerar_html_pais(aba)
            else:
                html = processa_aba_gera_html(aba)

            if html is None:
                print(f"❌ HTML retornado como None para aba: {aba}")
                erros.append(f"{aba}: Erro ao gerar HTML")
                continue

            # Atualiza página WordPress
            resposta = atualizar_pagina_wp(abas_links[aba], html)

            if not resposta:
                print(f"❌ Falha ao atualizar página: {abas_links[aba]}")
                erros.append(f"{aba}: Falha ao atualizar página WordPress")
            else:
                print(f"✅ Página atualizada com sucesso: {abas_links[aba]}")
                sucessos.append(aba)

        except Exception as e:
            error_msg = f"{aba}: {str(e)}"
            print(f"❌ Erro inesperado: {error_msg}")
            erros.append(error_msg)

    # Pausa entre lotes (se não for o último lote)
    if i + tamanho_lote < len(abas_selecionadas):
        print(f"\n⏱️  Pausa de 60 segundos antes do próximo lote...")
        time.sleep(60)

# =========================================
# Relatório Final
# =========================================
print(f"\n{'='*50}")
print("📊 RELATÓRIO DE EXECUÇÃO")
print(f"{'='*50}")
print(f"✅ Sucessos: {len(sucessos)}")
print(f"❌ Erros: {len(erros)}")
print(f"📋 Total processado: {len(abas_selecionadas)}")

if sucessos:
    print(f"\n✅ Abas atualizadas com sucesso: {sucessos}")

if erros:
    print(f"\n❌ Erros encontrados:")
    for e in erros:
        print(f"   - {e}")

if not erros:
    print("\n🎉 Todas as abas selecionadas foram atualizadas com sucesso!")
else:
    print(f"\n⚠️  {len(erros)} erro(s) ocorreram durante o processamento.")
