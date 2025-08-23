import pandas as pd
from criarHTML import processa_aba_gera_html
from atualizador_WP import atualizar_pagina_wp
from pitchs import gerar_html_pitchs_via_api
from criaHTMLPais import gerar_html_pais
from criarHTML_3col import gerar_html_3COL

# Lê apenas a coluna A a partir da linha 2 da aba 'CHECAR ABAS'
df_checagem = pd.read_excel("links_startups.xlsx", sheet_name="CHECAR ABAS", usecols="A", skiprows=1)
abas_selecionadas = df_checagem.iloc[:, 0].dropna().tolist()  # remove valores nulos

print("Abas que serão atualizadas:", abas_selecionadas)

# Lê os links de todas as abas (como antes)
links_df = pd.read_excel("links_startups.xlsx")
abas_links = dict(zip(links_df['ABA'], links_df['LINK']))

abas_pais = [
    "ASSOCIAÇÕES EMPRESARIAIS",
    "FINANCIAMENTO A INOVAÇÃO",
    "HUBS E ECOSSISTEMAS",
    "INSTITUTOS E GRUPOS DE PESQUISA",
    "POLÍTICAS DE INOVAÇÃO",
    "PROPRIEDADE INTELECTUAL",
    "TESTE"
]

erros = []

tamanho_lote = 5  # número de abas por lote

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
        time.sleep(60)  # pausa de 60 segundos

if not erros:
    print("\nTodas as abas selecionadas foram atualizadas com sucesso.")
else:
    print("\nAlguns erros ocorreram:")
    for e in erros:
        print("-", e)
