import pandas as pd
from datetime import datetime
import pyperclip
from conexao_api import client
import gspread.exceptions
from atualizador_WP import atualizar_pagina_wp
import re
from io import StringIO

def get_video_id(link):
    match = re.search(r'(?:v=|\/)([0-9A-Za-z_-]{11})(?:[&?]|$)', link)
    return match.group(1) if match else ''

def processa_pitchs_com_historico():
    try:
        planilha = client.open("PORTAL DA INOVAÇÃO E STARTUPS")
        aba = planilha.worksheet("PITCHS DE STARTUPS")

        dados_raw = aba.get_all_values()
        if not dados_raw:
            print("Aba PITCHS DE STARTUPS vazia.")
            return None

        cabecalho = dados_raw[0]
        data_linhas = dados_raw[1:]

        # Mapeamento de colunas esperadas
        col_map = {nome.upper().strip(): i for i, nome in enumerate(cabecalho)}
        required = ["NOME", "CATEGORIA", "LINK", "STATUS"]
        if not all(col in col_map for col in required):
            print("Colunas obrigatórias ausentes em PITCHS DE STARTUPS.")
            return None

        idx_nome = col_map["NOME"]
        idx_categoria = col_map["CATEGORIA"]
        idx_link = col_map["LINK"]
        idx_status = col_map["STATUS"]

        novas_linhas_aba = [cabecalho]
        entradas, saidas = [], []

        for linha in data_linhas:
            if len(linha) <= idx_status:
                novas_linhas_aba.append(linha)
                continue

            status = str(linha[idx_status]).strip().upper()
            nome = linha[idx_nome] if len(linha) > idx_nome else ""
            categoria = linha[idx_categoria] if len(linha) > idx_categoria else ""
            link = linha[idx_link] if len(linha) > idx_link else ""
            timestamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

            if status == "ADICIONAR AO SITE":
                entradas.append(["ENTRADA", "PITCHS DE STARTUPS", timestamp, nome, categoria, link])
                linha_mod = list(linha)
                linha_mod[idx_status] = "ADICIONADO AO SITE"
                novas_linhas_aba.append(linha_mod)
            elif status == "REMOVER":
                saidas.append(["SAÍDA", "PITCHS DE STARTUPS", timestamp, nome, categoria, link])
            else:
                novas_linhas_aba.append(linha)

        # Atualiza aba original
        aba.clear()
        if novas_linhas_aba:
            aba.update(f"A1:{chr(65 + len(cabecalho) - 1)}{len(novas_linhas_aba)}", novas_linhas_aba)

        # Atualiza aba HISTÓRICO
        try:
            aba_historico = planilha.worksheet("HISTÓRICO")
        except gspread.exceptions.WorksheetNotFound:
            aba_historico = planilha.add_worksheet("HISTÓRICO", rows=1, cols=10)
            aba_historico.append_row(["TIPO OPERAÇÃO", "ABA", "DATA", "NOME", "CATEGORIA", "LINK"])

        historico = entradas + saidas
        if historico:
            aba_historico.append_rows(historico)
            print(f"{len(historico)} linhas adicionadas à aba HISTÓRICO.")
        return True

    except Exception as e:
        print(f"Erro ao processar histórico: {e}")
        return None

def gerar_html_pitchs(df):
    if df.empty:
        print("DataFrame vazio, nada para gerar.")
        return None

    html = StringIO()
    html.write("<!-- COMECA ATUALIZAR DAQUI -->\n")
    
    for _, row in df.iterrows():
        video_id = get_video_id(str(row.get("LINK", "")))
        if not video_id:
            continue

        html.write(f"""
      
        <tr class="organizationRow"
            data-categoria="{row.get('CATEGORIA', '')}"
            data-instituicao="{row.get('INSTITUIÇÃO', '')}"
            data-segmento="{row.get('SEGMENTO', '')}">
            <td scope="row" style="text-align: center; position: relative; width: 600px;">
                <div style="color: darkblue; font-weight: bold; margin-bottom: 5px;">
                    <a href="{row.get('LINK', '')}" target="_blank">{row.get('NOME', '')}</a>
                </div>
                <a href="#" onclick="openFullscreen('https://www.youtube.com/embed/{video_id}')"
                    style="display: inline-block; position: relative;">
                    <div class="tooltip-container">
                        <img src="https://img.youtube.com/vi/{video_id}/0.jpg" alt="Miniatura do vídeo"
                            style="width: 180px; cursor: pointer; display: block;">
                        <div class="play-button">▶</div>
                        <span class="tooltip-inner">{row.get('CONTEÚDO BALÃO', '')}</span>
                    </div>
                </a>
            </td>
            <td style="text-align: center; vertical-align: middle;">{row.get('CATEGORIA', '')}</td>
            <td style="text-align: center; vertical-align: middle;">{row.get('INSTITUIÇÃO', '')}</td>
            <td style="text-align: center; vertical-align: middle;">{row.get('SEGMENTO', '')}</td>
        </tr>


        """)

    html.write("</tbody></table>\n") # Fecha a tabela
    
    result_html = html.getvalue()
    pyperclip.copy(result_html)
    print("HTML copiado para a área de transferência.")
    return result_html

def gerar_html_pitchs_via_api():
    try:
        planilha = client.open("PORTAL DA INOVAÇÃO E STARTUPS")
        aba = planilha.worksheet("PITCHS DE STARTUPS")
        dados = aba.get_all_records()
        df = pd.DataFrame(dados)

        processa_pitchs_com_historico()
        return gerar_html_pitchs(df)

    except Exception as e:
        print(f"Erro ao gerar HTML dos pitchs: {e}")
        return None
