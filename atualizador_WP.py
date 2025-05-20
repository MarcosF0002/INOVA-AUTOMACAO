import requests
from urllib.parse import urlparse
from dotenv import load_dotenv
import os
import re
from requests.auth import HTTPBasicAuth

def atualizar_pagina_wp(pagina_url, nova_tabela_html):
    slug = urlparse(pagina_url).path.strip('/')

    search_url = "https://inova.ufpr.br/wp-json/wp/v2/pages"
    headers = {
        "User-Agent": "Mozilla/5.0",
        "Content-Type": "application/json"
    }

    # Buscar a página pelo slug
    resp = requests.get(search_url, params={'slug': slug}, headers=headers)
    print("Código de status da busca:", resp.status_code)
    if resp.status_code != 200:
        print("Erro ao buscar página:", resp.text)
        return False

    pages = resp.json()
    if not pages:
        print(f"Nenhuma página encontrada com slug '{slug}'")
        return False

    page_id = pages[0]['id']
    print(f"ID da página encontrada: {page_id}")

    page_url = f"https://inova.ufpr.br/wp-json/wp/v2/pages/{page_id}?context=edit"

    load_dotenv()
    WP_USER = os.getenv("WP_USER")
    WP_APP_PASSWORD = os.getenv("WP_APP_PASSWORD")

    # Obter conteúdo com contexto de edição (necessário para acessar 'raw')
    resp_get = requests.get(
        page_url,
        auth=HTTPBasicAuth(WP_USER, WP_APP_PASSWORD),
        headers=headers
    )

    if resp_get.status_code != 200:
        print("Erro ao obter conteúdo:", resp_get.text)
        return False

    page_data = resp_get.json()
    conteudo = page_data.get("content", {}).get("raw")

    if not conteudo:
        print("Conteúdo 'raw' não encontrado. Verifique permissões do usuário.")
        return False

    # COMEÇA ATUALIZAR DAQUI
    # Substitui todo o bloco do comentário até o fechamento da tabela
    pattern = r'<!-- COMECA ATUALIZAR DAQUI -->.*?</table>'

    novo_conteudo, count = re.subn(
        pattern,
        f'{nova_tabela_html}',
        conteudo,
        flags=re.DOTALL
    )

    if count == 0:
        print("Aviso: não foi encontrado o marcador '<!-- COMECA ATUALIZAR DAQUI -->' com tabela associada.")
        return False

    # Atualizar a página via API REST
    data = {"content": novo_conteudo}
    resp_update = requests.post(
        f"https://inova.ufpr.br/wp-json/wp/v2/pages/{page_id}",
        auth=HTTPBasicAuth(WP_USER, WP_APP_PASSWORD),
        headers=headers,
        json=data
    )

    print("Código de status da atualização:", resp_update.status_code)
    try:
        print("Resposta da atualização:")
        print(resp_update.json())
    except Exception:
        print("Erro ao interpretar JSON da atualização:", resp_update.text)

    return resp_update.status_code == 200
