import os
import pandas as pd
import pyperclip

def processa_aba_gera_html(aba, 
                           sheet_xlsx_url="https://docs.google.com/spreadsheets/d/e/2PACX-1vTbpI-FEyM9QOXLJOxLnfVxxyMKUoDwFcBjDSyhko3GA8cIjpxxfkiYWpMThUuvSPrAnQ3L31Wm2vqm/pub?output=xlsx",
                           output_directory=r"C:\Users\marco\OneDrive\Área de Trabalho\Economia\INOVA\tabelas-atualizadas"):

    try:
        data = pd.read_excel(sheet_xlsx_url, sheet_name=aba)
    except Exception as e:
        print(f"Erro ao ler a planilha XLSX: {e}")
        return None

    if data is None or data.empty:
        print("Nenhum dado disponível para gerar o arquivo HTML.")
        return None

    def formatar_nome(nome):
        preposicoes = {'de', 'da', 'do', 'das', 'dos', 'em', 'no', 'na', 'nos', 'nas', 'a', 'o', 'e', 'com', 'para', 'por', 'sob', 'sem'}

        if not isinstance(nome, str) or not nome:
            return nome

        palavras = nome.split()
        nome_formatado = []

        for i, palavra in enumerate(palavras):
            if palavra.lower() in preposicoes:
                nome_formatado.append(palavra.lower())
            elif i == 0:
                # Primeira palavra: forçar só a primeira letra para maiúscula, mantendo o resto
                nome_formatado.append(palavra[0].upper() + palavra[1:] if palavra else '')
            else:
                nome_formatado.append(palavra)

        return ' '.join(nome_formatado)



    data['NOME'] = data['NOME'].apply(formatar_nome)
    data = data.sort_values(by='NOME', key=lambda col: col.str.lower())

    # Detecta se a 5ª coluna (índice 4) é CIDADE ou PAÍS
    colunas = list(data.columns)
    if len(colunas) < 5:
        print("A planilha não tem pelo menos cinco colunas.")
        return None

    quinta_coluna_nome = colunas[4]  # índice 4 = quinta coluna
    seletor_id = "cidadeSelect" if quinta_coluna_nome.upper() == "CIDADE" else "paisSelect"
    seletor_label = "Todas Cidades" if quinta_coluna_nome.upper() == "CIDADE" else "Todos os Países"
    data_attr = "cidade" if quinta_coluna_nome.upper() == "CIDADE" else "pais"

    def generate_html_table(data):
        html = f"""
<table class="table" id="organization_table">
<thead>
<tr>
<th scope="col"><p>Organização</p></th>
<th scope="col"><select id="ufSelect" onchange="filterTable()"><option value="">Todos Estados</option></select></th>
<th scope="col"><select id="{seletor_id}" onchange="filterTable()"><option value="">{seletor_label}</option></select></th>
<th scope="col"><select id="categoriaSelect" onchange="filterTable()"><option value="">Todas Categorias</option></select></th>
</tr>
</thead>
<tbody>
"""
        for _, row in data.iterrows():
            link = row['LINK'] if pd.notnull(row['LINK']) else '#'
            if not str(link).startswith(('http://', 'https://')):
                link = 'http://' + str(link)
            nome = row['NOME'] if pd.notnull(row['NOME']) else ''
            uf = row['UF'] if pd.notnull(row['UF']) else ''
            valor_quinta_coluna = row[quinta_coluna_nome] if pd.notnull(row[quinta_coluna_nome]) else ''
            categoria = row['CATEGORIA'] if pd.notnull(row['CATEGORIA']) else ''
            conteudo_balao = row['CONTEÚDO BALÃO'] if pd.notnull(row['CONTEÚDO BALÃO']) else ''
            html += f"""
<tr class="organizationRow" data-uf="{uf}" data-{data_attr}="{valor_quinta_coluna}" data-categoria="{categoria}">
<td scope="row">
<span data-bs-placement="bottom" data-bs-toggle="tooltip" title="{conteudo_balao}">
<a href="{link}" rel="noopener noreferrer" target="_blank">{nome}</a>
</span>
</td>
<td>{uf}</td>
<td>{valor_quinta_coluna}</td>
<td>{categoria}</td>
</tr>
"""
        html += """
</tbody>
</table>
"""
        return html

    html_table = generate_html_table(data)
    total_organizacoes = len(data)

    html = f"""
<!-- COMECA ATUALIZAR DAQUI -->
<div class="p-2 mr-2" id="count">
<p><b>Total de organizações:</b> {total_organizacoes}</p>
</div>
""" + html_table

    try:
        if not os.path.exists(output_directory):
            os.makedirs(output_directory)
        output_path = os.path.join(output_directory, f"{aba}.html")
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(html)
        print(f"Arquivo HTML '{output_path}' criado com sucesso.")
    except Exception as e:
        print(f"Erro ao escrever o arquivo HTML: {e}")

    try:
        pyperclip.copy(html)
        print("Código HTML copiado para a área de transferência.")
    except Exception as e:
        print(f"Erro ao copiar para área de transferência: {e}")

    return html
