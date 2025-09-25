# 🛰️ Atualizador de Páginas do Projeto INOVA – UFPR

Este projeto automatiza a atualização de páginas do site [inova.ufpr.br](https://inova.ufpr.br), que divulga organizações de inovação mapeadas pelo projeto de extensão **Inovação e Desenvolvimento Territorial** da UFPR.

## ✨ Objetivo

A equipe do projeto mantém uma planilha no Google Sheets com diversas abas (startups, aceleradoras, universidades, etc.). Este sistema:

Converte automaticamente o conteúdo das abas em HTML interativo.

Atualiza as páginas do WordPress todos os dias por meio de uma GitHub Action.

Identifica quais abas precisam de atualização através da coluna STATUS da planilha (ADICIONAR AO SITE, REMOVER, EDITAR).

Publica o HTML nas páginas corretas do WordPress via REST API.

---

## 🔧 Funcionalidades

✅ Atualização diária automática via GitHub Actions
✅ Geração de HTML com filtros por estado, cidade/país e categoria em JavaScript
✅ Integração direta com Google Sheets (via App Script + API)
✅ Mapeamento entre aba e URL configurado no arquivo main.py
✅ Substituição dinâmica do conteúdo no WordPress
✅ Processamento automatizado conforme o STATUS de cada aba

---

## 🛠️ Tecnologias utilizadas

- **Python 3**
- `pandas`
- `GitHub Actions`
- `GitHub Secrets`
- `requests`
- `openpyxl`
- **Google Sheets API + App Script**
- **WordPress REST API**

---

## 📂 Estrutura do Projeto

├── main.py # Script principal que orquestra a atualização

├── criarHTML.py # Geração do HTML com base nos dados da planilha

├── atualizador_WP.py # Conexão com a REST API do WordPress e substituição do conteúdo

├── conexao_api.py # Funções auxiliares para conexão e autenticação

├── requirements.txt # Dependências do projeto

---

## 🔐 Segurança

-As credenciais do WordPress e do Google estão armazenadas como GitHub Secrets, não sendo expostas no código.

-A autenticação com o WordPress é feita por senha de aplicativo, prática segura e recomendada para uso com a REST API.

---

## 👨‍💻 Autor

Desenvolvido por **Marcos Felipe Lopes Rodrigues**, aluno de Ciências Econômicas da UFPR e integrante do projeto de extensão *Inovação e Desenvolvimento Territorial*.

---

## 📃 Licença

Este projeto é de uso acadêmico e institucional, vinculado à Universidade Federal do Paraná.
