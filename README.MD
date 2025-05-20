# 🛰️ Atualizador de Páginas do Projeto INOVA – UFPR

Este projeto automatiza a atualização de páginas do site [inova.ufpr.br](https://inova.ufpr.br), que divulga organizações de inovação mapeadas pelo projeto de extensão **Inovação e Desenvolvimento Territorial** da UFPR.

## ✨ Objetivo

A equipe do projeto mantém uma planilha no Google Sheets com diversas abas (startups, aceleradoras, universidades, etc.). Este sistema:

- Converte automaticamente o conteúdo das abas em HTML interativo
- Permite selecionar as abas a atualizar por meio de uma interface gráfica
- Publica o HTML nas páginas corretas do WordPress via REST API

---

## 🔧 Funcionalidades

✅ Interface gráfica (Tkinter) para selecionar as abas a atualizar  
✅ Geração automática de HTML com filtros por estado, cidade/país e categoria  
✅ Conexão direta com planilha online (Google Sheets)  
✅ Mapeamento entre aba e URL via Excel local  
✅ Substituição dinâmica do conteúdo no WordPress com autenticação segura  
✅ Feedback visual de sucesso ou erro na atualização

---

## 🛠️ Tecnologias utilizadas

- **Python 3**
- `pandas`
- `tkinter`
- `requests`
- `openpyxl`
- `python-dotenv`
- **Google Sheets** (via exportação pública em XLSX)
- **WordPress REST API**

---

## 📂 Estrutura do Projeto

├── criarHTML.py # Geração do HTML com base nos dados da planilha
├── atualizador_WP.py # Conexão com a REST API do WordPress e substituição do conteúdo
├── interface.py # Interface gráfica com Tkinter para escolher abas e executar atualização
├── links_startups.xlsx # Planilha local que relaciona nome da aba ao link da página
├── .env # Credenciais de acesso ao WordPress (não versionado)

---

## 🔐 Segurança

- As credenciais do WordPress são armazenadas localmente em um arquivo `.env`, que **não deve ser versionado**.
- A autenticação com o WordPress é feita por meio de senha de aplicativo, uma prática segura e recomendada para uso com a REST API.

---

## 👨‍💻 Autor

Desenvolvido por **[Marcos Felipe Lopes Rodrigues]**, aluno de Ciências Econômicas da UFPR e integrante do projeto de extensão *Inovação e Desenvolvimento Territorial*.

---

## 📃 Licença

Este projeto é de uso acadêmico e institucional, vinculado à Universidade Federal do Paraná.