# leitordedisco_ocr
API e scripts em Python para consulta ao DataJud (CNJ) e classificação automática de processos em Cível ou Criminal. Utiliza mapeamento por código numérico da Tabela Processual Unificada (TPU/SGT), eliminando erros de texto, web scraping e CAPTCHAs. Integrável via REST e Webhooks.
# ⚖️ Classificador de Processos Jurídicos (DataJud + TPU/CNJ)

Microsserviço em Python para consulta e classificação de processos judiciais brasileiros em **Cível** ou **Criminal** com 100% de precisão.

## 🚀 Sobre o Projeto

O número único do CNJ não indica se um processo é cível ou criminal. Este projeto resolve o problema consultando a **API Pública do DataJud** e classificando a matéria com base nos códigos numéricos da **Tabela Processual Unificada (TPU)** do CNJ, evitando instabilidades com *web scraping* ou erros de regex/texto.

## ✨ Funcionalidades

- 🔍 **Consulta ao DataJud:** Consumo direto da API oficial do CNJ via REST.
- 🎯 **Classificação Precisa:** Mapeamento O(1) por código de classe da TPU (sem regex ou texto).
- 🔄 **Automação TPU:** Script embutido para sincronização automática da árvore de classes do SGT/CNJ.
- ⚡ **Sem Scraping:** Livre de CAPTCHAs e bloqueios de navegação.

## 🛠️ Tecnologias Utilizadas

- **Linguagem:** Python 3.10+
- **Bibliotecas:** `requests`
- **APIs Integradas:** DataJud (CNJ) e SGT (CNJ)

- ## 📌 Dependencias

  python main.py

  pip install requests
