# 🔎 Sistema de Rastreamento de RMs (Material Requisitions Tracker)

![Python](https://img.shields.io/badge/Python-3.10%2B-blue)
![Streamlit](https://img.shields.io/badge/Streamlit-1.28-red)
![Pandas](https://img.shields.io/badge/Pandas-Data%20Analysis-green)
![Status](https://img.shields.io/badge/Status-Operational-brightgreen)

## 📋 Sobre o Projeto

Este projeto é uma ferramenta de automação desenvolvida em **Python** para otimizar o controle e rastreamento de Requisições de Materiais (RMs) em grandes projetos de engenharia.

O sistema resolve o problema da verificação manual em planilhas Excel massivas. Ele permite que o usuário carregue a base de dados oficial (cronograma/suprimentos), mapeie as colunas dinamicamente e consulte o status de materiais em tempo real, gerando relatórios formatados automaticamente para comunicação de restrições.

### 🚀 Funcionalidades Principais

* **Leitor Universal de Excel:** Aceita qualquer formato de planilha (`.xlsx`), permitindo que o usuário mapeie as colunas (Status, RMT, Previsão, Revisão) via interface gráfica.
* **Detector Automático:** Identifica e carrega automaticamente a planilha mais recente salva na pasta local.
* **Busca Inteligente:** Filtra RMs por código parcial ou total, ignorando formatações incorretas.
* **Gerador de Relatórios:** Formata automaticamente os dados encontrados (Status, Data, Revisão) em um texto padrão vertical pronto para copiar e colar (Ctrl+C) em planilhas de controle de restrições.
* **Interface Profissional:** UI moderna desenvolvida com **Streamlit**, com tema escuro e indicadores visuais de status.
* **Cache de Performance:** Utiliza o cache do Streamlit para garantir buscas instantâneas sem recarregar o Excel repetidamente.

---

## 🛠️ Tecnologias Utilizadas

* **Linguagem:** Python 3
* **Interface (Frontend/Backend):** Streamlit
* **Manipulação de Dados:** Pandas
* **Leitura de Arquivos:** OpenPyXL, OS

---

## 📦 Como Rodar o Projeto

Este projeto foi desenhado para ser portátil. Siga os passos abaixo para executar na sua máquina.

### Pré-requisitos

Você precisa ter o [Python](https://www.python.org/downloads/) instalado na sua máquina.

### Passo a Passo

1.  **Clone o repositório ou baixe a pasta:**
    ```bash
    git clone [https://github.com/SEU-USUARIO/sistema-busca-rms.git](https://github.com/SEU-USUARIO/sistema-busca-rms.git)
    ```

2.  **Instale as bibliotecas necessárias:**
    Abra o terminal na pasta do projeto e execute:
    ```bash
    pip install -r requirements.txt
    ```

3.  **Execute o Sistema:**
    * **Opção A (Windows):** Dê um duplo clique no arquivo `Iniciar_Sistema.bat`.
    * **Opção B (Terminal):** Digite o comando:
        ```bash
        python -m streamlit run app.py
        ```

---

## 📂 Estrutura de Arquivos

```text
/
├── app.py                     # Código principal da aplicação (Streamlit)
├── gerar_base_real.py         # Script auxiliar para gerar dados de teste
├── Iniciar_Sistema.bat        # Atalho para execução rápida no Windows
├── requirements.txt           # Lista de dependências do projeto
├── README.md                  # Documentação do projeto
└── RMs_Criticas_*.xlsx        # (Opcional) Planilhas de dados locais
