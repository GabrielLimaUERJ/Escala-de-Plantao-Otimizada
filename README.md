# 📅 Sistema de Geração de Escalas

Sistema desenvolvido em Python + Streamlit para geração automática de escalas de plantão, considerando regras como:

* Intervalo mínimo entre plantões
* Distribuição equilibrada
* Férias
* Regiões
* Tipos de plantão (Noturno, Recesso, etc.)

## 🚀 Como rodar

1. Clone o repositório:

```
git clone https://github.com/seu-usuario/seu-repo.git
```

2. Instale as dependências:

```
pip install -r requirements.txt
```

3. Execute:

```
streamlit run app.py
```

## 📂 Estrutura dos dados

Os dados estão na pasta `/data`:

* `dados.csv` → histórico de plantões
* `ferias.csv` → controle de férias
* `regioes.csv` → mapeamento por região
* `calendario_raw.csv` → feriados e recessos

## 🔗 Fonte alternativa

Os dados também podem ser obtidos via Google Sheets:
https://docs.google.com/spreadsheets/d/1cvGJ2WLMaJNIXdmPbHkHfRIzE3Ph2DbBJs4otGtqn5E/edit

## 🧠 Tecnologias

* Python
* Pandas
* Streamlit
