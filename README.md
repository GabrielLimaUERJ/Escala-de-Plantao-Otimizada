# 📅 Sistema de Escala Automatizada

Aplicação em Python para geração automática de escalas de plantão, com foco em distribuição equilibrada, respeito a restrições operacionais e redução de conflitos.

Interface construída com Streamlit para facilitar simulações e ajustes em tempo real.

---

## 💡 Contexto Real

Este projeto foi inspirado em um fluxo real de organização de escalas, onde a alocação manual consome tempo e está sujeita a erros, especialmente em cenários com múltiplas restrições (férias, intervalos mínimos, tipos de plantão, etc.).

---

## 🎯 Objetivo

Automatizar a geração de escalas, garantindo:

* Distribuição equilibrada entre colaboradores
* Respeito a intervalos mínimos entre plantões
* Consideração de férias e indisponibilidades
* Separação por regiões operacionais
* Redução de conflitos e retrabalho

---

## ⚙️ Lógica do Sistema

O algoritmo considera múltiplas regras, como:

* Intervalo mínimo entre plantões
* Limite mensal por colaborador
* Bloqueio de plantões próximos
* Controle de carga nas últimas semanas
* Priorização por histórico (última execução)

A geração ocorre em duas etapas:

1. Alocação de plantões especiais (noturno, recesso, etc.)
2. Distribuição de plantões diurnos por região

---

## 🛠️ Tecnologias

* Python
* Pandas
* Streamlit

---

## 📚 Funcionalidades

* Geração automática de escala mensal
* Configuração dinâmica via interface (Streamlit)
* Controle por região (limites e quantidade diária)
* Visualização da escala e distribuição de carga
* Contador de plantões por colaborador

---

## 📂 Estrutura do Projeto

```bash
/projeto-escala
│
├── app.py
├── requirements.txt
├── dados.csv
├── ferias.csv
├── regioes.csv
└── calendario_raw.csv
```

---

## 📄 Estrutura dos Dados

Os dados são organizados em arquivos CSV:

* `dados.csv` → histórico de plantões
* `ferias.csv` → controle de férias
* `regioes.csv` → mapeamento por região
* `calendario_raw.csv` → feriados e recessos

⚠️ Os dados foram estruturados para simular um ambiente real e podem ser adaptados conforme necessidade.

---

## 🔗 Fonte alternativa (Google Sheets)

Os dados também podem ser obtidos via Google Sheets nas últimas abas:

https://docs.google.com/spreadsheets/d/1cvGJ2WLMaJNIXdmPbHkHfRIzE3Ph2DbBJs4otGtqn5E/edit

---

## ▶️ Como executar

Clone o repositório:

```bash
git clone https://github.com/seu-usuario/seu-repo.git
cd seu-repo
```

Crie um ambiente virtual (opcional, recomendado):

```bash
python -m venv venv
source venv/bin/activate  # Linux/macOS
venv\Scripts\activate     # Windows
```

Instale as dependências:

```bash
pip install -r requirements.txt
```

Execute a aplicação:

```bash
streamlit run app.py
```

---

## 🧠 Problema Resolvido

A organização manual de escalas pode gerar:

* conflitos de agenda
* distribuição desigual de carga
* inclusão de pessoas indisponíveis
* retrabalho constante

Este sistema automatiza esse processo, tornando a geração de escalas mais rápida, consistente e confiável.

---

## 🚀 Possíveis Melhorias

* Redução de conflitos por questões subjetivas ou multiobjetivas
* A quantidade de nomes disponíveis pode não ser suficiente
* Nomes ausentes por férias ou licenças podem precisar de ajustes manuais para serem escalados
* Interface de edição manual pós-geração
* Deploy em ambiente web
