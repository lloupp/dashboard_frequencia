# Dashboard de Frequência (Streamlit)

App em Streamlit para:

- Enviar uma planilha `.xlsx` de frequência
- Gerar e baixar um **dashboard HTML** pronto (standalone)
- Gerar e baixar **PDFs de frequência** (1 por aluno, compactado em `.zip`)
- Inspecionar **registros inválidos**

## Rodar localmente

```bash
python -m venv .venv
# Windows
.\.venv\Scripts\python -m pip install -r requirements.txt
.\.venv\Scripts\streamlit run app.py
```

## Hospedar no Streamlit Community Cloud

Pré-requisitos:

- Repositório no GitHub
- Arquivos no repo: `app.py` e `requirements.txt`

Passo a passo:

1. Acesse o Streamlit Community Cloud e clique em **New app**.
2. Selecione o repositório (GitHub) e a branch que você quer publicar.
3. Defina:
   - **Main file path**: `app.py`
4. Clique em **Deploy**.

Observação: no Streamlit Cloud não existe a planilha no `Downloads`, então use o upload dentro do app.

## Logos / Cabeçalho no HTML gerado

O HTML gerado embute as imagens em base64 para ficar “standalone”.

Use uma das opções abaixo:

- Cabeçalho inteiro: `santa_ufcspa.png` (na raiz do projeto **ou** em `assets/`)
- Ou logos separados: `assets/santa_casa.png` e `assets/ufcspa.png`

## Saídas geradas

- **Dashboard HTML**: baixado pelo botão do app
- **PDFs**: baixados como `frequencias.zip`

