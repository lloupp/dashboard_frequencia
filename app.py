from __future__ import annotations

from pathlib import Path

import streamlit as st

from generate_online_bundle import generate_dashboard_artifacts


PAGE_TITLE = "Gerador de Dashboard HTML"
DEFAULT_FILE = Path.home() / "Downloads" / "FREQUÊNCIA.xlsx"


st.set_page_config(page_title=PAGE_TITLE, layout="wide")

st.title("Gerador de Dashboard HTML")
st.caption("Envie a planilha e baixe um dashboard HTML pronto, sem depender de instalação extra para quem for visualizar.")

uploaded_file = st.file_uploader(
    "Envie a planilha Excel",
    type=["xlsx"],
    help="O app usa as abas principais da planilha e gera um HTML standalone para download.",
)

use_default = st.toggle(
    "Usar automaticamente a planilha do Downloads",
    value=uploaded_file is None and DEFAULT_FILE.exists(),
)

file_bytes: bytes | None = None
file_name = "dashboard_online.html"

if uploaded_file is not None:
    file_bytes = uploaded_file.getvalue()
    file_name = f"{Path(uploaded_file.name).stem}_dashboard.html"
elif use_default and DEFAULT_FILE.exists():
    file_bytes = DEFAULT_FILE.read_bytes()

if not file_bytes:
    st.info("Envie a planilha `.xlsx` para gerar o dashboard HTML.")
    st.stop()

with st.spinner("Gerando dashboard HTML..."):
    payload, html = generate_dashboard_artifacts(file_bytes)

meta = payload["meta"]
col1, col2, col3, col4 = st.columns(4)
col1.metric("Registros válidos", f"{meta['total_registros']:,}".replace(",", "."))
col2.metric("Frequência geral", f"{meta['frequencia_geral']:.1%}")
col3.metric("Alunos", f"{meta['total_alunos']:,}".replace(",", "."))
col4.metric("Disciplinas", f"{meta['total_disciplinas']:,}".replace(",", "."))

st.success("HTML gerado com sucesso.")
st.download_button(
    "Baixar dashboard HTML",
    data=html.encode("utf-8"),
    file_name=file_name,
    mime="text/html",
    use_container_width=True,
)

with st.expander("Prévia do HTML gerado"):
    st.components.v1.html(html, height=900, scrolling=True)

with st.expander("Resumo da geração"):
    st.write(
        {
            "periodo_inicio": meta["periodo_inicio"],
            "periodo_fim": meta["periodo_fim"],
            "gerado_em": meta["gerado_em"],
        }
    )
