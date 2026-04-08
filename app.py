from __future__ import annotations

from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZIP_DEFLATED, ZipFile

import streamlit as st

from gerar_frequencias import processar_frequencias
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
col1, col2, col3, col4, col5 = st.columns(5)
col1.metric("Registros válidos", f"{meta['total_registros']:,}".replace(",", "."))
col2.metric("Registros inválidos", f"{int(meta.get('total_registros_invalidos', 0)):,}".replace(",", "."))
col3.metric("Frequência geral", f"{meta['frequencia_geral']:.1%}")
col4.metric("Alunos", f"{meta['total_alunos']:,}".replace(",", "."))
col5.metric("Disciplinas", f"{meta['total_disciplinas']:,}".replace(",", "."))

st.success("HTML gerado com sucesso.")
tab_dashboard, tab_invalidos, tab_resumo = st.tabs(["Dashboard", "Registros inválidos", "Resumo"])

with tab_dashboard:
    st.download_button(
        "Baixar dashboard HTML",
        data=html.encode("utf-8"),
        file_name=file_name,
        mime="text/html",
        use_container_width=True,
    )

    with st.expander("Prévia do HTML gerado"):
        st.components.v1.html(html, height=900, scrolling=True)

    st.divider()
    st.subheader("Frequências em PDF")
    st.caption("Gera um PDF por aluno (compactado em .zip) usando a planilha enviada.")

    if st.button("Gerar PDFs de frequência", use_container_width=True):
        try:
            with st.spinner("Gerando PDFs..."):
                with TemporaryDirectory(prefix="frequencias_") as temp_dir:
                    temp_path = Path(temp_dir)
                    input_path = temp_path / "FREQUENCIA.xlsx"
                    output_dir = temp_path / "frequencias"
                    input_path.write_bytes(file_bytes)
                    output_dir.mkdir(parents=True, exist_ok=True)

                    total = processar_frequencias(
                        str(input_path),
                        str(output_dir),
                        str(Path(__file__).resolve().parent),
                    )

                    zip_path = temp_path / "frequencias.zip"
                    with ZipFile(zip_path, "w", compression=ZIP_DEFLATED) as zip_file:
                        for pdf_path in sorted(output_dir.glob("*.pdf")):
                            zip_file.write(pdf_path, arcname=pdf_path.name)

                    st.success(f"PDFs gerados: {total}")
                    st.download_button(
                        "Baixar PDFs (ZIP)",
                        data=zip_path.read_bytes(),
                        file_name="frequencias.zip",
                        mime="application/zip",
                        use_container_width=True,
                    )
        except Exception as error:
            st.error(f"Falha ao gerar PDFs: {error}")

with tab_invalidos:
    invalid_records = payload.get("invalid_records") or []
    st.caption("Registros fora do padrão (não são Presente nem Ausente).")
    st.metric("Total de registros inválidos", f"{len(invalid_records):,}".replace(",", "."))

    if not invalid_records:
        st.info("Nenhum registro inválido encontrado.")
    else:
        st.dataframe(invalid_records, use_container_width=True, hide_index=True)
        invalid_csv = "\n".join(
            [
                "DATA,DATA_EXIBICAO,CURSO,TURNO,DISCIPLINA,ALUNO,STATUS,MOTIVO",
                *[
                    ",".join(
                        str(row.get(key, "")).replace("\n", " ").replace("\r", " ").replace(",", ";")
                        for key in [
                            "DATA",
                            "DATA_EXIBICAO",
                            "CURSO",
                            "TURNO",
                            "DISCIPLINA",
                            "ALUNO",
                            "STATUS",
                            "MOTIVO",
                        ]
                    )
                    for row in invalid_records
                ],
            ]
        )
        st.download_button(
            "Baixar inválidos (CSV)",
            data=invalid_csv.encode("utf-8"),
            file_name="registros_invalidos.csv",
            mime="text/csv",
            use_container_width=True,
        )

with tab_resumo:
    st.write(
        {
            "periodo_inicio": meta["periodo_inicio"],
            "periodo_fim": meta["periodo_fim"],
            "gerado_em": meta["gerado_em"],
            "registros_brutos": meta.get("total_registros_brutos", ""),
            "registros_validos": meta.get("total_registros", ""),
            "registros_invalidos": meta.get("total_registros_invalidos", ""),
        }
    )
