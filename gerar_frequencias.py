
import pandas as pd
import os
import sys
import unicodedata
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.enums import TA_CENTER, TA_LEFT

# Cores institucionais (nova paleta)
AZUL_ESCURO = colors.HexColor('#0F2854')
AZUL_MEDIO = colors.HexColor('#1C4D8D')
AZUL_CLARO = colors.HexColor('#4988C4')
AZUL_BEM_CLARO = colors.HexColor('#BDE8F5')
CINZA = colors.HexColor('#F5F5F5')
VERDE = colors.HexColor('#1E7F4E')
VERMELHO = colors.HexColor('#C00000')

def _normalize_column_name(value: object) -> str:
    text = "" if value is None else str(value)
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    text = text.strip().upper()
    text = text.replace("\n", " ").replace("\r", " ")
    text = "_".join(part for part in text.replace("/", " ").replace("-", " ").split() if part)
    return text


def _canonicalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.rename(columns={column: _normalize_column_name(column) for column in df.columns})

    synonyms = {
        "ALUNO": {"ALUNO", "ALUNOA", "NOME", "NOME_ALUNO", "ESTUDANTE"},
        "DISCIPLINA": {"DISCIPLINA", "DISCIPLINAS", "MATERIA", "MODULO"},
        "PRESENTE": {"PRESENTE", "PRESENCA", "PRESENCAS", "PRESENTES", "PRES"},
        "AUSENTE": {"AUSENTE", "AUSENCIA", "AUSENCIAS", "AUSENTES", "FALTA", "FALTAS"},
        "STATUS": {"STATUS", "SITUACAO", "SITUACAO_DO_ALUNO", "SITUACAO_DE_FREQUENCIA"},
    }

    for canonical, candidates in synonyms.items():
        if canonical in df.columns:
            continue
        for candidate in candidates:
            if candidate in df.columns:
                df = df.rename(columns={candidate: canonical})
                break

    return df


def _as_bool(series: pd.Series) -> pd.Series:
    if series.dtype == bool:
        return series.fillna(False)
    if pd.api.types.is_numeric_dtype(series):
        return series.fillna(0).astype(float) != 0

    text = series.fillna("").astype(str).str.strip().str.lower()
    truthy = {"1", "true", "t", "sim", "s", "x", "ok", "presente", "p"}
    falsy = {"0", "false", "f", "nao", "não", "n", ""}
    return text.apply(lambda value: True if value in truthy else False if value in falsy else False)


def _get_presence_absence(df: pd.DataFrame) -> tuple[pd.Series, pd.Series]:
    if "PRESENTE" in df.columns and "AUSENTE" in df.columns:
        return _as_bool(df["PRESENTE"]), _as_bool(df["AUSENTE"])

    if "STATUS" in df.columns:
        status = df["STATUS"].fillna("").astype(str).str.strip().str.lower()
        present = status.isin({"presente", "p"})
        absent = status.isin({"ausente", "a"})
        return present, absent

    raise KeyError("Colunas de frequência não encontradas (PRESENTE/AUSENTE ou STATUS).")

def gerar_pdf_aluno(nome_aluno, turma, disciplinas_data, output_path, base_path):
    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        leftMargin=2*cm, rightMargin=2*cm,
        topMargin=2*cm, bottomMargin=2*cm
    )

    styles = getSampleStyleSheet()
    story = []

    imagem_central_path = os.path.join(base_path, 'santa_ufcspa.png')
    if os.path.exists(imagem_central_path):
        imagem_central = Image(imagem_central_path, width=10*cm, height=5*cm)
        imagem_table = Table([[imagem_central]], colWidths=[17*cm])
        imagem_table.setStyle(TableStyle([
            ('ALIGN', (0,0), (0,0), 'CENTER'),
            ('VALIGN', (0,0), (0,0), 'MIDDLE'),
            ('BOTTOMPADDING', (0,0), (0,0), 6),
        ]))
        story.append(imagem_table)
        story.append(Spacer(1, 0.2*cm))

    # Cabeçalho
    inst_style = ParagraphStyle('inst', fontSize=11, textColor=colors.white, alignment=TA_CENTER, fontName='Helvetica-Bold')
    prog_style = ParagraphStyle('prog', fontSize=10, textColor=colors.white, alignment=TA_CENTER, fontName='Helvetica')
    header_style = ParagraphStyle('header', fontSize=9, textColor=colors.white, alignment=TA_CENTER, fontName='Helvetica')

    header_data = [
        [Paragraph('UFCSPA / SANTA CASA DE PORTO ALEGRE', inst_style)],
        [Paragraph('Pós-Médica — Programa de Especialização Médica', prog_style)],
        [Paragraph(f'Especialização em {turma}', header_style)],
    ]
    header_table = Table(header_data, colWidths=[17*cm])
    header_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), AZUL_ESCURO),
        ('PADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(header_table)
    story.append(Spacer(1, 0.5*cm))

    # Título e dados do aluno
    titulo_style = ParagraphStyle('titulo', fontSize=13, textColor=AZUL_ESCURO, alignment=TA_CENTER, fontName='Helvetica-Bold', spaceAfter=6)
    story.append(Paragraph('COMPROVANTE DE FREQUÊNCIA', titulo_style))
    
    aluno_style = ParagraphStyle('aluno', fontSize=11, textColor=AZUL_ESCURO, fontName='Helvetica-Bold')
    label_style = ParagraphStyle('label', fontSize=9, textColor=colors.HexColor('#555555'), fontName='Helvetica')
    aluno_data = [[Paragraph('ALUNO(A)', label_style), Paragraph(nome_aluno, aluno_style)]]
    aluno_table = Table(aluno_data, colWidths=[3*cm, 14*cm])
    aluno_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), AZUL_BEM_CLARO),
        ('PADDING', (0,0), (-1,-1), 8),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ]))
    story.append(aluno_table)
    story.append(Spacer(1, 0.6*cm))

    # Tabela de frequências
    col_header_style = ParagraphStyle('colh', fontSize=8, textColor=colors.white, fontName='Helvetica-Bold', alignment=TA_CENTER)
    cell_style = ParagraphStyle('cell', fontSize=9, fontName='Helvetica', alignment=TA_CENTER)
    disc_style = ParagraphStyle('disc', fontSize=9, fontName='Helvetica', alignment=TA_LEFT)

    table_data = [[
        Paragraph('DISCIPLINA', col_header_style),
        Paragraph('PRESENÇAS', col_header_style),
        Paragraph('AUSÊNCIAS', col_header_style),
        Paragraph('TOTAL AULAS', col_header_style),
        Paragraph('% FREQUÊNCIA', col_header_style),
    ]]

    row_styles = []
    total_presencas, total_ausencias = 0, 0
    for i, (disciplina, presentes, ausentes) in enumerate(disciplinas_data):
        total = presentes + ausentes
        pct = (presentes / total * 100) if total > 0 else 0
        total_presencas += presentes
        total_ausencias += ausentes
        
        pct_text = f'<font color="{VERDE if pct >= 75 else VERMELHO}"><b>{pct:.0f}%</b></font>'
        
        table_data.append([
            Paragraph(disciplina, disc_style),
            Paragraph(str(presentes), cell_style),
            Paragraph(str(ausentes), cell_style),
            Paragraph(str(total), cell_style),
            Paragraph(pct_text, cell_style),
        ])
        
        # Define a cor de fundo para a linha
        bg_color = colors.white if i % 2 == 0 else CINZA
        row_styles.append(('BACKGROUND', (0, i + 1), (-1, i + 1), bg_color))


    # Linha de total
    total_geral = total_presencas + total_ausencias
    pct_geral = (total_presencas / total_geral * 100) if total_geral > 0 else 0
    pct_geral_text = f'<font color="{VERDE if pct_geral >= 75 else VERMELHO}"><b>{pct_geral:.0f}%</b></font>'
    total_style = ParagraphStyle('tot', fontSize=9, fontName='Helvetica-Bold', alignment=TA_CENTER)
    
    table_data.append([
        Paragraph('<b>TOTAL GERAL</b>', ParagraphStyle('tot_label', fontSize=9, fontName='Helvetica-Bold', alignment=TA_LEFT)),
        Paragraph(f'<b>{total_presencas}</b>', total_style),
        Paragraph(f'<b>{total_ausencias}</b>', total_style),
        Paragraph(f'<b>{total_geral}</b>', total_style),
        Paragraph(pct_geral_text, total_style),
    ])

    freq_table = Table(table_data, colWidths=[7.9*cm, 2.5*cm, 2.5*cm, 2.5*cm, 2.6*cm], repeatRows=1)
    
    # Monta o estilo da tabela
    style = [
        ('BACKGROUND', (0,0), (-1,0), AZUL_MEDIO),
        ('BACKGROUND', (0,-1), (-1,-1), AZUL_BEM_CLARO),
        ('GRID', (0,0), (-1,-1), 0.5, colors.lightgrey),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING', (0,0), (-1,-1), 6),
    ]
    style.extend(row_styles) # Adiciona os estilos das linhas
    freq_table.setStyle(TableStyle(style))

    story.append(freq_table)
    story.append(Spacer(1, 0.8*cm))

    # Rodapé
    nota_style = ParagraphStyle('nota', fontSize=8, textColor=colors.darkgrey, fontName='Helvetica-Oblique', alignment=TA_CENTER)
    story.append(Paragraph('Frequência mínima: 75%. Em caso de dúvidas, contate a coordenação.', nota_style))
    story.append(Spacer(1, 0.2*cm))
    
    sep_table = Table([['']], colWidths=[17*cm], rowHeights=[1])
    sep_table.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,-1), AZUL_ESCURO)]))
    story.append(sep_table)
    story.append(Spacer(1, 0.2*cm))
    
    rodape_style = ParagraphStyle('rodape', fontSize=7, textColor=colors.grey, fontName='Helvetica', alignment=TA_CENTER)
    story.append(Paragraph('Pós-Médica UFCSPA e Santa Casa · Porto Alegre, RS', rodape_style))

    doc.build(story)

def processar_frequencias(input_file, output_dir, base_path):
    try:
        dfs = pd.read_excel(input_file, sheet_name=None, header=1)
    except FileNotFoundError:
        print(f"Erro: Arquivo de entrada não encontrado em '{input_file}'")
        return 0

    total_gerados = 0
    for turma, df in dfs.items():
        df = _canonicalize_columns(df)
        if "ALUNO" not in df.columns or "DISCIPLINA" not in df.columns:
            continue

        presentes, ausentes = _get_presence_absence(df)
        df = df.copy()
        df["_PRESENTES"] = presentes.astype(int)
        df["_AUSENTES"] = ausentes.astype(int)

        summary = (
            df.groupby(["ALUNO", "DISCIPLINA"], dropna=False, as_index=False)
            .agg(
                PRESENTES=("_PRESENTES", "sum"),
                AUSENTES=("_AUSENTES", "sum"),
            )
            .reset_index(drop=True)
        )

        for aluno in summary['ALUNO'].unique():
            dados_aluno = summary[summary['ALUNO'] == aluno]
            disciplinas_data = [
                (row['DISCIPLINA'], int(row['PRESENTES']), int(row['AUSENTES']))
                for _, row in dados_aluno.iterrows()
            ]
            
            nome_arquivo = "".join(c for c in aluno if c.isalnum() or c in " .-_").rstrip().replace(' ', '_')
            turma_safe = "".join(c for c in turma if c.isalnum() or c in " .-_").rstrip().replace(' ', '_')
            output_path = os.path.join(output_dir, f'{turma_safe}__{nome_arquivo}.pdf')
            
            gerar_pdf_aluno(aluno, turma, disciplinas_data, output_path, base_path)
            total_gerados += 1
            
    return total_gerados

def main():
    base_path = os.path.dirname(__file__)
    
    if len(sys.argv) > 2:
        input_file = sys.argv[1]
        output_dir = sys.argv[2]
    else:
        input_file = os.path.join(base_path, 'FREQUÊNCIA.xlsx')
        output_dir = os.path.join(base_path, 'frequencias')
        
    os.makedirs(output_dir, exist_ok=True)
    
    num_gerados = processar_frequencias(input_file, output_dir, base_path)
    print(f'PDFs gerados: {num_gerados}')

if __name__ == "__main__":
    main()
