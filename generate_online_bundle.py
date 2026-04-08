from __future__ import annotations

import json
from io import BytesIO
from pathlib import Path

import pandas as pd

from generate_dashboards import PRIMARY_SHEETS, find_source_file, load_data, summarize_data


ROOT = Path(__file__).resolve().parent
OUTPUT_HTML = ROOT / "dashboard_online.html"
OUTPUT_JSON = ROOT / "dashboard_data.json"


def build_payload(data: pd.DataFrame, summary: dict[str, pd.DataFrame | dict[str, object]]) -> dict[str, object]:
    records = data[["DATA", "DATA_EXIBICAO", "CURSO", "TURNO", "DISCIPLINA", "ALUNO", "STATUS"]].copy()
    records["DATA"] = records["DATA"].dt.strftime("%Y-%m-%d")
    records["DIA_SEMANA"] = pd.to_datetime(records["DATA"]).dt.day_name().map(
        {
            "Monday": "Segunda",
            "Tuesday": "Terça",
            "Wednesday": "Quarta",
            "Thursday": "Quinta",
            "Friday": "Sexta",
            "Saturday": "Sábado",
            "Sunday": "Domingo",
        }
    )
    records = records[records["STATUS"].isin(["Presente", "Ausente"])].reset_index(drop=True)

    return {
        "meta": summary["meta"],
        "records": records.to_dict(orient="records"),
    }


def load_uploaded_excel(file_bytes: bytes) -> pd.DataFrame:
    excel_buffer = BytesIO(file_bytes)
    excel = pd.ExcelFile(excel_buffer)
    frames: list[pd.DataFrame] = []

    for sheet_name in excel.sheet_names:
        if sheet_name not in PRIMARY_SHEETS:
            continue

        frame = pd.read_excel(BytesIO(file_bytes), sheet_name=sheet_name, header=1)
        frame.columns = [str(column).strip() for column in frame.columns]
        frame = frame.dropna(how="all")
        required = ["DATA", "TURNO", "DISCIPLINA", "ALUNO", "PRESENTE", "AUSENTE"]
        if not set(required).issubset(frame.columns):
            continue

        frame = frame[required].copy()
        frame["CURSO"] = sheet_name
        frames.append(frame)

    if not frames:
        raise ValueError("Nenhuma aba válida foi encontrada para consolidar.")

    data = pd.concat(frames, ignore_index=True)
    data["DATA"] = pd.to_datetime(data["DATA"], errors="coerce")
    data["DATA_EXIBICAO"] = data["DATA"].dt.strftime("%d/%m/%Y")
    data["TURNO"] = data["TURNO"].fillna("Não informado").astype(str).str.strip()
    data["DISCIPLINA"] = data["DISCIPLINA"].fillna("Não informada").astype(str).str.strip()
    data["ALUNO"] = data["ALUNO"].fillna("Não informado").astype(str).str.strip()
    data["PRESENTE"] = data["PRESENTE"].fillna(False).astype(bool)
    data["AUSENTE"] = data["AUSENTE"].fillna(False).astype(bool)
    data["STATUS"] = data.apply(
        lambda row: "Presente"
        if row["PRESENTE"]
        else "Ausente"
        if row["AUSENTE"]
        else "Sem marcação",
        axis=1,
    )
    data["MES"] = data["DATA"].dt.strftime("%Y-%m")
    data["AULAS"] = 1
    return data.sort_values(["DATA", "CURSO", "DISCIPLINA", "ALUNO"]).reset_index(drop=True)


def generate_dashboard_artifacts(file_bytes: bytes) -> tuple[dict[str, object], str]:
    data = load_uploaded_excel(file_bytes)
    summary = summarize_data(data)
    payload = build_payload(data, summary)
    html = render_html(payload)
    return payload, html


def render_html(payload: dict[str, object]) -> str:
    payload_json = json.dumps(payload, ensure_ascii=False)
    return f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Dashboard de Frequência</title>
  <style>
    :root {{
      --bg: #f4f7fb;
      --card: rgba(255,255,255,0.92);
      --text: #102033;
      --muted: #5f6b7a;
      --stroke: rgba(15, 23, 42, 0.08);
      --shadow: 0 18px 40px rgba(15, 23, 42, 0.08);
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      font-family: "Segoe UI", "Helvetica Neue", sans-serif;
      color: var(--text);
      background:
        radial-gradient(circle at top left, rgba(59,130,246,0.18), transparent 32%),
        radial-gradient(circle at right center, rgba(245,158,11,0.14), transparent 25%),
        linear-gradient(180deg, #f8fbff 0%, var(--bg) 100%);
    }}
    .wrap {{ max-width: 1280px; margin: 0 auto; padding: 32px 20px 48px; }}
    .hero {{ display: flex; justify-content: space-between; gap: 20px; align-items: end; margin-bottom: 24px; }}
    .hero h1 {{ margin: 0; font-size: clamp(2rem, 4vw, 3.2rem); line-height: 1; letter-spacing: -0.03em; }}
    .hero p {{ margin: 10px 0 0; color: var(--muted); max-width: 760px; font-size: 1rem; }}
    .stamp {{ background: rgba(255,255,255,0.85); border: 1px solid var(--stroke); box-shadow: var(--shadow); border-radius: 18px; padding: 16px 18px; min-width: 220px; }}
    .grid {{ display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 16px; margin-bottom: 18px; }}
    .card, .panel {{ background: var(--card); border: 1px solid var(--stroke); box-shadow: var(--shadow); border-radius: 22px; }}
    .card {{ padding: 18px; }}
    .label {{ color: var(--muted); font-size: 0.9rem; margin-bottom: 12px; }}
    .value {{ font-size: 2rem; font-weight: 700; letter-spacing: -0.03em; }}
    .panel {{ padding: 20px; }}
    .filters {{ display: grid; grid-template-columns: repeat(6, minmax(0, 1fr)); gap: 14px; margin-bottom: 16px; }}
    .filters label {{ display: block; font-size: 0.82rem; color: var(--muted); margin-bottom: 8px; text-transform: uppercase; letter-spacing: 0.05em; }}
    .filters input, .filters select {{ width: 100%; border: 1px solid rgba(15, 23, 42, 0.1); background: rgba(255,255,255,0.9); border-radius: 14px; padding: 12px 14px; color: var(--text); font-size: 0.96rem; outline: none; }}
    .panels {{ display: grid; grid-template-columns: 1.4fr 1fr; gap: 16px; margin-bottom: 16px; }}
    .panel h2 {{ margin: 0 0 18px; font-size: 1.1rem; }}
    .bars {{ display: grid; gap: 12px; }}
    .bar-row {{ display: grid; grid-template-columns: 220px 1fr 68px; gap: 12px; align-items: center; }}
    .bar-track {{ height: 14px; background: #e7edf6; border-radius: 999px; overflow: hidden; }}
    .bar-fill {{ height: 100%; border-radius: 999px; background: linear-gradient(90deg, #2563eb 0%, #38bdf8 100%); }}
    .mini-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 16px; }}
    table {{ width: 100%; border-collapse: collapse; font-size: 0.95rem; }}
    th, td {{ padding: 10px 0; border-bottom: 1px solid rgba(15, 23, 42, 0.08); text-align: left; }}
    th {{ color: var(--muted); font-weight: 600; font-size: 0.82rem; text-transform: uppercase; letter-spacing: 0.04em; }}
    .timeline {{ height: 290px; display: flex; align-items: end; gap: 8px; padding-top: 18px; overflow-x: auto; }}
    .timeline-col {{ flex: 1; display: flex; flex-direction: column; justify-content: end; align-items: center; gap: 8px; min-width: 42px; }}
    .stack {{ width: 100%; max-width: 42px; display: flex; flex-direction: column; justify-content: end; border-radius: 14px 14px 8px 8px; overflow: hidden; background: #edf2f8; min-height: 14px; }}
    .present {{ background: linear-gradient(180deg, #22c55e 0%, #15803d 100%); }}
    .absent {{ background: linear-gradient(180deg, #fb7185 0%, #b91c1c 100%); }}
    .date-label {{ writing-mode: vertical-rl; transform: rotate(180deg); font-size: 0.72rem; color: var(--muted); letter-spacing: 0.04em; }}
    .foot {{ margin-top: 18px; color: var(--muted); font-size: 0.86rem; text-align: center; }}
    @media (max-width: 980px) {{
      .grid, .panels, .mini-grid, .filters {{ grid-template-columns: 1fr; }}
      .hero {{ flex-direction: column; align-items: start; }}
      .bar-row {{ grid-template-columns: 1fr; }}
    }}
  </style>
</head>
<body>
  <div class="wrap">
    <section class="hero">
      <div>
        <h1>Dashboard de Frequência</h1>
        <p>Versão pronta para publicação com atualização por arquivo JSON. Ideal para OneDrive, SharePoint ou hospedagem estática interna.</p>
      </div>
      <div class="stamp">
        <div class="label">Arquivo gerado</div>
        <div class="value" style="font-size:1.2rem;" id="gerado-em"></div>
      </div>
    </section>

    <section class="grid">
      <article class="card"><div class="label">Registros válidos</div><div class="value" id="registros"></div></article>
      <article class="card"><div class="label">Frequência geral</div><div class="value" id="frequencia"></div></article>
      <article class="card"><div class="label">Alunos</div><div class="value" id="alunos"></div></article>
      <article class="card"><div class="label">Disciplinas</div><div class="value" id="disciplinas"></div></article>
    </section>

    <section class="panel" style="margin-bottom:16px;">
      <h2>Filtros rápidos</h2>
      <div class="filters">
        <div><label for="course-filter">Curso</label><select id="course-filter"></select></div>
        <div><label for="shift-filter">Turno</label><select id="shift-filter"></select></div>
        <div><label for="discipline-filter">Disciplina</label><input id="discipline-filter" type="text" placeholder="Digite parte do nome" /></div>
        <div><label for="student-filter">Aluno</label><input id="student-filter" type="text" placeholder="Buscar aluno" /></div>
        <div><label for="start-date">Data inicial</label><input id="start-date" type="date" /></div>
        <div><label for="end-date">Data final</label><input id="end-date" type="date" /></div>
        <div><label for="frequency-filter">Faixa de frequência</label><select id="frequency-filter"><option value="0">Sem corte</option><option value="below_75">Abaixo de 75%</option><option value="0.75">75%+</option><option value="0.8">80%+</option><option value="0.9">90%+</option></select></div>
        <div><label for="weekday-filter">Dia da semana</label><select id="weekday-filter"><option value="Todos">Todos</option><option value="Segunda">Segunda</option><option value="Terça">Terça</option><option value="Quarta">Quarta</option><option value="Quinta">Quinta</option><option value="Sexta">Sexta</option><option value="Sábado">Sábado</option><option value="Domingo">Domingo</option></select></div>
      </div>
    </section>

    <section class="panels">
      <article class="panel">
        <h2>Frequência por curso</h2>
        <div class="bars" id="course-bars"></div>
      </article>
      <article class="panel">
        <h2>Resumo executivo</h2>
        <table>
          <tbody>
            <tr><th>Presenças</th><td id="presencas"></td></tr>
            <tr><th>Ausências</th><td id="ausencias"></td></tr>
            <tr><th>Período</th><td id="periodo"></td></tr>
            <tr><th>Cursos</th><td id="cursos-total"></td></tr>
          </tbody>
        </table>
      </article>
    </section>

    <section class="mini-grid">
      <article class="panel">
        <h2>Top alunos com mais ausências</h2>
        <table><thead><tr><th>Aluno</th><th>Curso</th><th>Ausências</th><th>Freq.</th></tr></thead><tbody id="students-table"></tbody></table>
      </article>
      <article class="panel">
        <h2>Disciplinas com mais ausências</h2>
        <table><thead><tr><th>Disciplina</th><th>Curso</th><th>Ausências</th><th>Freq.</th></tr></thead><tbody id="disciplines-table"></tbody></table>
      </article>
    </section>

    <section class="panel" style="margin-top:16px;">
      <h2>Evolução diária de presença e ausência</h2>
      <div class="timeline" id="timeline"></div>
    </section>

    <div class="foot">Quando publicado em HTTP/HTTPS, este HTML tenta ler automaticamente o arquivo <code>dashboard_data.json</code>.</div>
  </div>

  <script>
    const embeddedData = {payload_json};
    const number = (value) => new Intl.NumberFormat('pt-BR').format(value || 0);
    const percent = (value) => ((value || 0) * 100).toFixed(1).replace('.', ',') + '%';

    async function loadData() {{
      if (window.location.protocol.startsWith('http')) {{
        try {{
          const response = await fetch('./dashboard_data.json', {{ cache: 'no-store' }});
          if (response.ok) return await response.json();
        }} catch (error) {{
          console.warn('Fallback para dados embutidos.', error);
        }}
      }}
      return embeddedData;
    }}

    function normalizeDate(value) {{
      return value ? new Date(value + 'T00:00:00') : null;
    }}

    function aggregate(records, keys) {{
      const grouped = new Map();
      records.forEach((record) => {{
        const id = keys.map((key) => record[key]).join('||');
        if (!grouped.has(id)) {{
          const item = {{}};
          keys.forEach((key) => item[key] = record[key]);
          item.Registros = 0;
          item.Presencas = 0;
          item.Ausencias = 0;
          grouped.set(id, item);
        }}
        const item = grouped.get(id);
        item.Registros += 1;
        item.Presencas += record.STATUS === 'Presente' ? 1 : 0;
        item.Ausencias += record.STATUS === 'Ausente' ? 1 : 0;
      }});
      return Array.from(grouped.values()).map((item) => ({{
        ...item,
        Frequencia: item.Registros ? item.Presencas / item.Registros : 0,
      }}));
    }}

    function matchesFrequency(value, mode) {{
      if (mode === 'below_75') return value < 0.75;
      return value >= Number(mode || '0');
    }}

    loadData().then((data) => {{
      const records = data.records;
      const ids = {{
        course: document.getElementById('course-filter'),
        shift: document.getElementById('shift-filter'),
        discipline: document.getElementById('discipline-filter'),
        student: document.getElementById('student-filter'),
        start: document.getElementById('start-date'),
        end: document.getElementById('end-date'),
        freq: document.getElementById('frequency-filter'),
        weekday: document.getElementById('weekday-filter'),
      }};

      const uniqueCourses = ['Todos', ...Array.from(new Set(records.map((row) => row.CURSO)))];
      const uniqueShifts = ['Todos', ...Array.from(new Set(records.map((row) => row.TURNO)))];
      ids.course.innerHTML = uniqueCourses.map((item) => `<option value="${{item}}">${{item}}</option>`).join('');
      ids.shift.innerHTML = uniqueShifts.map((item) => `<option value="${{item}}">${{item}}</option>`).join('');
      document.getElementById('gerado-em').textContent = data.meta.gerado_em;
      const orderedDates = records.map((row) => row.DATA).sort();
      ids.start.value = orderedDates[0] || '';
      ids.end.value = orderedDates[orderedDates.length - 1] || '';

      function render() {{
        const filtered = records.filter((row) => {{
          const rowDate = normalizeDate(row.DATA);
          return (ids.course.value === 'Todos' || row.CURSO === ids.course.value)
            && (ids.shift.value === 'Todos' || row.TURNO === ids.shift.value)
            && (!ids.discipline.value.trim() || row.DISCIPLINA.toLowerCase().includes(ids.discipline.value.trim().toLowerCase()))
            && (!ids.student.value.trim() || row.ALUNO.toLowerCase().includes(ids.student.value.trim().toLowerCase()))
            && (!ids.start.value || rowDate >= normalizeDate(ids.start.value))
            && (!ids.end.value || rowDate <= normalizeDate(ids.end.value))
            && (ids.weekday.value === 'Todos' || row.DIA_SEMANA === ids.weekday.value);
        }});

        let courseSummary = aggregate(filtered, ['CURSO']).sort((a, b) => a.CURSO.localeCompare(b.CURSO, 'pt-BR'));
        let studentSummary = aggregate(filtered, ['CURSO', 'ALUNO']).sort((a, b) => b.Ausencias - a.Ausencias || a.ALUNO.localeCompare(b.ALUNO, 'pt-BR'));
        let disciplineSummary = aggregate(filtered, ['CURSO', 'DISCIPLINA']).sort((a, b) => b.Ausencias - a.Ausencias || a.DISCIPLINA.localeCompare(b.DISCIPLINA, 'pt-BR'));

        courseSummary = courseSummary.filter((row) => matchesFrequency(row.Frequencia, ids.freq.value));
        studentSummary = studentSummary.filter((row) => matchesFrequency(row.Frequencia, ids.freq.value));
        disciplineSummary = disciplineSummary.filter((row) => matchesFrequency(row.Frequencia, ids.freq.value));

        const freqGeral = filtered.length ? filtered.filter((row) => row.STATUS === 'Presente').length / filtered.length : 0;
        document.getElementById('registros').textContent = number(filtered.length);
        document.getElementById('frequencia').textContent = percent(freqGeral);
        document.getElementById('alunos').textContent = number(new Set(filtered.map((row) => row.CURSO + '::' + row.ALUNO)).size);
        document.getElementById('disciplinas').textContent = number(new Set(filtered.map((row) => row.CURSO + '::' + row.DISCIPLINA)).size);
        document.getElementById('presencas').textContent = number(filtered.filter((row) => row.STATUS === 'Presente').length);
        document.getElementById('ausencias').textContent = number(filtered.filter((row) => row.STATUS === 'Ausente').length);
        document.getElementById('periodo').textContent = `${{ids.start.value || data.meta.periodo_inicio}} a ${{ids.end.value || data.meta.periodo_fim}}`;
        document.getElementById('cursos-total').textContent = number(courseSummary.length);

        const courseBars = document.getElementById('course-bars');
        courseBars.innerHTML = '';
        courseSummary.forEach((row) => {{
          const studentsInCourse = new Set(filtered.filter((item) => item.CURSO === row.CURSO).map((item) => item.ALUNO)).size;
          const node = document.createElement('div');
          node.className = 'bar-row';
          node.innerHTML = `
            <div><strong>${{row.CURSO}}</strong><div class="label">${{number(studentsInCourse)}} alunos</div></div>
            <div class="bar-track"><div class="bar-fill" style="width:${{Math.max(row.Frequencia * 100, 4)}}%"></div></div>
            <div style="text-align:right;font-weight:700;">${{percent(row.Frequencia)}}</div>
          `;
          courseBars.appendChild(node);
        }});

        const studentsTable = document.getElementById('students-table');
        studentsTable.innerHTML = '';
        studentSummary.slice(0, 12).forEach((row) => {{
          const tr = document.createElement('tr');
          tr.innerHTML = `<td>${{row.ALUNO}}</td><td>${{row.CURSO}}</td><td>${{row.Ausencias}}</td><td>${{percent(row.Frequencia)}}</td>`;
          studentsTable.appendChild(tr);
        }});

        const disciplinesTable = document.getElementById('disciplines-table');
        disciplinesTable.innerHTML = '';
        disciplineSummary.slice(0, 12).forEach((row) => {{
          const tr = document.createElement('tr');
          tr.innerHTML = `<td>${{row.DISCIPLINA}}</td><td>${{row.CURSO}}</td><td>${{row.Ausencias}}</td><td>${{percent(row.Frequencia)}}</td>`;
          disciplinesTable.appendChild(tr);
        }});

        const timelineData = Array.from(filtered.reduce((map, row) => {{
          if (!map.has(row.DATA_EXIBICAO)) map.set(row.DATA_EXIBICAO, {{ DATA_EXIBICAO: row.DATA_EXIBICAO, Presente: 0, Ausente: 0 }});
          map.get(row.DATA_EXIBICAO)[row.STATUS] += 1;
          return map;
        }}, new Map()).values());

        const maxTotal = Math.max(...timelineData.map((row) => row.Presente + row.Ausente), 1);
        const timeline = document.getElementById('timeline');
        timeline.innerHTML = '';
        timelineData.forEach((row) => {{
          const presentHeight = ((row.Presente || 0) / maxTotal) * 220;
          const absentHeight = ((row.Ausente || 0) / maxTotal) * 220;
          const node = document.createElement('div');
          node.className = 'timeline-col';
          node.innerHTML = `
            <div class="stack" title="Presenças: ${{row.Presente}} | Ausências: ${{row.Ausente}}">
              <div class="present" style="height:${{presentHeight}}px"></div>
              <div class="absent" style="height:${{absentHeight}}px"></div>
            </div>
            <div class="date-label">${{row.DATA_EXIBICAO}}</div>
          `;
          timeline.appendChild(node);
        }});
      }}

      Object.values(ids).forEach((element) => element.addEventListener(element.tagName === 'INPUT' && element.type === 'text' ? 'input' : 'change', render));
      render();
    }});
  </script>
</body>
</html>
"""


def main() -> None:
    source_file = find_source_file()
    data = load_data(source_file)
    summary = summarize_data(data)
    payload = build_payload(data, summary)
    OUTPUT_JSON.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    OUTPUT_HTML.write_text(render_html(payload), encoding="utf-8")
    print(json.dumps(
        {
            "source": str(source_file),
            "html": str(OUTPUT_HTML),
            "json": str(OUTPUT_JSON),
        },
        ensure_ascii=False,
        indent=2,
    ))


if __name__ == "__main__":
    main()
