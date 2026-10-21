"""Aderência analítica diária — Streamlit.

Executar: streamlit run app.py
Base incorporada: comparação dos planos recebidos em 16/10/2026.
Não contém resultados, usuários ou dados de produção do monitor de exemplo.
Regras pendentes continuam fora do indicador. Atualizar BASE_REGRAS_COMPACTADA para rever
os vínculos ou incorporar novos pontos, versões e frequências do plano.
"""
from __future__ import annotations

import base64
from collections import Counter, defaultdict
from datetime import date, datetime
from html import escape
from io import BytesIO
import json
import re
import unicodedata
import zlib

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.datetime import from_excel

BASE_REGRAS_COMPACTADA = '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'
VERSION_BASE = "Planos industriais + cana + levedura interna (AN.SMI.PLV.01.13) · vínculos validados e pendentes"
REQUIRED = ["Grupo", "Material", "Equipamento", "Local", "Característica",
            "Tipo de Análise", "Dt. Resultado", "Vl. Resultado"]
KEY_FIELDS = REQUIRED[:6]
ITEM_COLS = ["ID", "ID plano", "Grupo", "Material", "Equipamento", "Local",
             "Característica", "Tipo de análise", "Previsto", "Realizado",
             "Faltantes", "Excedentes", "Atendidos", "Aderência (%)",
             "Condição", "Situação", "Observação", "Linhas do monitor",
             "Arquivo do plano", "Referência no plano"]
SUMMARY_COLS = ["Grupo", "Previsto", "Realizado", "Faltantes", "Excedentes",
                "Atendidos", "Aderência (%)", "Itens comparáveis",
                "Itens pendentes", "Registros no monitor", "Registros vinculados",
                "Registros pendentes", "Registros sem vínculo"]


def normalize(value) -> str:
    text = unicodedata.normalize("NFKD", str(value or ""))
    text = text.encode("ascii", "ignore").decode().lower()
    return re.sub(r"[^a-z0-9]+", "", text)


def load_rules() -> list[dict]:
    return json.loads(zlib.decompress(base64.b64decode(BASE_REGRAS_COMPACTADA)))


def is_blank(value) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def parse_datetime(value, epoch):
    if is_blank(value):
        return None
    if isinstance(value, (datetime, date)):
        return pd.Timestamp(value).to_pydatetime().replace(tzinfo=None)
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        try:
            converted = from_excel(value, epoch)
            return converted if isinstance(converted, datetime) else None
        except (ValueError, OverflowError):
            return None
    text = str(value).strip()
    for fmt in ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%d/%m/%Y",
                "%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def read_monitor(content: bytes) -> tuple[list[dict], list[dict], str]:
    """Lê uma única aba de monitor, detecta cabeçalho e preserva zeros/textos."""
    if len(content) > 30 * 1024 * 1024:
        raise ValueError("O arquivo excede 30 MB. Exporte um período menor.")
    try:
        wb = load_workbook(BytesIO(content), data_only=True, read_only=True)
    except Exception as exc:
        raise ValueError("Não foi possível abrir o Excel. Envie o monitor no formato .xlsx.") from exc
    try:
        matches = []
        required_names = {normalize(c) for c in REQUIRED}
        for sheet in wb:
            for rownum, row in enumerate(sheet.iter_rows(max_row=20, values_only=True), 1):
                names = [normalize(v) for v in row]
                if required_names.issubset(names):
                    if any(names.count(name) > 1 for name in required_names):
                        raise ValueError("Há colunas obrigatórias duplicadas no cabeçalho.")
                    matches.append((sheet, rownum, names))
                    break
        if not matches:
            raise ValueError("Cabeçalho não encontrado. Colunas necessárias: " + ", ".join(REQUIRED))
        if len(matches) != 1:
            raise ValueError("Há mais de uma aba de monitor. Envie uma exportação com apenas uma aba de resultados.")
        sheet, header_row, names = matches[0]
        positions = {c: names.index(normalize(c)) for c in REQUIRED}
        extra = {c: names.index(normalize(c)) for c in ["Dt. Registro", "Usuário"] if normalize(c) in names}
        records, invalid = [], []
        for rownum, row in enumerate(sheet.iter_rows(min_row=header_row + 1, values_only=True), header_row + 1):
            if all(is_blank(v) for v in row):
                continue
            rec = {c: row[i] if i < len(row) else None for c, i in {**positions, **extra}.items()}
            rec["Linha na origem"] = rownum
            dt = parse_datetime(rec["Dt. Resultado"], wb.epoch)
            missing = [c for c in KEY_FIELDS if is_blank(rec[c])]
            reasons = []
            if missing:
                reasons.append("Identificação incompleta: " + ", ".join(missing))
            if dt is None:
                reasons.append("Data do resultado inválida ou vazia")
            if is_blank(rec["Vl. Resultado"]):
                reasons.append("Resultado vazio (não conta como realizado)")
            if reasons:
                invalid.append({**rec, "Motivo": "; ".join(reasons)})
                continue
            for field in KEY_FIELDS:
                rec[field] = str(rec[field]).strip()
            rec["Dt. Resultado"] = dt
            rec["Data"] = dt.date()
            rec["_key"] = tuple(normalize(rec[f]) for f in KEY_FIELDS)
            records.append(rec)
        return records, invalid, sheet.title
    finally:
        wb.close()


def rule_keys(rule: dict) -> set[tuple]:
    base = tuple(normalize(rule[k]) for k in
                 ["group", "material", "equipment", "local", "test", "sample_type"])
    keys = {base}
    keys.update(tuple(normalize(v) for v in key) for key in rule.get("observed_keys", []))
    return keys


def sugar_product(value):
    """Identifica a classe sem confundir VVHP com VHP."""
    value = normalize(value)
    if "vvhp" in value:
        return "VVHP"
    return "VHP" if "vhp" in value else None


def select_sugar_rules(rules, day):
    products = {sugar_product(r["Material"]) for r in day
                if normalize(r["Grupo"]) == normalize("Produção de Açúcar")}
    products.discard(None)
    active = next(iter(products)) if len(products) == 1 else None
    selected, excluded = [], []
    for original in rules:
        rule = dict(original)
        product = sugar_product(rule.get("plan_material", ""))
        if product:
            rule["group"] = "Produção de Açúcar"
            if active and product != active:
                excluded.append({"ID": rule["id"], "Produto": product,
                    "Ensaio": rule["plan_test"], "Ponto": rule["local"],
                    "Previsto original (24h)": rule["expected"],
                    "Situação": "Não aplicável", "Motivo": f"Produção exclusiva de {active} no dia"})
                continue
            if not active:
                rule["status"] = "Pendente"
                rule["notes"] = "Conferir produto do dia: ambos ou nenhum identificado em Produção de Açúcar. " + rule["notes"]
        selected.append(rule)
    return selected, excluded, active, sorted(products)


def evaluate(records: list[dict], selected_date: date, rules=None) -> dict:
    """Apura cada obrigação antes de somar. Nunca compensa faltas com excedentes."""
    rules = load_rules() if rules is None else rules
    day = [r for r in records if r["Data"] == selected_date]
    rules, not_applicable, active_sugar, detected_sugars = select_sugar_rules(rules, day)
    by_key = defaultdict(list)
    for rec in day:
        by_key[rec["_key"]].append(rec)
    prepared = []
    owners = defaultdict(list)
    for rule in rules:
        matched = {r["Linha na origem"]: r for key in rule_keys(rule) for r in by_key.get(key, [])}
        match = list(matched.values())
        counts = Counter((r["_key"], r["Dt. Resultado"]) for r in match)
        status = rule["status"]
        note = rule["notes"]
        if status == "Comparável" and any(n > 1 for n in counts.values()):
            status = "Pendente"
            note = "Chave/horário repetido no monitor. Confirmar correção, reanálise ou amostra distinta."
        item = {"rule": rule, "match": match, "count": len(counts), "status": status, "note": note}
        prepared.append(item)
        if status == "Comparável":
            for rec in match:
                owners[rec["Linha na origem"]].append(item)
    for items in owners.values():
        if len(items) > 1:
            for item in items:
                item["status"] = "Pendente"
                item["note"] = "Registro associado a mais de uma obrigação. Rever vínculo antes do cálculo."

    rows, associations = [], defaultdict(list)
    for item in prepared:
        rule, actual, status = item["rule"], item["count"], item["status"]
        planned = int(rule["expected"])
        comparable = status == "Comparável"
        fulfilled = min(planned, actual) if comparable else None
        rows.append({
            "ID": rule["id"], "ID plano": rule["plan_id"], "Grupo": rule["group"],
            "Material": rule["material"], "Equipamento": rule["equipment"], "Local": rule["local"],
            "Característica": rule["test"], "Tipo de análise": rule["sample_type"],
            "Previsto": planned, "Realizado": actual,
            "Faltantes": max(planned-actual, 0) if comparable else None,
            "Excedentes": max(actual-planned, 0) if comparable else None,
            "Atendidos": fulfilled,
            "Aderência (%)": fulfilled/planned*100 if comparable and planned else None,
            "Condição": status,
            "Situação": ("A menos" if actual < planned else "A mais" if actual > planned
                         else "Na quantidade prevista") if comparable else "Pendente de validação",
            "Observação": item["note"],
            "Linhas do monitor": ", ".join(str(r["Linha na origem"]) for r in item["match"]),
            "Arquivo do plano": rule["source_file"], "Referência no plano": rule["source_rows"],
        })
        for rec in item["match"]:
            associations[rec["Linha na origem"]].append((rule["id"], status))

    linked_rows = []
    for rec in day:
        links = associations[rec["Linha na origem"]]
        state = "Comparável" if any(s == "Comparável" for _, s in links) else "Pendente" if links else "Sem vínculo no recorte"
        linked_rows.append({**{k: v for k, v in rec.items() if not k.startswith("_")},
                            "IDs das combinações": "; ".join(i for i, _ in links), "Vínculo": state})
    detail = pd.DataFrame(rows, columns=ITEM_COLS)
    linked = pd.DataFrame(linked_rows)
    groups = sorted(({r["Grupo"] for r in day} | {r["group"] for r in rules}) - {"Sem grupo confirmado"})
    summaries = []
    for group in groups:
        items = detail[detail["Grupo"] == group]
        ok = items[items["Condição"] == "Comparável"]
        recs = [r for r in linked_rows if r["Grupo"] == group]
        n = int(ok["Previsto"].sum())
        summaries.append({"Grupo": group,
            **{col: int(ok[col].sum()) if n else None for col in
               ["Previsto", "Realizado", "Faltantes", "Excedentes", "Atendidos"]},
            "Aderência (%)": float(ok["Atendidos"].sum())/n*100 if n else None,
            "Itens comparáveis": len(ok), "Itens pendentes": int((items["Condição"] == "Pendente").sum()),
            "Registros no monitor": len(recs),
            "Registros vinculados": sum(r["Vínculo"] == "Comparável" for r in recs),
            "Registros pendentes": sum(r["Vínculo"] == "Pendente" for r in recs),
            "Registros sem vínculo": sum(r["Vínculo"] == "Sem vínculo no recorte" for r in recs)})
    summary = pd.DataFrame(summaries, columns=SUMMARY_COLS)
    return {"summary": summary, "detail": detail, "linked": linked,
            "no_group": detail[detail["Grupo"] == "Sem grupo confirmado"], "date": selected_date,
            "not_applicable": pd.DataFrame(not_applicable), "active_sugar": active_sugar,
            "detected_sugars": detected_sugars}


def total_row(summary):
    result = {"Grupo": "TOTAL PARCIAL"}
    for col in SUMMARY_COLS[1:]:
        if col != "Aderência (%)":
            result[col] = int(summary[col].sum())
    result["Aderência (%)"] = result["Atendidos"]/result["Previsto"]*100 if result["Previsto"] else None
    return result


def csv_bytes(frame):
    # Evita que textos de um arquivo enviado sejam interpretados como fórmulas.
    safe = frame.copy()
    for col in safe.columns:
        safe[col] = safe[col].map(lambda v: "'" + v if isinstance(v, str) and v.startswith(("=", "+", "-", "@")) else v)
    return safe.to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig")


def percent(value):
    return "—" if value is None or pd.isna(value) else f"{value:.1f}%".replace(".", ",")


CSS = """
<style>
/* Paleta do aplicativo de referência Carta controle NIR DS-3. */
[data-testid="stAppViewContainer"] {background:#e8edf3; color:#083050; color-scheme:light;}
[data-testid="stHeader"] {background:#e8edf3;}
[data-testid="stSidebar"] {background:#f5f7fa; border-right:1px solid #dce3eb; color:#083050;}
.block-container {padding-top:4.5rem; padding-bottom:2.2rem; max-width:1600px;}
.qmp-header {background:linear-gradient(118deg,#00983a 0%,#004881 64%,#083050 100%);
 border-radius:12px; padding:24px 34px 20px; margin-bottom:6px; box-shadow:0 3px 10px #08305012;}
.qmp-header h1 {color:#fff !important; font-size:31px; font-weight:750; line-height:1.25; padding:0; margin:0;}
.qmp-header p {color:#eaf5fa; font-size:14px; margin:8px 0 0;}
h2,h3 {color:#083050 !important;}
[data-testid="stAlert"] {border-radius:9px;}
[data-testid="stExpander"] {background:#f5f7fa; border-radius:12px;}
[data-testid="stFileUploaderDropzone"] {background:#fff; border:1px dashed #bccbd8; border-radius:10px;}
[data-testid="stDownloadButton"] button {background:#004881; color:#fff; border:1px solid #004881; border-radius:8px;}
[data-testid="stDownloadButton"] button:hover {background:#083050; color:#fff; border-color:#083050;}
[data-testid="stDownloadButton"] button:focus-visible {outline:3px solid #00983a; outline-offset:3px;}
[data-baseweb="tab"] {color:#083050;}
[data-baseweb="tab"][aria-selected="true"] {color:#007c30; font-weight:700;}
[data-baseweb="tab-highlight"] {background:#00983a;}
.group-grid {display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:16px; margin:20px 0 28px;}
.group-card {background:#fff; color:#083050; border:1px solid #d7e0e9;
 border-top:4px solid #004881; border-radius:12px; padding:20px 16px; text-align:center;
 min-height:185px; box-shadow:0 3px 12px #0830500b;}
.group-card.complete {border-top-color:#00983a;}
.group-card.complete .group-value {color:#007c30;}
.group-card.empty {border-top-color:#9aaaba; background:#f5f7fa;}
.group-card.empty .group-value {color:#64778a;}
.group-title {font-weight:650; min-height:46px; font-size:15px;}
.group-value {color:#004881; font-size:36px; font-weight:750; line-height:1.3; margin:7px 0;}
.group-note {color:#506579; font-size:12px; line-height:1.7;}
@media(max-width:1050px){.group-grid{grid-template-columns:repeat(3,minmax(0,1fr));}}
@media(max-width:760px){.group-grid{grid-template-columns:repeat(2,minmax(0,1fr));}
 .qmp-header{padding:22px;} .qmp-header h1{font-size:26px;}}
@media(max-width:440px){.group-grid{grid-template-columns:1fr;}}
</style>
"""


def style_summary(table):
    """Apresentação da tabela; não altera os valores calculados."""
    def row_colors(row):
        if row["Grupo"] == "TOTAL PARCIAL":
            return ["background-color:#083050;color:#ffffff;font-weight:bold"] * len(row)
        base = "background-color:#ffffff;color:#083050"
        styles = [base] * len(row)
        for column in ("Faltantes", "Excedentes", "Aderência (%)"):
            value = row[column]
            if pd.isna(value):
                color = "background-color:#edf1f5;color:#64778a"
            elif column == "Faltantes" and value > 0:
                color = "background-color:#fff3de;color:#805200"
            elif column == "Excedentes" and value > 0:
                color = "background-color:#e3eff9;color:#004881"
            elif column == "Aderência (%)" and value >= 100:
                color = "background-color:#e2f3e8;color:#007c30;font-weight:bold"
            else:
                color = base
            styles[row.index.get_loc(column)] = color
        return styles
    return table.style.apply(row_colors, axis=1)


def cards_html(summary):
    cards = []
    for row in summary.to_dict("records"):
        value = row["Aderência (%)"]
        empty = pd.isna(value)
        cls = "empty" if empty else "complete" if value >= 100 else ""
        note = ("Vínculos aguardando validação" if row["Itens pendentes"] else "Sem base diária vinculada") if empty else (
            f"{int(row['Realizado'])} realizados / {int(row['Previsto'])} previstos<br>"
            f"{int(row['Faltantes'])} faltantes · {int(row['Excedentes'])} excedentes")
        cards.append(f'<div class="group-card {cls}"><div class="group-title">{escape(row["Grupo"])}</div>'
                     f'<div class="group-value">{percent(value)}</div><div class="group-note">{note}<br>'
                     f'{int(row["Itens comparáveis"])} itens comparáveis · {int(row["Itens pendentes"])} pendentes'
                     '</div></div>')
    return '<div class="group-grid">' + ''.join(cards) + '</div>'


def main(monitor_bytes=None):
    import streamlit as st

    st.set_page_config(page_title="Aderência Analítica", page_icon="📋", layout="wide")
    st.markdown(CSS, unsafe_allow_html=True)
    st.markdown('<header class="qmp-header"><h1>Aderência Analítica do Laboratório</h1>'
                '<p>Previsto × realizado · Rotina diária · 24 horas</p></header>', unsafe_allow_html=True)
    with st.sidebar:
        st.header("Monitor de digitação")
        uploaded = st.file_uploader("Importar monitor (.xlsx)", type=["xlsx"])
        st.caption("A base dos planos já está incorporada ao aplicativo.")
    with st.expander("Plano diário — Laboratório de Cana"):
        cana = [r for r in load_rules() if r["group"] == "Laboratório de Cana"]
        st.caption("14 itens por ensaio/ponto · 49 resultados previstos em 24 horas. Índice de preparo: 1 por turno de cada CP, considerando 3 turnos/dia. Somente vínculos confirmados entram no percentual.")
        st.dataframe(pd.DataFrame([{"Material do plano": r["plan_material"],
            "Ensaio": r["plan_test"], "Equipamento": r["equipment"], "Condição": r["status"], "Frequência": r["frequency"],
            "Previsto em 24h": r["expected"], "Coleta": r["sample_type"],
            "Validação necessária": r["notes"]} for r in cana]), hide_index=True, width="stretch")
        st.caption("Impureza mineral e Bolo úmido excluídos conforme orientação do usuário. Fora do total diário fixo: análises por carga amostrada, por frente de colheita e uma vez na safra. Comparações NIR não são automaticamente resultados da rotina.")
    with st.expander("Plano diário — Produção de Levedura"):
        yeast = [r for r in load_rules() if r["id"].startswith("LEV")]
        st.caption("Somente laboratório interno. Quatro ensaios diários; 21 resultados por ciclo de 24h do plano, antes de definir a abrangência por produto. As faixas de proteína são especificações e não seis ensaios diferentes.")
        st.dataframe(pd.DataFrame([{"Ensaio": r["plan_test"], "Frequência": r["frequency"],
            "Previsto em 24h no plano": r["expected"], "Condição": r["status"],
            "Validação necessária": r["notes"]} for r in yeast]), hide_index=True, width="stretch")
        st.info("Confirmar frequência conjunta ou por produto 400/420 e vínculos com o monitor. Até essa validação, os quatro itens permanecem pendentes e não alteram o percentual.")
        st.caption("Excluídos por orientação do usuário: Odor/Aroma, Cor (Pantone) e Corpos estranhos. Fora do cálculo diário: Cinzas (Matéria Mineral), de frequência semanal, e todos os ensaios de laboratórios externos. Plano AN.SMI.PLV.01.13, página 1. As regras anteriores de creme de levedura permanecem separadas.")
    content = uploaded.getvalue() if uploaded is not None else monitor_bytes
    if content is None:
        st.info("Importe o monitor de digitação para visualizar os cartões e a aderência por grupo.")
        with st.expander("Critérios da avaliação"):
            st.write("Frequências diárias: 2h = 12; 4h = 6; 8h = 3; diária = 1, além de 2 ou 3 vezes ao dia.")
            counts = Counter(r["status"] for r in load_rules())
            st.write(f"A base contém {counts['Comparável']} combinações comparáveis e {counts['Pendente']} pendentes. O indicador é parcial.")
        return
    try:
        cached_read = st.cache_data(show_spinner=False)(read_monitor)
        records, invalid, sheet_name = cached_read(content)
    except ValueError as exc:
        st.error(str(exc))
        return
    if not records:
        st.error("Nenhum resultado com identificação, valor e data válidos foi encontrado.")
        if invalid:
            st.dataframe(pd.DataFrame(invalid), hide_index=True, width="stretch")
        return
    dates = sorted({r["Data"] for r in records}, reverse=True)
    with st.sidebar:
        selected_date = st.selectbox("Data do resultado", dates, format_func=lambda d:d.strftime("%d/%m/%Y"))
    result = evaluate(records, selected_date)
    summary, detail, linked = result["summary"], result["detail"], result["linked"]
    if result["active_sugar"]:
        st.success(f"Produção de açúcar em {selected_date:%d/%m/%Y}: {result['active_sugar']}. "
                   "O outro produto está como Não aplicável, fora do previsto e dos faltantes.")
    else:
        st.warning("Conferir produção de açúcar: " +
                   ("VHP e VVHP encontrados no mesmo dia." if result["detected_sugars"] else
                    "nenhum VHP/VVHP identificado no grupo Produção de Açúcar.") +
                   " As obrigações específicas dos produtos ficam pendentes, fora do percentual, até a conferência.")
    if not result["not_applicable"].empty:
        with st.expander("Análises não aplicáveis — produto não produzido"):
            st.dataframe(result["not_applicable"], hide_index=True, width="stretch")
    with st.sidebar:
        selected_group = st.selectbox("Grupo", ["Todos"] + summary["Grupo"].tolist())
        st.caption(f"Data analisada: {selected_date:%d/%m/%Y}\n\nOperação considerada: 00h às 24h")
    filtered = summary if selected_group == "Todos" else summary[summary["Grupo"] == selected_group]
    filtered_detail = detail if selected_group == "Todos" else detail[detail["Grupo"] == selected_group]
    filtered_linked = linked if selected_group == "Todos" else linked[linked["Grupo"] == selected_group]
    total = total_row(filtered)
    st.info(f"Aderência parcial: {percent(total['Aderência (%)'])}. "
            "Calculada somente nos itens comparáveis; excedentes não compensam faltantes.")
    st.markdown(cards_html(filtered), unsafe_allow_html=True)
    st.subheader("Previsto e realizado por grupo")
    display_cols = ["Grupo", "Previsto", "Realizado", "Faltantes", "Excedentes", "Aderência (%)"]
    table = pd.concat([filtered, pd.DataFrame([total])], ignore_index=True)
    configurations = {col: st.column_config.NumberColumn(col, format="%d") for col in display_cols[1:-1]}
    configurations["Aderência (%)"] = st.column_config.NumberColumn("Aderência parcial (%)", format="%.1f%%")
    st.dataframe(style_summary(table[display_cols]), hide_index=True, width="stretch", height=min(620, 38*(len(table)+1)),
                 column_config=configurations)
    st.caption("Células vazias indicam grupo sem base diária vinculada. Mesmo 100% se refere apenas aos itens comparáveis.")
    st.download_button("Baixar tabela por grupo (CSV)", csv_bytes(table),
                       file_name=f"aderencia_grupos_{selected_date:%Y-%m-%d}.csv", mime="text/csv")
    tabs = st.tabs(["Análises por ponto", "Vínculos do monitor", "Pendências e critérios"])
    with tabs[0]:
        situation = st.selectbox("Situação", ["Todas", "A menos", "A mais", "Na quantidade prevista", "Pendente de validação"])
        shown = filtered_detail if situation == "Todas" else filtered_detail[filtered_detail["Situação"] == situation]
        st.dataframe(shown, hide_index=True, width="stretch")
        st.download_button("Baixar análises por ponto (CSV)", csv_bytes(shown),
                           file_name=f"aderencia_itens_{selected_date:%Y-%m-%d}.csv", mime="text/csv")
    with tabs[1]:
        st.caption("Cada registro conserva a linha do Excel original e os IDs dos itens relacionados.")
        state = st.selectbox("Vínculo", ["Todos", "Comparável", "Pendente", "Sem vínculo no recorte"])
        shown = filtered_linked if state == "Todos" else filtered_linked[filtered_linked["Vínculo"] == state]
        st.dataframe(shown, hide_index=True, width="stretch")
        st.download_button("Baixar vínculos do monitor (CSV)", csv_bytes(shown),
                           file_name=f"vinculos_monitor_{selected_date:%Y-%m-%d}.csv", mime="text/csv")
    with tabs[2]:
        coverage = filtered[["Grupo", "Itens comparáveis", "Itens pendentes", "Registros no monitor",
                            "Registros vinculados", "Registros pendentes", "Registros sem vínculo"]]
        st.dataframe(coverage, hide_index=True, width="stretch")
        st.caption(f"Existem ainda {len(result['no_group'])} combinações do plano sem grupo confirmado, fora dos percentuais por grupo.")
        if invalid:
            st.warning(f"{len(invalid)} linhas do arquivo foram excluídas por dados incompletos ou inválidos. Confira antes de concluir a avaliação.")
            st.dataframe(pd.DataFrame(invalid), hide_index=True, width="stretch")
            st.download_button("Baixar linhas inválidas (CSV)", csv_bytes(pd.DataFrame(invalid)),
                               file_name="linhas_invalidas.csv", mime="text/csv")
        st.markdown("""
- O previsto vem da frequência diária do plano, por ensaio, equipamento, local e tipo de coleta.
- A data de referência é **Dt. Resultado**. O fechamento é de 00h às 24h.
- Aderência = soma de **mínimo(previsto, realizado)** por item ÷ soma dos previstos comparáveis.
- VHP e VVHP são exclusivos por dia: a identificação usa Material no grupo Produção de Açúcar; expedição não define a produção. O produto inativo fica Não aplicável. Ambos ou nenhum exigem conferência.
- As regras e os pontos conhecidos permanecem previstos mesmo sem resultado no dia, exceto o produto de açúcar identificado como não produzido. Pressupõe-se 24 horas de operação.
- Faltantes significam registros não encontrados na exportação, não prova de análise não executada.
- Tipos de coleta diferentes, equivalências técnicas e versões divergentes ficam pendentes.
- Mesma identificação e horário repetidos exigem validação; o item sai do subtotal comparável.
- Resultado zero conta. Resultado vazio não conta. Não há avaliação dos valores contra limites de especificação.
- Pontos novos e registros sem vínculo não são automaticamente excedentes nem entram no previsto.
- Paradas, mudanças de versão, novos equipamentos e equivalências precisam de atualização da base antes de uma avaliação global.
- Os registros do monitor podem incluir cálculos e análises fora da rotina diária. O total do arquivo não é o denominador da aderência.
""")
        st.caption(f"{VERSION_BASE} · Aba importada: {sheet_name}")
    if invalid:
        st.warning("Há linhas inválidas no arquivo. Consulte a aba Pendências e critérios.")


if __name__ == "__main__":
    main()
