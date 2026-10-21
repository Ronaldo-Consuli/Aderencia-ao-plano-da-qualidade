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
    rules = [dict(r) for r in rules]
    for rule in rules:
        if rule["id"].startswith("LEV"):
            keys = list(rule.get("observed_keys", []))
            for rec in day:
                material = normalize(rec["Material"])
                if (re.fullmatch(r"levedurainativaseca[0-9]*", material)
                        and normalize(rec["Grupo"]) == normalize(rule["group"])
                        and normalize(rec["Equipamento"]) == "silo"
                        and normalize(rec["Local"]) == normalize(rule["local"])
                        and normalize(rec["Característica"]) == normalize(rule["test"])
                        and normalize(rec["Tipo de Análise"]) == "composta"):
                    keys.append([rec[f] for f in KEY_FIELDS])
            rule["observed_keys"] = keys
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
        st.caption("Somente laboratório interno. Proteína a cada 2 horas (12/dia); umidade, pH e densidade aparente a cada 8 horas (3/dia cada). Total: 21 resultados previstos em 24 horas, independentemente da faixa. As faixas de proteína são especificações e não seis ensaios diferentes.")
        st.dataframe(pd.DataFrame([{"Ensaio": r["plan_test"], "Frequência": r["frequency"],
            "Previsto em 24h": r["expected"], "Condição": r["status"],
            "Validação necessária": r["notes"]} for r in yeast]), hide_index=True, width="stretch")
        st.info("Vínculos validados. Os resultados das faixas de levedura são somados por característica, sem multiplicar a quantidade prevista pelo número de faixas. Registros repetidos no mesmo ponto, característica e horário continuam sujeitos à revisão.")
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
