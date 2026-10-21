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
        count_each_row = rule["id"].startswith("LEV")
        if count_each_row:
            note += " Cada linha com resultado conta como análise, inclusive horários repetidos ou próximos, conforme orientação do usuário."
        if not count_each_row and status == "Comparável" and any(n > 1 for n in counts.values()):
            status = "Pendente"
            note = "Chave/horário repetido no monitor. Confirmar correção, reanálise ou amostra distinta."
        item = {"rule": rule, "match": match, "count": len(match) if count_each_row else len(counts), "status": status, "note": note}
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
        st.caption("Somente laboratório interno. Proteína, umidade, pH e densidade aparente a cada 2 horas (12/dia cada). Total: 48 resultados previstos em 24 horas, independentemente da faixa. As faixas de proteína são especificações e não seis ensaios diferentes.")
        st.dataframe(pd.DataFrame([{"Ensaio": r["plan_test"], "Frequência": r["frequency"],
            "Previsto em 24h": r["expected"], "Condição": r["status"],
            "Validação necessária": r["notes"]} for r in yeast]), hide_index=True, width="stretch")
        st.info("Vínculos validados. Os resultados das faixas de levedura são somados por característica, sem multiplicar a quantidade prevista pelo número de faixas. Cada linha com resultado conta como uma análise, mesmo com horário repetido ou próximo.")
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
