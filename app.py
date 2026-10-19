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

BASE_REGRAS_COMPACTADA = "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"
VERSION_BASE = "Planos industriais + plano de cana incluído em 19/10/2026 · vínculos validados e pendentes"
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
        st.caption("16 itens por ensaio/ponto · 67 resultados previstos em 24 horas. Índice de preparo: 6 por desintegrador. Somente vínculos confirmados entram no percentual.")
        st.dataframe(pd.DataFrame([{"Material do plano": r["plan_material"],
            "Ensaio": r["plan_test"], "Equipamento": r["equipment"], "Condição": r["status"], "Frequência": r["frequency"],
            "Previsto em 24h": r["expected"], "Coleta": r["sample_type"],
            "Validação necessária": r["notes"]} for r in cana]), hide_index=True, width="stretch")
        st.caption("Fora do total diário fixo: análises por carga amostrada, por frente de colheita e uma vez na safra. Comparações NIR não são automaticamente resultados da rotina.")
    content = uploaded.getvalue() if uploaded is not None else monitor_bytes
    if content is None:
        st.info("Importe o monitor de digitação para visualizar os cartões e a aderência por grupo.")
        with st.expander("Critérios da avaliação"):
            st.write("Frequências diárias: 2h = 12; 4h = 6; 8h = 3; diária = 1, além de 2 ou 3 vezes ao dia.")
            st.write("A base contém 305 combinações comparáveis e 217 pendentes. O indicador é parcial.")
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
