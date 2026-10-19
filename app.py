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
            st.write("A base contém 313 combinações comparáveis e 209 pendentes. O indicador é parcial.")
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
