"""
Versao WEB (stlite / WebAssembly) do Gerador de Applications MFT (CA WA).

Corre 100% no browser - os ficheiros carregados nunca saem do PC do utilizador.
Usa a mesma logica da app desktop (mft_core.py); em vez de pastas, os EAI sao
carregados (XML soltos ou ZIP) e o resultado e descarregado num ZIP.
"""

import io
import json
import os
import zipfile

import pandas as pd
import streamlit as st

import mft_core as core


DEFAULT_SHEET = "CAWA_EAI_TO_MFT_APPLICATIONS"
DEFAULT_COLS = dict(as_is="A", job_name="B", app="L", args="N", job_id="G")
COL_LABELS = dict(
    as_is="Application (AS-IS)", job_name="Nome do Job (AS-IS)", app="Application (TO-BE)",
    args="Argumento2", job_id="Job_ID",
)
INFRA_LABELS = dict(
    agent=("Agente (rununit)", "ex.: NOME-DO-AGENTE"),
    snmp_agent=("SNMP agent (por defeito)", "ex.: NOME-SNMP-AGENT"),
    userid=("User id do job", "ex.: utilizador"),
    cmdname=("Comando (cmdname)", "ex.: /caminho/para/mft_script_submit.sh"),
)


def read_eai_uploads(uploads):
    """{nome_sem_extensao: (nome_ficheiro, texto)} a partir de XML soltos e/ou ZIPs."""
    sources = {}
    errors = []

    def add(filename, data):
        base = os.path.basename(filename)
        if not base.lower().endswith(".xml"):
            return
        try:
            text = data.decode("utf-8")
        except UnicodeDecodeError:
            errors.append(f"'{base}' nao esta em UTF-8, ignorado.")
            return
        sources[base[:-4]] = (base, text)

    for up in uploads or []:
        data = up.getvalue()
        if up.name.lower().endswith(".zip"):
            with zipfile.ZipFile(io.BytesIO(data)) as zf:
                for info in zf.infolist():
                    if not info.is_dir():
                        add(info.filename, zf.read(info))
        else:
            add(up.name, data)
    return sources, errors


def run_generation(raw_rows, cols, start_row, eai_sources, infra):
    """Gera tudo e devolve (resultado, linhas_de_log, bytes_do_zip)."""
    log_lines = []
    loader = (lambda name: eai_sources.get(name)) if eai_sources else None
    result = core.generate(raw_rows, cols, start_row, loader, "ficheiros EAI carregados", infra, log_lines.append)

    if result["manual_setup"]:
        log_lines.append(
            f"\n{len(result['manual_setup'])} alerta(s)/script(s) usados que nao vem no import - "
            f"lista exportada para '{core.MANUAL_SETUP_CSV}'."
        )
    if result["missing_eai"]:
        log_lines.append(
            f"\n{len(result['missing_eai'])} ficheiro(s) EAI em falta nos ficheiros carregados - "
            f"lista exportada para '{core.MISSING_EAI_CSV}'."
        )
    log_lines += core.summary_log_lines(result)

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, content in result["files"].items():
            # mesmo fim de linha que a app desktop em Windows
            data = content if name.endswith(".csv") else content.replace("\n", "\r\n")
            zf.writestr(name, data.encode("utf-8"))
        zf.writestr("log_geracao.txt", "\n".join(log_lines).replace("\n", "\r\n").encode("utf-8"))
    return result, log_lines, buf.getvalue()


def main():
    st.set_page_config(page_title="Gerador de Applications MFT", page_icon="🔁", layout="wide")
    st.title("🔁 Gerador de Applications MFT (CA WA)")
    st.caption(
        "Corre no teu browser (nada é enviado para servidores). Lê o Excel de migração "
        "EAI → MFT e gera um XML por application MFT, com o horário, dependências e jobs "
        "ligados copiados do EAI de origem."
    )

    # --- configuracao de infraestrutura -----------------------------------
    for f in core.INFRA_FIELDS:
        st.session_state.setdefault(f"infra_{f}", "")
    infra_ok = all(st.session_state[f"infra_{f}"].strip() for f in core.INFRA_FIELDS)

    with st.expander("⚙️ Configuração do CAWA (agente, SNMP, user, comando)", expanded=not infra_ok):
        st.write(
            "Estes valores entram em todos os XML gerados. Carrega o ficheiro de configuração "
            "da equipa (`.json`) ou preenche à mão e descarrega-o para a próxima vez."
        )
        cfg_file = st.file_uploader("Carregar configuração (.json)", type=["json"], key="cfg_upload")
        if cfg_file is not None and st.session_state.get("cfg_loaded") != cfg_file.file_id:
            try:
                loaded = json.loads(cfg_file.getvalue().decode("utf-8-sig"))
                for f in core.INFRA_FIELDS:
                    if isinstance(loaded.get(f), str):
                        st.session_state[f"infra_{f}"] = loaded[f]
                st.session_state["cfg_loaded"] = cfg_file.file_id
                st.rerun()
            except (ValueError, UnicodeDecodeError) as e:
                st.error(f"Ficheiro de configuração inválido: {e}")
        c1, c2 = st.columns(2)
        for i, f in enumerate(core.INFRA_FIELDS):
            label, ph = INFRA_LABELS[f]
            with (c1 if i % 2 == 0 else c2):
                st.text_input(label, placeholder=ph, key=f"infra_{f}")
        infra = {f: st.session_state[f"infra_{f}"] for f in core.INFRA_FIELDS}
        st.download_button(
            "⬇️ Descarregar configuração (.json)",
            data=json.dumps(infra, ensure_ascii=False, indent=2).encode("utf-8"),
            file_name="config_gerador_mft.json", mime="application/json",
        )
    infra_ok = all(v.strip() for v in infra.values())

    # --- 1. Excel ---------------------------------------------------------
    st.subheader("1. Excel de migração (.xlsx ou .csv)")
    excel = st.file_uploader("Ficheiro", type=["xlsx", "xlsm", "csv"], key="excel")
    raw_rows = None
    if excel is not None:
        data = excel.getvalue()
        sheet = None
        if not excel.name.lower().endswith(".csv"):
            sheets = core.list_sheet_names(data)
            sheet = st.selectbox(
                "Folha", sheets, index=sheets.index(DEFAULT_SHEET) if DEFAULT_SHEET in sheets else 0,
            )
        try:
            raw_rows = core.load_raw_rows(data, sheet_name=sheet, filename=excel.name)
        except Exception as e:
            st.error(f"Erro ao ler o ficheiro: {e}")

    st.markdown("**Mapeamento de colunas** (letra da coluna no Excel)")
    cc = st.columns(6)
    col_letters = {}
    for i, key in enumerate(DEFAULT_COLS):
        with cc[i]:
            col_letters[key] = st.text_input(COL_LABELS[key], value=DEFAULT_COLS[key], key=f"col_{key}")
    with cc[5]:
        start_row = st.number_input("1ª linha de dados", min_value=1, value=2, step=1)

    if raw_rows:
        max_cols = max((len(r) for r in raw_rows), default=0)
        letters = [core.index_to_letter(i) for i in range(max_cols)]
        preview = pd.DataFrame(
            [list(r) + [""] * (max_cols - len(r)) for r in raw_rows[:200]],
            columns=letters, index=range(1, min(len(raw_rows), 200) + 1),
        )
        st.caption(f"{len(raw_rows)} linhas, {max_cols} colunas (pré-visualização das primeiras 200).")
        st.dataframe(preview, use_container_width=True, height=260)

    # --- 2. EAI -----------------------------------------------------------
    st.subheader("2. XML dos EAI de origem")
    st.write(
        "Carrega os XML da pasta `entrada_EAI` (podes selecionar vários) ou um **ZIP** com a pasta. "
        "São usados para copiar o horário real, comentários, SNMP, dependências e jobs ligados."
    )
    eai_uploads = st.file_uploader(
        "Ficheiros EAI", type=["xml", "zip"], accept_multiple_files=True, key="eai",
    )
    eai_sources, eai_errors = read_eai_uploads(eai_uploads)
    for err in eai_errors:
        st.warning(err)
    if eai_sources:
        st.caption(f"{len(eai_sources)} ficheiro(s) EAI carregado(s).")
    no_eai_ok = True
    if not eai_sources:
        no_eai_ok = st.checkbox(
            "Gerar sem EAI (todas as applications só com os valores fixos por defeito: sem "
            "schedule real, sem comentários, sem SNMP, sem dependências do EAI)",
            value=False,
        )

    # --- 3. Gerar ---------------------------------------------------------
    st.subheader("3. Gerar")
    problems = []
    if not infra_ok:
        problems.append("preencher a configuração do CAWA (em cima)")
    if not raw_rows:
        problems.append("carregar o Excel")
    if not no_eai_ok:
        problems.append("carregar os EAI ou confirmar que queres gerar sem eles")
    try:
        cols = {k: core.col_letter_to_index(v) for k, v in col_letters.items()}
    except ValueError as e:
        problems.append(f"corrigir o mapeamento de colunas ({e})")
        cols = None
    if problems:
        st.info("Falta: " + "; ".join(problems) + ".")

    if st.button("Gerar XMLs", type="primary", disabled=bool(problems)):
        with st.spinner("A gerar..."):
            result, log_lines, zip_bytes = run_generation(raw_rows, cols, int(start_row), eai_sources, infra)
        st.session_state["last"] = dict(result=result, log=log_lines, zip=zip_bytes)

    last = st.session_state.get("last")
    if last:
        r = last["result"]
        m = st.columns(4)
        m[0].metric("Applications geradas", r["generated"])
        m[1].metric("Conflitos (não gerados)", r["conflicts"])
        m[2].metric("EAI em falta", len(r["missing_eai"]))
        m[3].metric("Alertas/scripts a criar", len(r["manual_setup"]))
        st.download_button(
            "⬇️ Descarregar ZIP (XML + listas + log)", data=last["zip"],
            file_name="applications_MFT.zip", mime="application/zip", type="primary",
        )
        if r["conflicts"]:
            st.warning(f"{r['conflicts']} conflito(s) (Argumento2 diferente para a mesma Application) - NÃO gerados, ver log.")
        if r["missing_eai"]:
            st.warning(
                f"{r['eai_missing']} referência(s) a EAI não encontradas nos ficheiros carregados "
                f"({len(r['missing_eai'])} ficheiro(s) distinto(s) - ver '{core.MISSING_EAI_CSV}' no ZIP)."
            )
        if r["manual_setup"]:
            st.warning(
                f"{len(r['manual_setup'])} alerta(s)/script(s) para criar manualmente no CAWA - "
                f"ver '{core.MANUAL_SETUP_CSV}' no ZIP."
            )
        with st.expander(f"Log ({len(r['row_errors']) + len(r['warnings'])} aviso(s)/linha(s) ignorada(s))"):
            st.code("\n".join(last["log"]), language=None)


if __name__ == "__main__":
    main()
