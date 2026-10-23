"""
Versao WEB (stlite / WebAssembly) do CAWA XML Manager - lojas SALES.

Corre 100% no browser - os ficheiros carregados nunca saem do PC do utilizador.
Junta numa so pagina as duas ferramentas desktop:
- app_csv_upload.py      -> separador "Adicionar / remover lojas";
- set_only_on_request.py -> separador "Only on request".
A logica e a mesma (lojas_sales_core.py).
"""

import hashlib
import os
import xml.etree.ElementTree as ET

import streamlit as st

import lojas_sales_core as core


def _stem(filename):
    return os.path.splitext(os.path.basename(filename))[0]


def tab_lojas():
    st.write(
        "Carrega a application XML com os jobs `SALES_<loja>` / `nb_sales_files_fo_ok_<loja>`. "
        "As lojas novas são criadas a partir da primeira loja do XML (jobs e ligações `link`)."
    )
    uploaded_xml = st.file_uploader("Application XML (obrigatório)", type=["xml"], key="lojas_xml")
    uploaded_csv = st.file_uploader(
        "CSV de lojas (opcional)", type=["csv"], key="lojas_csv",
        help="Uma loja por linha na 1ª coluna (ex.: 00123, 123, loja_123). Pode ter cabeçalho.",
    )
    if uploaded_xml is None:
        st.info("Carrega um ficheiro XML para continuar.")
        return

    xml_bytes = uploaded_xml.getvalue()
    xml_hash = hashlib.sha1(xml_bytes).hexdigest()
    if st.session_state.get("xml_hash") != xml_hash:
        try:
            st.session_state.tree = core.parse_xml(xml_bytes)
        except ET.ParseError as exc:
            st.error(f"Erro ao analisar XML: {exc}")
            return
        st.session_state.xml_hash = xml_hash
        st.session_state.messages = []
    tree = st.session_state.tree

    for kind, msg in st.session_state.get("messages", []):
        getattr(st, kind)(msg)
    st.session_state.messages = []

    def done(messages):
        st.session_state.messages = messages
        st.rerun()

    store_ids = core.get_store_ids(tree)

    csv_store_ids = []
    if uploaded_csv is not None:
        try:
            csv_store_ids = core.parse_csv_stores(uploaded_csv.getvalue())
        except Exception as exc:
            st.error(f"Erro ao ler CSV: {exc}")

    col1, col2 = st.columns(2)
    with col1:
        st.subheader(f"Lojas existentes ({len(store_ids)})")
        st.write(", ".join(store_ids) or "Nenhuma loja encontrada")
        if uploaded_csv is not None:
            st.subheader(f"Lojas no CSV ({len(csv_store_ids)})")
            if csv_store_ids:
                st.write(", ".join(csv_store_ids))
            else:
                st.warning("Nenhuma loja válida encontrada no CSV.")

    with col2:
        st.subheader("Ações")
        new_store = st.text_input("Número de loja", value="")
        if st.button("Adicionar loja"):
            if not new_store.strip():
                st.warning("Indica um número de loja válido.")
            else:
                try:
                    core.add_store_jobs(tree, new_store.strip())
                    done([("success", f"Loja {new_store.strip()} adicionada com sucesso.")])
                except Exception as exc:
                    st.error(str(exc))

        if uploaded_csv is not None and st.button("Adicionar lojas do CSV", disabled=not csv_store_ids):
            messages = []
            for store_id in csv_store_ids:
                if store_id in store_ids:
                    messages.append(("info", f"Loja {store_id} já existe e foi ignorada."))
                    continue
                try:
                    core.add_store_jobs(tree, store_id)
                    store_ids.append(store_id)
                    messages.append(("success", f"Loja {store_id} adicionada com sucesso."))
                except Exception as exc:
                    messages.append(("error", f"Erro ao adicionar {store_id}: {exc}"))
            done(messages)

        remove_store = st.selectbox("Remover loja", [""] + store_ids)
        if st.button("Remover loja"):
            if not remove_store:
                st.warning("Seleciona uma loja para remover.")
            else:
                try:
                    core.remove_store_jobs(tree, remove_store)
                    done([("success", f"Loja {remove_store} removida com sucesso.")])
                except Exception as exc:
                    st.error(str(exc))

    st.markdown("---")
    updated_xml = core.xml_to_string(tree)
    st.download_button(
        "⬇️ Descarregar XML atualizado", data=updated_xml.encode("utf-8"),
        file_name=f"{_stem(uploaded_xml.name)}_atualizado.xml", mime="application/xml", type="primary",
    )
    with st.expander("Pré-visualização do XML"):
        st.code(updated_xml[:10000], language="xml")


def tab_only_on_request():
    st.write(
        "Põe os jobs `SALES_<loja>` e `nb_sales_files_fo_ok_<loja>` das lojas do CSV em "
        "**Only on request** (`<request>true</request>`). Gera um XML novo; o original não é alterado."
    )
    uploaded_xml = st.file_uploader("Application XML", type=["xml"], key="req_xml")
    uploaded_csv = st.file_uploader("CSV com a lista de lojas", type=["csv"], key="req_csv")
    column = st.text_input("Coluna das lojas no CSV", value="store")
    if uploaded_xml is None or uploaded_csv is None:
        st.info("Carrega o XML e o CSV para continuar.")
        return

    try:
        stores = core.read_stores_bytes(uploaded_csv.getvalue(), column)
        new_xml, changed, missing = core.set_only_on_request(uploaded_xml.getvalue(), stores)
    except (ValueError, UnicodeDecodeError, ET.ParseError) as exc:
        st.error(f"Erro: {exc}")
        return

    m = st.columns(3)
    m[0].metric("Lojas lidas do CSV", len(stores))
    m[1].metric("Jobs alterados", len(changed))
    m[2].metric("Lojas sem jobs no XML", len(missing))
    if missing:
        st.warning("Lojas sem jobs SALES/nb_sales no XML: " + ", ".join(f"{s:04d}" for s in missing))
    st.download_button(
        "⬇️ Descarregar XML", data=new_xml,
        file_name=f"{_stem(uploaded_xml.name)}_only_on_request.xml", mime="application/xml", type="primary",
    )
    if changed:
        with st.expander(f"Jobs alterados ({len(changed)})"):
            st.code("\n".join(changed), language=None)


def main():
    st.set_page_config(page_title="Lojas SALES (CAWA)", page_icon="🛒", layout="wide")
    st.title("🛒 Lojas SALES — CAWA XML Manager")
    st.caption("Corre no teu browser (nada é enviado para servidores).")
    t1, t2 = st.tabs(["Adicionar / remover lojas", "Only on request"])
    with t1:
        tab_lojas()
    with t2:
        tab_only_on_request()


if __name__ == "__main__":
    main()
