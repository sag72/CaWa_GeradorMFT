"""
Logica do CAWA XML Manager (lojas SALES / nb_sales_files_fo_ok), sem interface.

- Adicionar/remover lojas: funcoes de app_csv_upload.py (copiadas tal e qual).
- "Only on request": funcoes de set_only_on_request.py (copiadas tal e qual),
  com versoes que trabalham sobre bytes em vez de ficheiros (para o browser).
"""

from __future__ import annotations

import re
import csv
import xml.etree.ElementTree as ET
from copy import deepcopy
from io import BytesIO, StringIO

XML_NAMESPACE = "http://dto.wa.ca.com/application"
NS = {"appl": XML_NAMESPACE}


def parse_xml(xml_bytes):
    tree = ET.ElementTree(ET.fromstring(xml_bytes))
    return tree


def xml_to_string(tree):
    ET.register_namespace("", XML_NAMESPACE)
    bio = BytesIO()
    tree.write(bio, encoding="utf-8", xml_declaration=True)
    xml_string = bio.getvalue().decode("utf-8")
    xml_string = xml_string.replace(" />", "/>")
    if xml_string.startswith("<?xml"):
        xml_string = xml_string.replace("<?xml version='1.0' encoding='utf-8'?>", "<?xml version=\"1.0\" encoding=\"UTF-8\"?>")
    return xml_string


def zero_pad_store_id(store_id, width):
    return str(int(store_id)).zfill(width)


def canonical_store_id(value):
    if value is None:
        return None
    digits = re.search(r"^(0*)(\d+)$", value)
    if not digits:
        return value
    return str(int(digits.group(2)))


def normalize_csv_store_id(value):
    if not value:
        return None
    digits = re.search(r"(\d+)", value)
    if not digits:
        return None
    return canonical_store_id(digits.group(1))


def parse_csv_stores(csv_bytes):
    try:
        text = csv_bytes.decode("utf-8", errors="replace")
    except AttributeError:
        text = str(csv_bytes)
    if not text.strip():
        return []

    sample = "\n".join(text.splitlines()[:5])
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",;|\t")
    except csv.Error:
        dialect = csv.excel

    reader = csv.reader(StringIO(text), dialect)
    store_ids = []
    for row_index, row in enumerate(reader):
        if not row:
            continue
        first = row[0].strip()
        if not first:
            continue
        if row_index == 0 and re.search(r"^(store|loja|id|codigo|c[oó]digo)$", first, re.I) and not re.search(r"\d", first):
            continue
        normalized = normalize_csv_store_id(first)
        if normalized is not None:
            store_ids.append(normalized)
    return store_ids


def extract_store_id_from_job_name(name):
    if not name:
        return None
    match = re.search(r"nb_sales_files_fo_ok_(\d+)$", name)
    if match:
        return canonical_store_id(match.group(1))
    match = re.search(r"SALES_(\d+)$", name)
    if match:
        return canonical_store_id(match.group(1))
    return None


def get_store_ids(tree):
    root = tree.getroot()
    store_ids = set()
    for job in root.findall("appl:unix_job", NS):
        store_id = extract_store_id_from_job_name(job.get("name", ""))
        if store_id is not None:
            store_ids.add(store_id)
    return sorted(store_ids, key=lambda x: int(x))


def find_first_sample_job(tree, prefix):
    root = tree.getroot()
    for job in root.findall("appl:unix_job", NS):
        name = job.get("name", "")
        if name.startswith(prefix):
            return job
    return None


def update_relcount(link):
    dependencies = link.find("appl:dependencies", NS)
    if dependencies is None:
        return
    relcount = dependencies.find("appl:relcount", NS)
    relcond_list = dependencies.find("appl:relconditionlist", NS)
    if relcount is not None and relcond_list is not None:
        count = len(relcond_list.findall("appl:relcondition", NS))
        relcount.text = str(count)


def update_relconditions(tree, sample_name, new_name):
    root = tree.getroot()
    for link in root.findall("appl:link", NS):
        dependencies = link.find("appl:dependencies", NS)
        if dependencies is None:
            continue
        relcond_list = dependencies.find("appl:relconditionlist", NS)
        if relcond_list is None:
            continue
        for relcond in list(relcond_list.findall("appl:relcondition", NS)):
            succ = relcond.find("appl:successorname", NS)
            if succ is not None and succ.text == sample_name:
                new_rel = deepcopy(relcond)
                succ_new = new_rel.find("appl:successorname", NS)
                if succ_new is not None:
                    succ_new.text = new_name
                relcond_list.append(new_rel)
        update_relcount(link)


def remove_relconditions(tree, store_id):
    root = tree.getroot()
    for link in root.findall("appl:link", NS):
        dependencies = link.find("appl:dependencies", NS)
        if dependencies is None:
            continue
        relcond_list = dependencies.find("appl:relconditionlist", NS)
        if relcond_list is None:
            continue
        for relcond in list(relcond_list.findall("appl:relcondition", NS)):
            succ = relcond.find("appl:successorname", NS)
            if succ is not None:
                target_id = extract_store_id_from_job_name(succ.text)
                if target_id == canonical_store_id(store_id):
                    relcond_list.remove(relcond)
        update_relcount(link)


def format_store_name(sample_name, store_id):
    match = re.search(r"^(.*?)(\d+)$", sample_name)
    if not match:
        return f"{sample_name}_{store_id}"
    prefix, digits = match.groups()
    width = len(digits)
    return f"{prefix}{str(int(store_id)).zfill(width)}"


def replace_ids_in_element(element, replacements):
    if element is None:
        return
    for elem in element.iter():
        if elem.text:
            for old, new in replacements:
                elem.text = elem.text.replace(old, new)
        if elem.tail:
            for old, new in replacements:
                elem.tail = elem.tail.replace(old, new)
        for attr_name, attr_value in list(elem.attrib.items()):
            for old, new in replacements:
                attr_value = attr_value.replace(old, new)
            elem.attrib[attr_name] = attr_value


def replace_arg_store_id(element, target_id):
    if element is None or target_id is None:
        return
    pattern_store = re.compile(r"store:\d+")
    pattern_nb = re.compile(r"nb_sales_files_fo_ok\.sh\s+\d+")
    for elem in element.iter():
        if elem.text:
            elem.text = pattern_store.sub(f"store:{target_id}", elem.text)
            elem.text = pattern_nb.sub(f"nb_sales_files_fo_ok.sh {target_id}", elem.text)
        if elem.tail:
            elem.tail = pattern_store.sub(f"store:{target_id}", elem.tail)
            elem.tail = pattern_nb.sub(f"nb_sales_files_fo_ok.sh {target_id}", elem.tail)


def add_store_jobs(tree, store_id):
    normalized_new_id = canonical_store_id(store_id)
    if normalized_new_id is None:
        raise ValueError("Número de loja inválido.")

    existing_stores = get_store_ids(tree)
    if normalized_new_id in existing_stores:
        raise ValueError(f"Store {store_id} already exists in this application.")

    sample_nb = find_first_sample_job(tree, "nb_sales_files_fo_ok_")
    sample_sales = find_first_sample_job(tree, "SALES_")
    if sample_nb is None or sample_sales is None:
        raise ValueError("Could not find the sample store job definitions in the XML.")

    sample_nb_name = sample_nb.get("name")
    sample_sales_name = sample_sales.get("name")

    sample_nb_raw = re.search(r"nb_sales_files_fo_ok_(\d+)$", sample_nb_name)
    sample_sales_raw = re.search(r"SALES_(\d+)$", sample_sales_name)
    sample_nb_id = sample_nb_raw.group(1) if sample_nb_raw else None
    sample_sales_id = sample_sales_raw.group(1) if sample_sales_raw else None

    target_nb_id = zero_pad_store_id(normalized_new_id, len(sample_nb_id)) if sample_nb_id else normalized_new_id
    target_sales_id = zero_pad_store_id(normalized_new_id, len(sample_sales_id)) if sample_sales_id else normalized_new_id

    target_nb_name = format_store_name(sample_nb_name, target_nb_id)
    target_sales_name = format_store_name(sample_sales_name, target_sales_id)

    new_nb = deepcopy(sample_nb)
    new_sales = deepcopy(sample_sales)

    canonical_target_id = canonical_store_id(normalized_new_id)

    new_nb.set("name", target_nb_name)
    nb_replacements = [(f"nb_sales_files_fo_ok_{sample_nb_id}", f"nb_sales_files_fo_ok_{target_nb_id}")]
    if sample_nb_id and target_nb_id:
        replace_ids_in_element(new_nb, nb_replacements)
    if sample_sales_id and target_sales_id:
        sales_replacements = [
            (f"SALES_{sample_sales_id}", f"SALES_{target_sales_id}"),
        ]
        replace_ids_in_element(new_nb, sales_replacements)
    replace_arg_store_id(new_nb, canonical_target_id)

    new_sales.set("name", target_sales_name)
    sales_replacements = [
        (f"SALES_{sample_sales_id}", f"SALES_{target_sales_id}"),
    ]
    if sample_nb_id and target_nb_id:
        nb_replacements = [(f"nb_sales_files_fo_ok_{sample_nb_id}", f"nb_sales_files_fo_ok_{target_nb_id}")]
        replace_ids_in_element(new_sales, nb_replacements)
    replace_ids_in_element(new_sales, sales_replacements)
    replace_arg_store_id(new_sales, canonical_target_id)

    root = tree.getroot()
    root.append(new_nb)
    root.append(new_sales)

    update_relconditions(tree, f"{sample_nb_name}.0001", f"{target_nb_name}.0001")
    update_relconditions(tree, f"{sample_sales_name}.0001", f"{target_sales_name}.0001")


def remove_store_jobs(tree, store_id):
    normalized_id = canonical_store_id(store_id)
    root = tree.getroot()
    removed = False
    for job in list(root.findall("appl:unix_job", NS)):
        job_id = extract_store_id_from_job_name(job.get("name", ""))
        if job_id == normalized_id:
            root.remove(job)
            removed = True
    if not removed:
        raise ValueError(f"Store {store_id} not found in XML jobs.")
    remove_relconditions(tree, store_id)


# --------------------------------------------------------------------------
# "Only on request" (de set_only_on_request.py)
# --------------------------------------------------------------------------

JOB_PATTERNS = (
    re.compile(r"^SALES_(\d+)$"),
    re.compile(r"^nb_sales_files_fo_ok_(\d+)$"),
)


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def normalise_store(value: str) -> int:
    value = value.strip()
    if not value.isdecimal():
        raise ValueError(f"Loja inválida: {value!r}. A loja deve conter apenas dígitos.")
    return int(value)


def job_store(job_name: str) -> int | None:
    for pattern in JOB_PATTERNS:
        match = pattern.fullmatch(job_name)
        if match:
            return int(match.group(1))
    return None


def request_element(job: ET.Element) -> ET.Element:
    for child in job:
        if local_name(child.tag) == "request":
            return child

    namespace = job.tag[: job.tag.find("}") + 1] if job.tag.startswith("{") else ""
    element = ET.Element(f"{namespace}request")
    job.append(element)
    return element


# --------------------------------------------------------------------------
# "Only on request" sobre bytes (equivalente ao main() de set_only_on_request.py)
# --------------------------------------------------------------------------

def read_stores_bytes(csv_bytes: bytes, column: str = "store") -> set[int]:
    """Igual a read_stores() de set_only_on_request.py, mas a partir de bytes."""
    text = csv_bytes.decode("utf-8-sig")
    reader = csv.DictReader(StringIO(text, newline=""))
    if not reader.fieldnames:
        raise ValueError("O CSV não tem cabeçalho.")

    columns = {name.casefold(): name for name in reader.fieldnames}
    actual_column = columns.get(column.casefold())
    if actual_column is None:
        available = ", ".join(reader.fieldnames)
        raise ValueError(
            f"Coluna {column!r} não encontrada. Colunas disponíveis: {available}"
        )

    stores = {
        normalise_store(row[actual_column])
        for row in reader
        if row.get(actual_column, "").strip()
    }

    if not stores:
        raise ValueError("O CSV não contém lojas.")
    return stores


def set_only_on_request(xml_bytes: bytes, stores: set[int]):
    """Marca <request>true</request> nos jobs SALES_/nb_sales_files_fo_ok_ das lojas.

    Devolve (xml_bytes_novo, jobs_alterados, lojas_sem_jobs)."""
    tree = ET.ElementTree(ET.fromstring(xml_bytes))
    changed: list[str] = []
    found_stores: set[int] = set()
    for element in tree.getroot().iter():
        if local_name(element.tag) not in {"unix_job", "ext_job"}:
            continue

        name = element.get("name", "")
        store = job_store(name)
        if store is None or store not in stores:
            continue

        found_stores.add(store)
        request = request_element(element)
        if (request.text or "").strip().lower() != "true":
            request.text = "true"
            changed.append(name)

    missing = sorted(stores - found_stores)
    ET.indent(tree, space="    ")
    # namespace por omissao (<appl xmlns=...>, formato nativo do CAWA) em vez de ns0:,
    # para o resultado nao depender de xml_to_string() ja ter sido chamado antes
    ET.register_namespace("", XML_NAMESPACE)
    out = BytesIO()
    tree.write(out, encoding="utf-8", xml_declaration=True)
    # o script desktop escreve em modo texto no Windows (\r\n); igual aqui
    return out.getvalue().replace(b"\n", b"\r\n"), changed, missing
