"""
Logica do Gerador de Applications MFT (CA Workload Automation), sem interface.

Usada pela app desktop (app.py, Tkinter) e pela versao web (web/app_web.py,
stlite). Nao tem acesso ao disco: recebe as linhas do Excel/CSV e uma funcao
que devolve o conteudo de cada EAI, e devolve os ficheiros gerados em memoria.

Os valores de infraestrutura do CAWA (agente, agente SNMP, user id e comando)
nao estao aqui - sao passados em `infra` (ver INFRA_FIELDS), para este ficheiro
poder ser publicado sem dados internos.
"""

import csv
import io
import os
import re
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape

try:
    import openpyxl
    from openpyxl.utils import column_index_from_string
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

NS_URI = "http://dto.wa.ca.com/application"
NS = "{%s}" % NS_URI
ET.register_namespace("", NS_URI)

APPL_LEVEL_TAGS = {"appl_runtime_name", "version", "comment", "applopts", "defaults", "job_templates"}

# Valores de infraestrutura que entram nos templates (todos obrigatorios).
INFRA_FIELDS = ["agent", "snmp_agent", "userid", "cmdname"]

APPL_HEADER_TEMPLATE = (
    '<?xml version="1.0" encoding="UTF-8"?>\n'
    '<appl xmlns="http://dto.wa.ca.com/application" name="{app_name}">'
    '<appl_runtime_name>{app_name}_%APPL._SDD%APPL._SMM</appl_runtime_name>'
    '<version>12.0</version>'
    '<comment>{comment}</comment>'
    '<applopts><wait>false</wait><estimate_endtime>false</estimate_endtime>'
    '<prop_dueout>false</prop_dueout><hold>false</hold><noinherit>false</noinherit>'
    '<notrigger_ifactive>false</notrigger_ifactive>'
    '<suppress_nowork_notification>false</suppress_nowork_notification>'
    '<reason_required>false</reason_required></applopts>'
    '<defaults>{schedules_xml}<tag/>'
    '{snmp_notifylist_xml}'
    '<rununit><agent>{agent}</agent></rununit></defaults>'
    '<job_templates><app><appl name="Template"><version>11.1</version></appl>'
    '</app></job_templates>'
)

DEFAULT_SNMP_NOTIFYLIST_TEMPLATE = (
    '<snmp_notifylist><snmp_notify><returncode/><monitor_states>'
    '<monitor_state>OVERDUE</monitor_state><monitor_state>FAILED</monitor_state>'
    '<monitor_state>PREMEND</monitor_state><monitor_state>SUBERROR</monitor_state>'
    '<monitor_state>INACTIVE</monitor_state><monitor_state>UNKNOWN</monitor_state>'
    '<monitor_state>ABANDONSUBMISSION</monitor_state></monitor_states>'
    '<snmpagent>{snmp_agent}</snmpagent>'
    '<message>CAWA Application: %APPL._name Geracao: %APPL._gen Job_Name: '
    '%WOB._fullname Status: %WOB._state Completion_Code: %WOB._cmpc Agente: '
    '%WOB._agent</message></snmp_notify>'
    '<snmp_notify><returncode/><monitor_states>'
    '<monitor_state>AGENTDOWN</monitor_state></monitor_states>'
    '<snmpagent>{snmp_agent}</snmpagent>'
    '<message>CAWA Agente: %WOB._agent Down</message></snmp_notify>'
    '</snmp_notifylist>'
)

DEFAULT_DEPENDENCIES_XML = "<dependencies><relcount>0</relcount></dependencies>"
DEFAULT_ALERT_XML = "<alert_notifynodefaults>false</alert_notifynodefaults>"
DEFAULT_EXITCODELIST_XML = ""
DEFAULT_MAX_MIN_XML = ""

MAIN_JOB_TEMPLATE = (
    '<unix_job name="{job_name}" qualifier="0001">'
    '<comment>{comment}</comment>'
    '{dependencies_xml}'
    '<notifynodefaults>false</notifynodefaults>'
    '{alert_xml}'
    '<snmp_notifynodefaults>false</snmp_notifynodefaults>'
    '<hold>false</hold><request>false</request><conditional>false</conditional>'
    '<criticaljob>false</criticaljob>'
    '<job_ancestor_wait_default_ignore>true</job_ancestor_wait_default_ignore>'
    '{max_min_xml}'
    '<retry><retry_count>0</retry_count><retry_interval>0</retry_interval>'
    '<last_notify_email>false</last_notify_email>'
    '<last_notify_alerts>false</last_notify_alerts>'
    '<last_notify_snmp>false</last_notify_snmp></retry>'
    '{exitcodelist_xml}'
    '<subappl_wait>false</subappl_wait>'
    '<userid>{userid}</userid>'
    '<args>{args}</args>'
    '<cmdname>{cmdname}</cmdname>'
    '<isscript>true</isscript><shell>/bin/ksh</shell>'
    '<processpriority>NORMAL</processpriority>'
    '</unix_job>'
)

DEFAULT_SCHEDULES_XML = "<schedules><run><schedule>DAILY</schedule></run></schedules>"

MANUAL_SETUP_CSV = "alertas_scripts_criar_manualmente.csv"
MISSING_EAI_CSV = "eai_em_falta_na_pasta_entrada.csv"


# --------------------------------------------------------------------------
# Excel / CSV
# --------------------------------------------------------------------------

def load_raw_rows(source, sheet_name=None, filename=None):
    """Devolve uma lista de listas (linhas x colunas), tudo como string.

    `source` e um caminho ou um ficheiro em memoria (bytes/file-like); neste caso
    `filename` indica a extensao."""
    ext = os.path.splitext(filename or source)[1].lower()
    if isinstance(source, (bytes, bytearray)):
        source = io.BytesIO(source)
    if ext in (".xlsx", ".xlsm"):
        if not HAS_OPENPYXL:
            raise RuntimeError("openpyxl nao esta instalado (pip install openpyxl)")
        wb = openpyxl.load_workbook(source, data_only=True, read_only=True)
        ws = wb[sheet_name] if sheet_name else wb.active
        rows = []
        for row in ws.iter_rows(values_only=True):
            rows.append(["" if v is None else str(v) for v in row])
        return rows
    elif ext == ".csv":
        if isinstance(source, str):
            with open(source, newline="", encoding="utf-8-sig") as f:
                return [row for row in csv.reader(f)]
        text = source.read().decode("utf-8-sig")
        return [row for row in csv.reader(io.StringIO(text, newline=""))]
    else:
        raise RuntimeError(f"Extensao nao suportada: {ext}")


def list_sheet_names(source):
    """Nomes das folhas de um Excel (caminho ou bytes)."""
    if isinstance(source, (bytes, bytearray)):
        source = io.BytesIO(source)
    wb = openpyxl.load_workbook(source, read_only=True)
    return wb.sheetnames


def col_letter_to_index(letter):
    letter = letter.strip().upper()
    if not letter:
        raise ValueError("Coluna vazia")
    if HAS_OPENPYXL:
        return column_index_from_string(letter) - 1
    idx = 0
    for ch in letter:
        if not ("A" <= ch <= "Z"):
            raise ValueError(f"Letra de coluna invalida: {letter}")
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx - 1


def index_to_letter(idx):
    letters = ""
    idx += 1
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def derive_job_name(argumento2):
    m = re.search(r'"([^"]+)"', argumento2)
    if not m:
        return None
    parts = m.group(1).split(".")
    return parts[-1].strip().upper() if parts else None


def get_cycl_suffix(as_is_name):
    """Extrai o sufixo _CYCLxxxx do nome AS-IS (EAI), se existir, para ser acrescentado
    ao nome da application MFT (o Excel nao o inclui, mas deve manter-se, como no EAI)."""
    m = re.search(r"(_CYCL\w*)$", as_is_name or "", re.IGNORECASE)
    return m.group(1).upper() if m else ""


def apply_cycl_suffix(app_name, cycl_suffix):
    if not cycl_suffix or app_name.upper().endswith(cycl_suffix):
        return app_name
    return app_name + cycl_suffix


# --------------------------------------------------------------------------
# Leitura/estrutura do EAI de origem
# --------------------------------------------------------------------------

def local_tag(elem):
    t = elem.tag
    return t.split("}", 1)[1] if "}" in t else t


def get_job_elements(root):
    """Todos os elementos de job (unix_job, link, ext_job, ...), pela ordem do ficheiro."""
    return [c for c in root if local_tag(c) not in APPL_LEVEL_TAGS]


def is_soacnt_job(elem):
    if local_tag(elem) != "unix_job":
        return False
    args_elem = elem.find(f"{NS}args")
    return bool(args_elem is not None and args_elem.text and "run_soacnt.sh" in args_elem.text)


def get_job_id(elem):
    args_elem = elem.find(f"{NS}args")
    if args_elem is None or not args_elem.text:
        return None
    m = re.search(r"run_soacnt\.sh\s+(\d+)", args_elem.text)
    return m.group(1) if m else None


def find_matched_job(root, job_name, job_id):
    """Devolve (job_elem_ou_None, aviso_ou_None). Casa primeiro pelo nome exato do job
    (coluna 'nome' do Excel - mais fiavel, cobre tambem nomes com formato App@Job), com o
    job_ID (extraido do args) como alternativa."""
    soacnt_jobs = [j for j in get_job_elements(root) if is_soacnt_job(j)]
    if not soacnt_jobs:
        return None, "o EAI de origem nao tem nenhum job run_soacnt.sh reconhecivel"

    job_name = (job_name or "").strip()
    job_id = (job_id or "").strip()

    if job_name:
        for j in soacnt_jobs:
            if j.get("name", "") == job_name:
                return j, None

    if job_id:
        for j in soacnt_jobs:
            if get_job_id(j) == job_id:
                note = (
                    f"nome do job '{job_name}' nao bateu certo com nenhum job do EAI, usei o job_ID '{job_id}'"
                    if job_name
                    else None
                )
                return j, note

    if len(soacnt_jobs) == 1:
        note = None
        if job_name or job_id:
            note = (
                f"nome do job '{job_name}' e job_ID '{job_id}' nao bateram certo com nenhum job do EAI, "
                f"mas so ha 1 job run_soacnt.sh no ficheiro, a usar esse"
            )
        return soacnt_jobs[0], note

    return None, (
        f"o EAI de origem tem {len(soacnt_jobs)} jobs run_soacnt.sh e nenhum bate com o nome '{job_name}' "
        f"nem com o job_ID '{job_id}'"
    )


def _strip_ws(elem):
    if elem.text and not elem.text.strip():
        elem.text = None
    elif elem.text:
        elem.text = elem.text.strip()
    if elem.tail and not elem.tail.strip():
        elem.tail = None
    elif elem.tail:
        elem.tail = elem.tail.strip()
    for child in elem:
        _strip_ws(child)


def render_passthrough_job(elem, old_ref, new_ref, old_name, new_name):
    """Serializa um elemento copiado do EAI. O ElementTree redeclara sempre o namespace
    default no elemento de topo de qualquer subarvore serializada isoladamente; como o
    <appl> ja o declara na raiz, essa redeclaracao e removida para o XML final ficar
    identico ao formato de referencia (so 1 xmlns, na raiz)."""
    _strip_ws(elem)
    text = ET.tostring(elem, encoding="unicode")
    text = text.replace(f' xmlns="{NS_URI}"', "")
    text = re.sub(r">\s+<", "><", text).strip()
    text = text.replace(old_ref, new_ref)
    text = text.replace(old_name, new_name)
    return text


def compute_rename(matched_job, job_name):
    """(old_ref, new_ref, old_name, new_name) para atualizar referencias (successorname)
    do nome antigo do job (EAI) para o novo nome do job (MFT)."""
    old_name = matched_job.get("name", "")
    old_qualifier = matched_job.get("qualifier", "0001")
    return f"{old_name}.{old_qualifier}", f"{job_name}.0001", old_name, job_name


def get_dependencies_xml(matched_job, rename):
    """Copia o <dependencies> do job de origem (preserva relconditionlist/successorname,
    ou seja, as ligacoes de saida deste job para outros jobs), tal e qual."""
    dep = matched_job.find(f"{NS}dependencies")
    if dep is None:
        return DEFAULT_DEPENDENCIES_XML
    return render_passthrough_job(dep, *rename)


def get_exitcodelist_xml(matched_job, rename):
    """Copia <exitcodelist> do job de origem tal e qual (define codigos de saida
    especiais, ex: exit 212 tratado como SUCCESS, usados nas condicoes de fluxo)."""
    elem = matched_job.find(f"{NS}exitcodelist")
    if elem is None:
        return DEFAULT_EXITCODELIST_XML
    return render_passthrough_job(elem, *rename)


def get_max_min_xml(matched_job, rename):
    """Copia <max_min> do job de origem tal e qual (dependencia de tempo maximo/minimo
    de execucao do job)."""
    elem = matched_job.find(f"{NS}max_min")
    if elem is None:
        return DEFAULT_MAX_MIN_XML
    return render_passthrough_job(elem, *rename)


def get_appl_comment(root):
    """Comentario da application (nivel appl) do EAI de origem."""
    elem = root.find(f"{NS}comment")
    return (elem.text or "").strip() if elem is not None else ""


def get_job_comment(matched_job):
    """Comentario do job (nivel unix_job) do EAI de origem."""
    elem = matched_job.find(f"{NS}comment")
    return (elem.text or "").strip() if elem is not None else ""


def get_alert_xml(matched_job, rename):
    """Copia <alert_notifynodefaults>/<alert_notifylist> do job de origem tal e qual (o
    alerta que dispara, por exemplo, um job link associado)."""
    notif = matched_job.find(f"{NS}alert_notifynodefaults")
    notif_val = (notif.text or "false").strip() if notif is not None else "false"
    parts = [f"<alert_notifynodefaults>{notif_val}</alert_notifynodefaults>"]
    alert_list = matched_job.find(f"{NS}alert_notifylist")
    if alert_list is not None:
        parts.append(render_passthrough_job(alert_list, *rename))
    return "".join(parts)


def get_passthrough_jobs(root, matched_job, rename):
    """Todos os job elements do EAI exceto o job migrado e outros jobs run_soacnt.sh irmaos
    (que pertencem a outras applications MFT), com o successorname atualizado para o novo
    nome do job MFT."""
    extras = []
    for elem in get_job_elements(root):
        if elem is matched_job or is_soacnt_job(elem):
            continue
        extras.append(render_passthrough_job(elem, *rename))
    return extras


def build_cross_app_map(root, as_is_name, matched_job, job_id_to_target):
    """Para os outros jobs run_soacnt.sh do mesmo EAI (irmaos, que migram para outras
    applications MFT), devolve um dict {old_ref: new_ref} usando o formato
    NomeApp@NomeJob.0001 para reescrever referencias cruzadas (successorname), e a lista
    de old_refs que nao foi possivel resolver (esse job irmao nao foi migrado)."""
    mapping = {}
    unresolved = []
    for elem in get_job_elements(root):
        if elem is matched_job or not is_soacnt_job(elem):
            continue
        old_ref = f"{elem.get('name', '')}.{elem.get('qualifier', '0001')}"
        target = job_id_to_target.get((as_is_name, get_job_id(elem)))
        if target:
            target_app, target_job = target
            mapping[old_ref] = f"{target_app}@{target_job}.0001"
        else:
            unresolved.append(old_ref)
    return mapping, unresolved


def apply_cross_refs(text, mapping):
    for old_ref, new_ref in mapping.items():
        text = text.replace(old_ref, new_ref)
    return text


def find_alerts_and_scripts(xml_content):
    """Alertas (<alert>) e scripts (<script_reference>) usados na application gerada -
    objetos que nao vem no import e precisam de ser criados manualmente no CAWA."""
    alerts = sorted(set(re.findall(r"<alert>([^<]+)</alert>", xml_content)))
    scripts = sorted(set(re.findall(r"<script_reference>([^<]+)</script_reference>", xml_content)))
    return alerts, scripts


def extract_run_blocks(content):
    """Devolve a lista de blocos <run>...</run> encontrados dentro de <schedules>
    no XML EAI (texto bruto), com <schedule> normalizado para maiusculas."""
    m = re.search(r"<schedules>(.*?)</schedules>", content, re.DOTALL)
    if not m:
        return []
    inner = re.sub(r">\s+<", "><", m.group(1)).strip()
    inner = re.sub(
        r"(<schedule>)(.*?)(</schedule>)",
        lambda mm: mm.group(1) + mm.group(2).strip().upper() + mm.group(3),
        inner,
    )
    return re.findall(r"<run>.*?</run>", inner, re.DOTALL)


def extract_snmp_notifylist(content):
    """Extrai o bloco <snmp_notifylist>...</snmp_notifylist> tal como esta no EAI de
    origem (agente SNMP e monitor_states podem variar de application para application)."""
    m = re.search(r"<snmp_notifylist>.*?</snmp_notifylist>", content, re.DOTALL)
    if not m:
        return None
    return re.sub(r">\s+<", "><", m.group(0)).strip()


def build_schedules_xml(run_blocks):
    unique = []
    seen = set()
    for r in run_blocks:
        if r not in seen:
            seen.add(r)
            unique.append(r)
    if not unique:
        return DEFAULT_SCHEDULES_XML
    return "<schedules>" + "".join(unique) + "</schedules>"


# --------------------------------------------------------------------------
# Construcao do XML final
# --------------------------------------------------------------------------

def build_appl_xml(
    app_name,
    job_name,
    argumento2,
    schedules_xml,
    extras_xml_list,
    infra,
    appl_comment="",
    job_comment="",
    dependencies_xml=DEFAULT_DEPENDENCIES_XML,
    alert_xml=DEFAULT_ALERT_XML,
    exitcodelist_xml=DEFAULT_EXITCODELIST_XML,
    snmp_notifylist_xml=None,
    max_min_xml=DEFAULT_MAX_MIN_XML,
):
    if snmp_notifylist_xml is None:
        snmp_notifylist_xml = DEFAULT_SNMP_NOTIFYLIST_TEMPLATE.format(snmp_agent=escape(infra["snmp_agent"]))
    header = APPL_HEADER_TEMPLATE.format(
        app_name=escape(app_name),
        comment=escape(appl_comment),
        schedules_xml=schedules_xml,
        snmp_notifylist_xml=snmp_notifylist_xml,
        agent=escape(infra["agent"]),
    )
    main_job = MAIN_JOB_TEMPLATE.format(
        job_name=escape(job_name),
        comment=escape(job_comment),
        args=escape(argumento2),
        dependencies_xml=dependencies_xml,
        alert_xml=alert_xml,
        exitcodelist_xml=exitcodelist_xml,
        max_min_xml=max_min_xml,
        userid=escape(infra["userid"]),
        cmdname=escape(infra["cmdname"]),
    )
    return header + main_job + "".join(extras_xml_list) + "</appl>"


class EaiSourceCache:
    """Evita reler/reparsear o mesmo ficheiro EAI varias vezes na mesma geracao.

    `loader(as_is_name)` devolve (identificador, conteudo_texto) ou None se o EAI nao
    existir (o identificador - caminho ou nome do ficheiro - so serve para mensagens)."""

    def __init__(self, loader):
        self.loader = loader
        self._cache = {}

    def get(self, as_is_name):
        if as_is_name in self._cache:
            return self._cache[as_is_name]
        found = self.loader(as_is_name) if as_is_name else None
        if not found:
            result = (None, None, None)
        else:
            path, content = found
            try:
                root = ET.fromstring(content.encode("utf-8"))
            except ET.ParseError as e:
                result = (path, content, None)
                self._cache[as_is_name] = result
                raise ValueError(f"XML invalido em '{path}': {e}")
            result = (path, content, root)
        self._cache[as_is_name] = result
        return result


def dir_loader(eai_dir):
    """Loader de EAI a partir de uma pasta: <eai_dir>/<AS-IS>.xml."""
    def load(as_is_name):
        if not eai_dir or not as_is_name:
            return None
        candidate = os.path.join(eai_dir, f"{as_is_name}.xml")
        if not os.path.isfile(candidate):
            return None
        with open(candidate, encoding="utf-8") as f:
            return candidate, f.read()
    return load


def _csv_text(header, rows):
    buf = io.StringIO(newline="")
    writer = csv.writer(buf, delimiter=";")
    writer.writerow(header)
    for row in rows:
        writer.writerow(row)
    return "﻿" + buf.getvalue()


# --------------------------------------------------------------------------
# Geracao
# --------------------------------------------------------------------------

def generate(raw_rows, cols, start_row, eai_loader, eai_label, infra, log):
    """Gera as applications MFT.

    cols: dict com os indices (0-based) as_is, job_name, app, args, job_id.
    eai_loader: ver EaiSourceCache; None = gerar so com valores por defeito.
    eai_label: descricao da origem dos EAI (pasta/ficheiros) para as mensagens.
    infra: dict com INFRA_FIELDS.
    log: funcao chamada com cada linha de log.

    Devolve dict com files ({nome: conteudo}, pela ordem de geracao), generated,
    conflicts, eai_missing, missing_eai, manual_setup, row_errors e warnings.
    Os nomes dos CSV auxiliares sao MANUAL_SETUP_CSV e MISSING_EAI_CSV."""
    missing_infra = [f for f in INFRA_FIELDS if not (infra.get(f) or "").strip()]
    if missing_infra:
        raise ValueError(f"Valores de infraestrutura em falta: {', '.join(missing_infra)}")

    cache = EaiSourceCache(eai_loader) if eai_loader else None

    # --- 1a passagem: recolher e validar linhas ---------------------
    entries = []
    row_errors = []

    def cell(row, idx):
        return (row[idx] if idx < len(row) else "").strip()

    for i, row in enumerate(raw_rows[start_row - 1:], start=start_row):
        as_is_name = cell(row, cols["as_is"])
        job_name_as_is = cell(row, cols["job_name"])
        app_name = cell(row, cols["app"])
        argumento2 = cell(row, cols["args"])
        job_id = cell(row, cols["job_id"])

        if not app_name and not argumento2:
            continue
        if not app_name:
            row_errors.append(f"Linha {i}: Application (TO-BE) em branco, ignorada.")
            continue
        if not app_name.upper().startswith("MFT_"):
            row_errors.append(
                f"Linha {i}: Application (TO-BE) '{app_name}' nao comeca por 'MFT_' "
                f"(linha provavelmente incompleta/placeholder), ignorada."
            )
            continue
        if not argumento2:
            row_errors.append(f"Linha {i}: Argumento2 em branco para '{app_name}', ignorada.")
            continue

        entries.append(dict(
            row=i, as_is=as_is_name, job_name_as_is=job_name_as_is,
            app_name=app_name, argumento2=argumento2, job_id=job_id,
        ))

    # --- agrupar por Application (TO-BE) ------------------------------
    groups = {}
    order = []
    for e in entries:
        if e["app_name"] not in groups:
            groups[e["app_name"]] = []
            order.append(e["app_name"])
        groups[e["app_name"]].append(e)

    files = {}
    generated = 0
    conflicts = 0
    eai_missing = 0
    warnings = []
    manual_setup = []
    missing_eai = {}

    # --- 2a passagem: resolver nome final/job de cada grupo (sem tocar no EAI) -----
    resolved = {}
    for app_name in order:
        group = groups[app_name]
        distinct_args = sorted({g["argumento2"] for g in group})

        if len(distinct_args) > 1:
            conflicts += 1
            detail = "; ".join(f"linha {g['row']} (AS-IS={g['as_is']}, args={g['argumento2']!r})" for g in group)
            warnings.append(
                f"CONFLITO '{app_name}': {len(group)} linhas com Argumento2 diferente, "
                f"NAO gerado - resolver no Excel primeiro. Detalhe: {detail}"
            )
            continue

        argumento2 = distinct_args[0]
        job_name = derive_job_name(argumento2)
        if not job_name:
            warnings.append(
                f"Linha(s) {[g['row'] for g in group]} ({app_name}): nao foi possivel derivar o "
                f"nome do job a partir do Argumento2 '{argumento2}', ignorado."
            )
            continue

        cycl_suffix = ""
        for g in group:
            cycl_suffix = get_cycl_suffix(g["as_is"])
            if cycl_suffix:
                break
        final_app_name = apply_cycl_suffix(app_name, cycl_suffix)
        resolved[app_name] = dict(job_name=job_name, argumento2=argumento2, final_app_name=final_app_name)

    # --- indice global job_id -> (application MFT, job MFT), para resolver ---------
    # referencias cruzadas entre applications (jobs do mesmo EAI que migram para
    # applications MFT diferentes).
    job_id_to_target = {}
    for app_name, info in resolved.items():
        for g in groups[app_name]:
            if g["as_is"] and g["job_id"]:
                job_id_to_target[(g["as_is"], g["job_id"])] = (info["final_app_name"], info["job_name"])

    # --- 3a passagem: gerar cada application, usando o EAI de origem ---------------
    for app_name in order:
        if app_name not in resolved:
            continue
        group = groups[app_name]
        job_name = resolved[app_name]["job_name"]
        argumento2 = resolved[app_name]["argumento2"]
        final_app_name = resolved[app_name]["final_app_name"]

        run_blocks = []
        extras = []
        seen_extras = set()
        dependencies_xml = DEFAULT_DEPENDENCIES_XML
        alert_xml = DEFAULT_ALERT_XML
        exitcodelist_xml = DEFAULT_EXITCODELIST_XML
        max_min_xml = DEFAULT_MAX_MIN_XML
        snmp_notifylist_xml = None
        appl_comment = ""
        job_comment = ""
        dependencies_set = False

        if cache is not None:
            for g in group:
                if not g["as_is"]:
                    warnings.append(f"Linha {g['row']} ({app_name}): sem Application (AS-IS), sem EAI de origem.")
                    continue
                try:
                    eai_path, content, root = cache.get(g["as_is"])
                except ValueError as e:
                    warnings.append(f"Linha {g['row']} ({app_name}): {e}")
                    continue
                if not eai_path:
                    eai_missing += 1
                    missing_eai.setdefault(g["as_is"], set()).add(final_app_name)
                    warnings.append(
                        f"Linha {g['row']} ({app_name}): nao encontrei '{g['as_is']}.xml' em '{eai_label}'."
                    )
                    continue

                run_blocks.extend(extract_run_blocks(content))

                if snmp_notifylist_xml is None:
                    snmp_notifylist_xml = extract_snmp_notifylist(content)
                if not appl_comment:
                    appl_comment = get_appl_comment(root)

                matched_job, note = find_matched_job(root, g["job_name_as_is"], g["job_id"])
                if note:
                    warnings.append(f"Linha {g['row']} ({app_name}): {note}")
                if matched_job is None:
                    continue

                rename = compute_rename(matched_job, job_name)
                cross_map, unresolved = build_cross_app_map(root, g["as_is"], matched_job, job_id_to_target)

                if not dependencies_set:
                    dependencies_xml = apply_cross_refs(get_dependencies_xml(matched_job, rename), cross_map)
                    alert_xml = apply_cross_refs(get_alert_xml(matched_job, rename), cross_map)
                    exitcodelist_xml = apply_cross_refs(get_exitcodelist_xml(matched_job, rename), cross_map)
                    max_min_xml = apply_cross_refs(get_max_min_xml(matched_job, rename), cross_map)
                    job_comment = get_job_comment(matched_job)
                    dependencies_set = True
                for extra in get_passthrough_jobs(root, matched_job, rename):
                    extra = apply_cross_refs(extra, cross_map)
                    if extra not in seen_extras:
                        seen_extras.add(extra)
                        extras.append(extra)

                combined = dependencies_xml + alert_xml + exitcodelist_xml + "".join(extras)
                for old_ref in unresolved:
                    if old_ref in combined:
                        warnings.append(
                            f"Linha {g['row']} ({app_name}): referencia cruzada para '{old_ref}' nao foi "
                            f"possivel resolver (esse job nao foi migrado para nenhuma application MFT); "
                            f"referencia mantida como no EAI original."
                        )

        schedules_xml = build_schedules_xml(run_blocks)

        xml_content = build_appl_xml(
            final_app_name,
            job_name,
            argumento2,
            schedules_xml,
            extras,
            infra,
            appl_comment=appl_comment,
            job_comment=job_comment,
            dependencies_xml=dependencies_xml,
            alert_xml=alert_xml,
            exitcodelist_xml=exitcodelist_xml,
            snmp_notifylist_xml=snmp_notifylist_xml,
            max_min_xml=max_min_xml,
        )
        files[f"{final_app_name}.xml"] = xml_content
        generated += 1

        alerts, scripts = find_alerts_and_scripts(xml_content)
        for alert in alerts:
            manual_setup.append((final_app_name, "Alerta", alert))
        for script in scripts:
            manual_setup.append((final_app_name, "Script", script))

        rows_str = ",".join(str(g["row"]) for g in group)
        extra_note = ""
        if alerts or scripts:
            extra_note = f", alertas={alerts}, scripts={scripts}"
        log(
            f"{final_app_name}.xml gerado (linha(s) {rows_str}, job={job_name}, "
            f"schedules={schedules_xml}, jobs extra={len(extras)}{extra_note})."
        )

    if manual_setup:
        files[MANUAL_SETUP_CSV] = _csv_text(
            ["Application MFT", "Tipo", "Nome", "Criado (Sim/Nao)"],
            [[*row, ""] for row in manual_setup],
        )
    if missing_eai:
        files[MISSING_EAI_CSV] = _csv_text(
            ["Ficheiro EAI em falta", "Applications MFT afetadas", "Colocado (Sim/Nao)"],
            [[f"{name}.xml", ", ".join(sorted(missing_eai[name])), ""] for name in sorted(missing_eai)],
        )

    return dict(
        files=files, generated=generated, conflicts=conflicts, eai_missing=eai_missing,
        missing_eai=missing_eai, manual_setup=manual_setup, row_errors=row_errors, warnings=warnings,
    )


def summary_log_lines(result):
    """Linhas finais do log (resumo, linhas ignoradas e avisos)."""
    lines = [
        f"\nConcluido: {result['generated']} application(oes) gerada(s), {result['conflicts']} conflito(s) "
        f"nao gerado(s), {len(result['row_errors'])} linha(s) ignorada(s), {len(result['warnings'])} aviso(s)."
    ]
    lines += ["  IGNORADO: " + e for e in result["row_errors"]]
    lines += ["  AVISO: " + w for w in result["warnings"]]
    return lines
