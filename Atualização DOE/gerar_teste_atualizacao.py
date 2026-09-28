# -*- coding: utf-8 -*-
"""Gera arquivo-teste de atualização SPGG a partir do último ato DOE por vaga."""

from __future__ import annotations

import re
import unicodedata
from datetime import datetime, date
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

FOLDER = Path(__file__).resolve().parent
DOE_FILE = FOLDER / "Cargos em Comissão Atualizado.xlsx"
DOE_CANDIDATOS = [
    FOLDER.parent / "AgenteDOE" / "FG_Cargos_Comissao_DOE_04-10-09-2026.xlsx",
]
SPGG_FILE = FOLDER / "TB_R_CARGOS_SPGG.xlsx"
OUT_FILE = FOLDER / "TESTE_atualizacao_SPGG_v5.xlsx"

ENTRY_ATOS = {"NOMEIA", "DESIGNA", "COLOCA", "ATRIBUI"}
EXIT_ATOS = {"EXONERA", "DISPENSA", "CESSA"}
OCCUPANCY_ATOS = ENTRY_ATOS | EXIT_ATOS
CANCEL_ATOS = {"TORNA SEM EFEITO", "DECLARA SEM EFEITO", "REVOGA", "SEM EFEITO"}
LOTACAO_ATOS = {"RETIFICA", "ALTERA", "APOSTILA", "RELOTA", "PRORROGA"}

ATO_PRIORITY = {}
for _a in ENTRY_ATOS:
    ATO_PRIORITY[_a] = 40
for _a in EXIT_ATOS:
    ATO_PRIORITY[_a] = 10
for _a in LOTACAO_ATOS:
    ATO_PRIORITY[_a] = 5
for _a in CANCEL_ATOS:
    ATO_PRIORITY[_a] = 0

HEADER_DATE_RE = re.compile(r"DATA\s+DO\s+DOE\s*:\s*(\d{1,2}/\d{1,2}/\d{4})", re.IGNORECASE)


def norm_name(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    try:
        if pd.isna(value):
            return ""
    except (ValueError, TypeError):
        pass
    text = str(value).strip()
    if text.lower() in {"", "nan", "none", "-", "<na>"}:
        return ""
    text = unicodedata.normalize("NFKD", text.upper())
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return re.sub(r"\s+", " ", text).strip()


def display_name(value) -> str | None:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    try:
        if pd.isna(value):
            return None
    except (ValueError, TypeError):
        pass
    text = str(value).strip()
    if text.lower() in {"", "nan", "none", "-", "<na>"}:
        return None
    return re.sub(r"\s+", " ", text).upper()


def has_name(value) -> bool:
    return display_name(value) is not None


def is_real_person(value) -> bool:
    key = norm_name(value)
    if not key or len(key) < 5:
        return False
    if "PRIMEIRO PROVIMENTO" in key or "PRIEMEIRO PROVIMENTO" in key:
        return False
    if key in {"SERVIDOR", "VAGA NOVA", "VAGO", "VAGA", "NAO INFORMADO"}:
        return False
    return True


def vaga_int(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, int):
        return value
    if isinstance(value, float) and value.is_integer():
        return int(value)
    text = str(value).strip()
    if not text or text.lower() in {"nan", "none", "-", "*"}:
        return None
    if "/" in text:
        return None
    digits = re.sub(r"\D", "", text)
    if not digits:
        return None
    try:
        return int(digits)
    except ValueError:
        return None


def parse_date(value):
    if value is None:
        return pd.NaT
    try:
        if pd.isna(value):
            return pd.NaT
    except (ValueError, TypeError):
        pass
    if isinstance(value, datetime):
        ts = pd.Timestamp(value)
        return ts if pd.isna(ts) else ts.normalize()
    if isinstance(value, date) and not isinstance(value, datetime):
        return pd.Timestamp(value)
    return pd.to_datetime(value, dayfirst=True, errors="coerce")


def fmt_date(value) -> str:
    ts = parse_date(value)
    if pd.isna(ts):
        return ""
    return ts.strftime("%d/%m/%Y")


def excel_date(value):
    ts = parse_date(value)
    if pd.isna(ts):
        return None
    return ts.to_pydatetime()


def ato_norm(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return re.sub(r"\s+", " ", str(value).strip().upper())


def as_id(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    try:
        if pd.isna(value):
            return None
    except (ValueError, TypeError):
        pass
    text = str(value).strip()
    if text.lower() in {"", "nan", "none", "-", "0"}:
        return None
    try:
        return int(float(text))
    except (TypeError, ValueError):
        return None


def dates_equal(a, b) -> bool:
    da, db = parse_date(a), parse_date(b)
    if pd.isna(da) and pd.isna(db):
        return True
    if pd.isna(da) or pd.isna(db):
        return False
    return da.normalize() == db.normalize()


def ids_equal(a, b) -> bool:
    return as_id(a) == as_id(b)


def load_doe(path=None, linha_offset: int = 0) -> pd.DataFrame:
    path = Path(path) if path is not None else DOE_FILE
    raw = pd.read_excel(path, sheet_name="FG e Cargos em Comissão", header=3)
    records = []
    section_date = pd.NaT
    section_label = ""

    for idx, row in raw.iterrows():
        nome_txt = "" if pd.isna(row.get("Nome")) else str(row.get("Nome")).strip()
        header_match = HEADER_DATE_RE.search(nome_txt)
        if header_match:
            section_date = parse_date(header_match.group(1))
            section_label = header_match.group(0)
            continue
        if re.search(r"DATA\s+DO\s+DOE\s*:\s*SEM\s+DATA", nome_txt, re.I):
            # Mantém a última data de seção válida; não zera a data dos atos abaixo.
            continue

        is_section = (
            pd.notna(row.get("Nome"))
            and pd.isna(row.get("Data DOE"))
            and pd.isna(row.get("Indicação"))
            and pd.isna(row.get("Vaga n°"))
        )
        if is_section:
            continue
        if raw.loc[[idx]].dropna(how="all").empty:
            continue

        data_celula = parse_date(row.get("Data DOE"))
        data_contar = parse_date(row.get("A contar de"))
        if pd.isna(data_contar):
            data_contar = parse_date(row.get("a contar de"))
        if pd.notna(data_contar):
            data_usada = data_contar
            origem_data = "A contar de"
        elif pd.notna(data_celula):
            data_usada = data_celula
            origem_data = "célula Data DOE"
        elif pd.notna(section_date):
            data_usada = section_date
            origem_data = f"seção {section_label}"
        else:
            data_usada = pd.NaT
            origem_data = "sem data"

        rec = row.to_dict()
        rec["linha_doe"] = int(idx + 5) + linha_offset
        rec["arquivo_doe"] = path.name
        rec["vaga_k"] = vaga_int(row.get("Vaga n°"))
        rec["data_celula"] = data_celula
        rec["data_contar"] = data_contar
        rec["data_secao"] = section_date
        rec["data_p"] = data_usada
        rec["origem_data"] = origem_data
        rec["ato"] = ato_norm(row.get("Indicação"))
        rec["pri"] = ATO_PRIORITY.get(rec["ato"], 15)
        rec["nome_disp"] = display_name(row.get("Nome"))
        rec["anterior_disp"] = display_name(row.get("Vaga Anterior"))
        records.append(rec)

    return pd.DataFrame(records)


def resolve_doe_fonte() -> Path:
    """Uma fonte por geração: o primeiro FG_Cargos_Comissao.xlsx que existir."""
    for path in DOE_CANDIDATOS:
        if path.exists():
            return path
    raise FileNotFoundError(
        "Não achei FG_Cargos_Comissao.xlsx em C:\\Atualização DOE nem em C:\\AgenteDOE"
    )


def load_all_doe() -> pd.DataFrame:
    fonte = resolve_doe_fonte()
    print("Fonte DOE (única):", fonte)
    return load_doe(fonte)


def _cancel_blob(row) -> str:
    parts = [row.get("Função"), row.get("Órgão / Secretaria"), row.get("Assunto")]
    text = " ".join(str(p) for p in parts if p is not None and not (isinstance(p, float) and pd.isna(p)))
    text = unicodedata.normalize("NFKD", text.upper())
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return re.sub(r"\s+", " ", text)


def classify_sem_efeito(row, grp: pd.DataFrame) -> str:
    """SEM EFEITO só cancela titular em posse/desistência/publicação indevida."""
    b = _cancel_blob(row)
    if re.search(
        r"NAO TER PREENCHIDO|NAO TER TOMADO POSSE|PRAZO LEGAL|DESISTENCIA DO CARGO|\bDESISTENCIA\b",
        b,
    ):
        return "cancela_posse"
    if re.search(r"CONSTADO INDEVIDAMENTE|CONSTOU INDEVIDAMENTE", b):
        return "cancela_publicacao"
    return "nao_mexe_titular"


def classify_retifica(row) -> str:
    """RETIFICA: lotação/texto vs confirmação de nomeação/designação vs FG/CC."""
    b = _cancel_blob(row)
    if re.search(r"ATO E DISPENSA|ATO E EXONERA|A EXONERACAO E|A DISPENSA E", b):
        return "confirma_saida"
    if re.search(r"TRATA DA NOMEACAO|SE TRATA DA NOMEACAO|NOMEIA PARA EXERCER|ATO E NOMEIA", b):
        return "confirma_nomeacao"
    if re.search(r"ATO E DESIGNA|DESIGNA PARA EXERCER", b):
        return "confirma_designacao"
    if re.search(r"FUNCAO GRATIFICADA", b) and re.search(r"CARGO COMISSIONADO", b):
        return "corrige_tipo_fg_cc"
    return "lotacao_texto"


def effective_occupancy(grp: pd.DataFrame) -> tuple[list, set]:
    """Devolve índices de atos de ocupação ainda válidos e os anulados por SEM EFEITO."""
    grp = grp.copy()
    ordered = grp.sort_values(["data_p", "linha_doe"], na_position="first")
    kept: list = []
    voided: set = set()
    for idx, row in ordered.iterrows():
        ato = row["ato"]
        if ato in CANCEL_ATOS:
            kind = classify_sem_efeito(row, grp)
            grp.at[idx, "sem_efeito_kind"] = kind
            if kind == "nao_mexe_titular":
                continue
            person = norm_name(row.get("nome_disp"))
            for pos in range(len(kept) - 1, -1, -1):
                prev = grp.loc[kept[pos]]
                if norm_name(prev.get("nome_disp")) == person and prev["ato"] in OCCUPANCY_ATOS:
                    voided.add(kept[pos])
                    kept.pop(pos)
                    break
            continue
        if ato in LOTACAO_ATOS:
            continue
        if ato in OCCUPANCY_ATOS:
            kept.append(idx)
    return kept, voided, grp


def select_last_acts(doe: pd.DataFrame):
    with_vaga = doe[doe["vaga_k"].notna()].copy()
    sem_vaga = doe[doe["vaga_k"].isna()].copy()

    selected_rows = []
    discarded_rows = []
    so_lotacao_rows = []

    for vaga, grp in with_vaga.groupby("vaga_k", sort=False):
        kept_idx, voided, grp_occ = effective_occupancy(grp)
        occ = grp_occ.loc[kept_idx].copy() if kept_idx else grp.iloc[0:0]

        if occ.empty:
            for _, row in grp.iterrows():
                item = row.to_dict()
                item["vaga_k"] = vaga
                if row.name in voided:
                    item["motivo_descarte"] = "anulado por TORNA/DECLARA SEM EFEITO"
                elif row["ato"] in LOTACAO_ATOS:
                    item["motivo_descarte"] = "ato de lotação/retificação — não define titular"
                elif row["ato"] in CANCEL_ATOS:
                    item["motivo_descarte"] = "ato de anulação (não define titular)"
                else:
                    item["motivo_descarte"] = "sem ato de ocupação válido nesta vaga"
                discarded_rows.append(item)
                so_lotacao_rows.append(item)
            continue

        dated = occ[occ["data_p"].notna()]
        pool = dated if len(dated) else occ
        last_date = pool["data_p"].max() if pool["data_p"].notna().any() else pd.NaT
        on_last = pool[pool["data_p"] == last_date] if pd.notna(last_date) else pool
        on_last = on_last.sort_values(["pri", "linha_doe"], ascending=[False, False])
        chosen = on_last.iloc[0].copy()

        saidas = on_last[on_last["ato"].isin(EXIT_ATOS)]
        if len(saidas) and chosen["ato"] in ENTRY_ATOS:
            saida_nome = display_name(saidas.sort_values("linha_doe").iloc[-1]["nome_disp"])
            chosen["saida_mesmo_dia"] = saida_nome if is_real_person(saida_nome) else None
        else:
            chosen["saida_mesmo_dia"] = None

        ret_tipo = []
        for _, rr in grp[grp["ato"] == "RETIFICA"].iterrows():
            kind = classify_retifica(rr)
            if kind in {"confirma_nomeacao", "confirma_designacao", "corrige_tipo_fg_cc"}:
                ret_tipo.append(rr)
        if ret_tipo:
            last_ret = max(
                ret_tipo,
                key=lambda r: (
                    r["data_p"] if pd.notna(r["data_p"]) else pd.Timestamp.min,
                    r.get("linha_doe") or 0,
                ),
            )
            chosen["retifica_tipo_funcao"] = last_ret.get("Função")
            chosen["retifica_tipo_orgao"] = last_ret.get("Órgão / Secretaria")
            chosen["retifica_kind"] = classify_retifica(last_ret)
        else:
            chosen["retifica_tipo_funcao"] = None
            chosen["retifica_tipo_orgao"] = None
            chosen["retifica_kind"] = None

        selected_rows.append(chosen)

        for _, row in grp.iterrows():
            if row.name == chosen.name:
                continue
            item = row.to_dict()
            item["vaga_k"] = vaga
            if row.name in voided:
                item["motivo_descarte"] = "anulado por TORNA/DECLARA SEM EFEITO"
            elif row["ato"] in LOTACAO_ATOS:
                item["motivo_descarte"] = "ato de lotação/retificação — não define titular"
            elif row["ato"] in CANCEL_ATOS:
                item["motivo_descarte"] = "ato de anulação (não define titular)"
            elif pd.notna(last_date) and row["data_p"] == last_date and row["ato"] in OCCUPANCY_ATOS:
                item["motivo_descarte"] = "mesmo dia: prevaleceu o ato de entrada (NOMEIA/DESIGNA)"
            elif row["ato"] in OCCUPANCY_ATOS:
                item["motivo_descarte"] = "data anterior ao último ato de ocupação"
            else:
                item["motivo_descarte"] = "não define titular"
            discarded_rows.append(item)

    selected = pd.DataFrame(selected_rows)
    discarded = pd.DataFrame(discarded_rows)
    so_lotacao = pd.DataFrame(so_lotacao_rows)
    return selected, discarded, sem_vaga, so_lotacao


def decide_names(last: pd.Series, spgg_nome, spgg_anterior) -> tuple[str | None, str | None, str]:
    ato = last["ato"]
    nome_doe = display_name(last.get("nome_disp"))
    anterior_doe = display_name(last.get("anterior_disp"))
    saida_mesmo_dia = display_name(last.get("saida_mesmo_dia"))
    nome_spgg = display_name(spgg_nome)
    anterior_spgg = display_name(spgg_anterior)
    if ato in EXIT_ATOS:
        novo_nome = None
        novo_anterior = nome_doe if is_real_person(nome_doe) else nome_spgg
        regra = "saída: limpa Nome atual; quem saiu vai para NOME ANTERIOR"
        return novo_nome, novo_anterior, regra

    if ato in ENTRY_ATOS:
        novo_nome = nome_doe
        if is_real_person(saida_mesmo_dia):
            novo_anterior = saida_mesmo_dia
            regra = "entrada: Nome DOE + quem saiu desta vaga no mesmo dia"
        elif is_real_person(anterior_doe):
            novo_anterior = anterior_doe
            regra = "entrada: Nome DOE + Vaga Anterior DOE"
        elif is_real_person(nome_spgg) and norm_name(nome_spgg) != norm_name(novo_nome):
            novo_anterior = nome_spgg
            regra = "entrada: Nome DOE + titular atual da SPGG (Vaga Anterior vazia)"
        else:
            novo_anterior = anterior_spgg if is_real_person(anterior_spgg) else None
            regra = "entrada: Nome DOE; sem nome anterior identificável"
        return novo_nome, novo_anterior, regra

    return nome_spgg, anterior_spgg, "sem ato de ocupação — titular da SPGG mantido"


def build_id_index(spgg: pd.DataFrame) -> dict[str, int]:
    mapping: dict[str, int] = {}
    for _, row in spgg.iterrows():
        nk = norm_name(row.get("NOME ANTERIOR"))
        idv = as_id(row.get("ID_RH ANT."))
        if nk and idv is not None:
            mapping[nk] = idv
    for _, row in spgg.iterrows():
        nk = norm_name(row.get("Nome"))
        idv = as_id(row.get("ID RH"))
        if nk and idv is not None:
            mapping[nk] = idv
    return mapping


def build_phone_index(spgg: pd.DataFrame) -> dict[str, str]:
    mapping: dict[str, str] = {}
    for _, row in spgg.iterrows():
        nk = norm_name(row.get("Nome"))
        phone = clean_phone(row.get("TELEFONE"))
        if nk and phone:
            mapping[nk] = phone
    return mapping


def build_code_defaults(spgg: pd.DataFrame) -> dict:
    defaults = {}
    tmp = spgg.copy()
    tmp["CC/FG"] = tmp["CC/FG"].astype(str).str.strip().str.upper()
    for (codigo, cc_fg), grp in tmp.groupby(["CÓDIGO", "CC/FG"], dropna=False):
        if pd.isna(codigo) or not str(cc_fg) or str(cc_fg) in {"NAN", "NONE"}:
            continue
        def mode_or_none(col):
            s = grp[col].dropna()
            return s.mode().iloc[0] if len(s) else None
        defaults[(str(codigo).strip(), str(cc_fg))] = {
            "fator": mode_or_none("FATOR"),
            "nivel": mode_or_none("NÍVEL"),
            "valor": mode_or_none("VALOR"),
        }
    return defaults


def resolve_id(name, mapping: dict[str, int], row_name, row_id, prefer_row: bool = False):
    if not has_name(name):
        return None, "em branco — sem nome"

    if prefer_row and norm_name(name) == norm_name(row_name) and as_id(row_id) is not None:
        return as_id(row_id), "acompanhou o nome nesta mesma vaga"

    found = mapping.get(norm_name(name))
    if found is not None:
        if norm_name(name) == norm_name(row_name) and as_id(row_id) == found:
            return found, "manteve — mesmo titular desta vaga"
        return found, "acompanhou o nome (já existia na SPGG)"

    return None, "em branco — nome novo (não existia na SPGG)"


def resolve_phone(name, mapping: dict[str, str], row_name, row_phone):
    if not has_name(name):
        return None, "em branco — sem nome"
    if norm_name(name) == norm_name(row_name) and clean_phone(row_phone):
        return clean_phone(row_phone), "acompanhou o nome nesta mesma vaga"
    found = mapping.get(norm_name(name))
    if found:
        return found, "acompanhou o nome (já existia na SPGG)"
    return None, "em branco — sem telefone cadastrado para este nome"


def vacancy_attrs_from_act(last, code_defaults: dict) -> dict:
    parsed = parse_funcao(last.get("Função"), last.get("Órgão / Secretaria"))
    attrs = {}
    ret_f = last.get("retifica_tipo_funcao")
    ret_o = last.get("retifica_tipo_orgao")
    if ret_f or ret_o:
        parsed_r = parse_funcao(ret_f, ret_o)
        if parsed_r:
            parsed = parsed_r
    if not parsed:
        enc = infer_encargo(last.get("Órgão / Secretaria"), last.get("Função"))
        if enc:
            attrs["encargo"] = enc
        return attrs
    attrs["cc_fg"] = parsed["cc_fg"]
    attrs["codigo"] = parsed["codigo"]
    attrs["nivel"] = parsed["nivel"]
    key = (parsed["codigo"], parsed["cc_fg"])
    defs = code_defaults.get(key) or {}
    if defs.get("fator") is not None:
        attrs["fator"] = defs["fator"]
    if defs.get("nivel") is not None:
        attrs["nivel"] = defs["nivel"]
    if defs.get("valor") is not None:
        attrs["valor"] = defs["valor"]
    enc = infer_encargo(last.get("Órgão / Secretaria"), last.get("Função"))
    if enc:
        attrs["encargo"] = enc
    orgao = infer_orgao(last.get("Órgão / Secretaria"))
    if orgao:
        attrs["orgao"] = orgao
    return attrs


ORGAO_ALIASES = [
    ("SISTEMAS PENAL", "SSPS"),
    ("SUSEPE", "SUSEPE"),
    ("SECRETARIA EXTRAORDINARIA GERAL DE GOVERNO", "GG"),
    ("GABINETE DO GOVERNADOR", "GG"),
    ("VICE-GOVERNADOR", "GVG"),
    ("CASA CIVIL", "CC"),
    ("CASA MILITAR", "CM"),
    ("SECRETARIA DA EDUCACAO", "SEDUC"),
    ("SECRETARIA DE EDUCACAO", "SEDUC"),
    ("SECRETARIA DA SAUDE", "SES"),
    ("SECRETARIA DE SAUDE", "SES"),
    ("SECRETARIA DA FAZENDA", "SEFAZ"),
    ("PLANEJAMENTO", "SPGG"),
    ("PROCURADORIA", "PGE"),
    ("SEGURANCA PUBLICA", "SSP"),
    ("POLICIA CIVIL", "PC"),
    ("BRIGADA", "BM"),
    ("TRABALHO E DESENVOLVIMENTO", "STDP"),
    ("AGRICULTURA", "SEAPI"),
    ("DESENVOLVIMENTO RURAL", "SDR"),
    ("DESENVOLVIMENTO SOCIAL", "SEDES"),
    ("DESENVOLVIMENTO ECONOMICO", "SEDEC"),
    ("JUSTICA, CIDADANIA", "SJCDH"),
    ("JUSTICA E CIDADANIA", "SJCDH"),
    ("TURISMO", "SETUR"),
    ("CULTURA", "SEDAC"),
    ("COMUNICACAO", "SECOM"),
    ("MEIO AMBIENTE", "SEMA"),
    ("HABITACAO", "SEHAB"),
    ("LOGISTICA E TRANSPORTES", "SELT"),
    ("OBRAS PUBLICAS", "SOP"),
    ("DESENVOLVIMENTO URBANO", "SEDUR"),
    ("ESPORTE", "SEL"),
    ("IRGA", "IRGA"),
    ("DETRAN", "DETRAN"),
    ("DAER", "DAER"),
    ("DEFESA CIVIL", "Defesa Civil"),
    ("BOMBEIRO", "CBM"),
    ("IPE PREV", "IPE PREV"),
    ("IPE SAUDE", "IPE SAÚDE"),
    ("AGERGS", "AGERGS"),
]

FUNC_PATTERNS = [
    (re.compile(r"\bCCS\s*-?\s*(\d{1,2})\b", re.I), "CC", "CCS/FGS"),
    (re.compile(r"\bFGS\s*-?\s*(\d{1,2})\b", re.I), "FG", "CCS/FGS"),
    (re.compile(r"\bCCT\s*-?\s*(\d{1,2})\b", re.I), "CC", "CCT/FGT"),
    (re.compile(r"\bFGT\s*-?\s*(\d{1,2})\b", re.I), "FG", "CCT/FGT"),
    (re.compile(r"CARGO COMISSIONADO SUPERIOR\s*-?\s*(\d{1,2})", re.I), "CC", "CCS/FGS"),
    (re.compile(r"FUN[CÇ][AÃ]O GRATIFICADA SUPERIOR\s*-?\s*(\d{1,2})", re.I), "FG", "CCS/FGS"),
    (re.compile(r"CARGO COMISSIONADO TRANSVERSAL\s*-?\s*(\d{1,2})", re.I), "CC", "CCT/FGT"),
    (re.compile(r"FUN[CÇ][AÃ]O GRATIFICADA TRANSVERSAL\s*-?\s*(\d{1,2})", re.I), "FG", "CCT/FGT"),
]


def parse_funcao(*parts) -> dict | None:
    text = " ".join(str(p) for p in parts if p is not None and not (isinstance(p, float) and pd.isna(p)))
    if not text.strip():
        return None
    blob = unicodedata.normalize("NFKD", text)
    blob = "".join(ch for ch in blob if not unicodedata.combining(ch))
    # RETIFICA: o valor novo vem depois de "para declarar que se trata"
    declare = re.search(
        r"para declarar que se trata.{0,120}?(Cargo Comissionado|Funcao Gratificada)\s+(Superior|Transversal)\s*-?\s*(\d{1,2})",
        blob,
        re.I,
    )
    if declare:
        kind, eixo, num = declare.group(1), declare.group(2), declare.group(3).zfill(2)
        cc_fg = "CC" if "cargo" in kind.lower() else "FG"
        prefix = "CCS/FGS" if "superior" in eixo.lower() else "CCT/FGT"
        return {"cc_fg": cc_fg, "codigo": f"{prefix} {num}", "nivel": num}

    for rx, cc_fg, prefix in FUNC_PATTERNS:
        m = rx.search(text)
        if m:
            num = m.group(1).zfill(2)
            return {"cc_fg": cc_fg, "codigo": f"{prefix} {num}", "nivel": num}
    return None


def infer_orgao(*parts) -> str | None:
    text = norm_name(" ".join(str(p) for p in parts if p is not None and not (isinstance(p, float) and pd.isna(p))))
    if not text:
        return None
    for needle, sigla in ORGAO_ALIASES:
        if needle in text:
            return sigla
    return None


def infer_encargo(*parts) -> str | None:
    text = " ".join(str(p) for p in parts if p is not None and not (isinstance(p, float) and pd.isna(p)))
    text = re.sub(r"[\r\n]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    m = re.search(
        r"(Assessor Superior I{1,3}|Assessor Especial V|Assessor Especial IV|Assessor Especial III|Assessor Especial II|Assessor Especial I|Assessor T[eé]cnico V|Assessor T[eé]cnico IV|Assessor T[eé]cnico III|Assessor T[eé]cnico II|Assessor T[eé]cnico I|Diretor de Departamento Adjunto|Diretor de Departamento|Diretor[- ]Geral|Chefe de Gabinete|Chefe de Divis[aã]o|Chefe de Se[cç][aã]o|Chefe de Setor|Secret[aá]rio Adjunto)",
        text,
        re.I,
    )
    if m:
        return m.group(1)
    return None


def clean_phone(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    text = str(value).replace("_x000d_", "").replace("_x000D_", "").replace("\r", "").replace("\n", "").strip()
    if text.lower() in {"", "nan", "none", "-", "0", "0.0"}:
        return None
    return text


def blank_if_nan(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, str) and v.strip().lower() in {"", "nan", "none", "-", "<na>"}:
        return None
    return v


def person_pack(spgg: pd.DataFrame) -> dict:
    packs = {}

    def put(nome, row, id_col, tel_from_nome=True):
        key = norm_name(nome)
        if not key:
            return
        pack = {
            "id": as_id(row.get(id_col)),
            "tel": clean_phone(row.get("TELEFONE")) if tel_from_nome else None,
            "via": blank_if_nan(row.get("VIA PARTIDO")),
            "carg": blank_if_nan(row.get("CARG_INDIC")),
            "indic": blank_if_nan(row.get("INDICAÇÃO")),
            "partido": blank_if_nan(row.get("PARTIDO")),
        }
        if key not in packs or (pack["id"] and not packs[key].get("id")):
            packs[key] = pack
        elif packs[key].get("id") is None and pack["id"]:
            packs[key] = pack

    for _, row in spgg.iterrows():
        put(row.get("Nome"), row, "ID RH", True)
        put(row.get("NOME ANTERIOR"), row, "ID_RH ANT.", False)
    return packs


def build_carga(audit, update_map, packs):
    rows = []
    for _, a in audit.iterrows():
        ato = str(a.get("Ato DOE usado") or "").upper()
        if ato in EXIT_ATOS:
            status = "LIVRE"
        elif ato in ENTRY_ATOS:
            status = "OCUPADO"
        else:
            status = a.get("STATUS SPGG")
        nova = str(a.get("Vaga nova")) == "SIM"
        nome = display_name(a.get("Nome SPGG (depois)"))
        saiu = display_name(a.get("NOME ANTERIOR (depois)"))
        p_in = packs.get(norm_name(nome), {})
        p_out = packs.get(norm_name(saiu), {})
        um = update_map.get(int(a["VAGA_1"]), {})
        rows.append(
            {
                "vaga": int(a["VAGA_1"]),
                "nova": nova,
                "nome": nome,
                "saiu": saiu,
                "id_rh": as_id(a.get("ID RH (depois)")) or p_in.get("id"),
                "id_ant": as_id(a.get("ID_RH ANT. (depois)")) or p_out.get("id"),
                "tel": p_in.get("tel"),
                "via": p_in.get("via") or "-",
                "carg": p_in.get("carg") or "-",
                "indic": p_in.get("indic") or "-",
                "partido": p_in.get("partido") or "-",
                "via_ant": p_out.get("via") or "-",
                "carg_ant": p_out.get("carg") or "-",
                "indic_ant": p_out.get("indic") or "-",
                "partido_ant": p_out.get("partido") or "-",
                "data": parse_date(a.get("DATA DOE SPGG (depois)")),
                "status": status,
                "cc_fg": blank_if_nan(a.get("CC/FG (depois)")),
                "codigo": blank_if_nan(a.get("CÓDIGO (depois)")),
                "encargo": blank_if_nan(a.get("ENCARGO (depois)")),
                "fator": um.get("fator"),
                "orgao": a.get("Órgão SPGG") if not nova else infer_orgao(a.get("Órgão DOE")),
                "unidade": a.get("Órgão DOE") if nova else None,
                "ato": ato,
            }
        )
    return rows


def infer_cc_fg(funcao) -> str | None:
    parsed = parse_funcao(funcao)
    return parsed["cc_fg"] if parsed else None


def make_audit_row(key, last, sp, novo_nome, novo_anterior, novo_id, origem_id,
                   novo_id_ant, origem_id_ant, data_depois, regra, vaga_nova, n_spgg=0):
    nome_antes = display_name(sp.get("Nome")) if sp is not None else None
    anterior_antes = display_name(sp.get("NOME ANTERIOR")) if sp is not None else None
    id_antes = as_id(sp.get("ID RH")) if sp is not None else None
    id_ant_antes = as_id(sp.get("ID_RH ANT.")) if sp is not None else None
    data_antes = parse_date(sp.get("DATA DOE")) if sp is not None else pd.NaT
    mudou_nome = norm_name(nome_antes) != norm_name(novo_nome)
    mudou_ant = norm_name(anterior_antes) != norm_name(novo_anterior)
    mudou_data = not dates_equal(data_antes, data_depois)
    mudou_id = not ids_equal(id_antes, novo_id)
    mudou_id_ant = not ids_equal(id_ant_antes, novo_id_ant)
    return {
        "VAGA_1": key,
        "Vaga nova": "SIM" if vaga_nova else "NÃO",
        "Órgão SPGG": None if sp is None else sp.get("Órgão"),
        "STATUS SPGG": "OCUPADO" if vaga_nova else (None if sp is None else sp.get("STATUS")),
        "DATA DOE SPGG (antes)": fmt_date(data_antes),
        "DATA DOE SPGG (depois)": fmt_date(data_depois),
        "Origem da data": last.get("origem_data"),
        "Nome SPGG (antes)": nome_antes,
        "Nome SPGG (depois)": novo_nome,
        "ID RH (antes)": id_antes,
        "ID RH (depois)": novo_id,
        "Origem do ID RH": origem_id,
        "NOME ANTERIOR (antes)": anterior_antes,
        "NOME ANTERIOR (depois)": novo_anterior,
        "ID_RH ANT. (antes)": id_ant_antes,
        "ID_RH ANT. (depois)": novo_id_ant,
        "Origem do ID_RH ANT.": origem_id_ant,
        "Mudou nome atual": "SIM" if mudou_nome else "NÃO",
        "Mudou nome anterior": "SIM" if mudou_ant else "NÃO",
        "Mudou data": "SIM" if mudou_data else "NÃO",
        "Mudou ID RH": "SIM" if mudou_id else "NÃO",
        "Alterou alguma coisa": "SIM",
        "Ato DOE usado": last["ato"],
        "Data DOE do ato": fmt_date(last["data_p"]),
        "Nome no ato DOE": last["nome_disp"],
        "Vaga Anterior no DOE": last["anterior_disp"],
        "Regra aplicada": regra,
        "Linha origem DOE": int(last["linha_doe"]),
        "VAGA_1 duplicada na SPGG": "SIM" if n_spgg > 1 else "NÃO",
        "Função DOE": last.get("Função"),
        "Órgão DOE": last.get("Órgão / Secretaria"),
    }, {
        "mudou_nome": mudou_nome,
        "mudou_ant": mudou_ant,
        "mudou_data": mudou_data,
        "mudou_id": mudou_id,
        "mudou_id_ant": mudou_id_ant,
    }


def build_updates(selected: pd.DataFrame, spgg: pd.DataFrame):
    spgg = spgg.copy()
    spgg["_vaga_k"] = spgg["VAGA_1"].map(vaga_int)
    id_index = build_id_index(spgg)
    phone_index = build_phone_index(spgg)
    code_defaults = build_code_defaults(spgg)

    first_by_vaga = (
        spgg.dropna(subset=["_vaga_k"])
        .drop_duplicates("_vaga_k", keep="first")
        .set_index("_vaga_k")
    )

    audit_rows = []
    no_match_rows = []
    update_map = {}

    for _, last in selected.iterrows():
        key = int(last["vaga_k"])
        vaga_existe = key in first_by_vaga.index
        vattrs = vacancy_attrs_from_act(last, code_defaults)

        if not vaga_existe:
            if last["ato"] not in ENTRY_ATOS:
                rec = last.to_dict()
                rec["motivo"] = "Vaga n° não existe na SPGG e o último ato não é nomeação/designação"
                no_match_rows.append(rec)
                continue
            novo_nome, novo_anterior, regra = decide_names(last, None, None)
            regra = "vaga nova (não existia na SPGG): cadastro do titular do último ato de ocupação"
            data_depois = last["data_p"]
            novo_id, origem_id = resolve_id(novo_nome, id_index, None, None)
            novo_id_ant, origem_id_ant = resolve_id(novo_anterior, id_index, None, None)
            novo_tel, origem_tel = resolve_phone(novo_nome, phone_index, None, None)
            audit_row, flags = make_audit_row(
                key, last, None, novo_nome, novo_anterior, novo_id, origem_id,
                novo_id_ant, origem_id_ant, data_depois, regra, True, 0,
            )
            audit_row["TELEFONE (depois)"] = novo_tel
            audit_row["Origem do telefone"] = origem_tel
            audit_row["CC/FG (depois)"] = vattrs.get("cc_fg")
            audit_row["CÓDIGO (depois)"] = vattrs.get("codigo")
            audit_row["ENCARGO (depois)"] = vattrs.get("encargo")
            update_map[key] = {
                "novo_nome": novo_nome,
                "novo_anterior": novo_anterior,
                "nova_data": data_depois,
                "novo_id": novo_id,
                "novo_id_ant": novo_id_ant,
                "novo_tel": novo_tel,
                "vaga_nova": True,
                "funcao": last.get("Função"),
                "orgao_doe": last.get("Órgão / Secretaria"),
                "ato": last["ato"],
                "mudou": True,
                "mudou_tel": True,
                **vattrs,
                **flags,
            }
            audit_rows.append(audit_row)
            continue

        sp = first_by_vaga.loc[key]
        novo_nome, novo_anterior, regra = decide_names(last, sp.get("Nome"), sp.get("NOME ANTERIOR"))
        nome_antes = display_name(sp.get("Nome"))
        anterior_antes = display_name(sp.get("NOME ANTERIOR"))
        id_antes = as_id(sp.get("ID RH"))
        id_ant_antes = as_id(sp.get("ID_RH ANT."))
        tel_antes = clean_phone(sp.get("TELEFONE"))
        data_depois = last["data_p"] if pd.notna(last["data_p"]) else parse_date(sp.get("DATA DOE"))

        novo_id, origem_id = resolve_id(
            novo_nome, id_index, nome_antes, id_antes, prefer_row=True
        )
        novo_id_ant, origem_id_ant = resolve_id(
            novo_anterior, id_index, nome_antes, id_antes, prefer_row=True
        )
        novo_tel, origem_tel = resolve_phone(novo_nome, phone_index, nome_antes, tel_antes)
        if (
            has_name(novo_anterior)
            and norm_name(novo_anterior) == norm_name(anterior_antes)
            and as_id(id_ant_antes) is not None
            and novo_id_ant is None
        ):
            novo_id_ant = id_ant_antes
            origem_id_ant = "manteve — mesmo nome anterior desta vaga"

        n_spgg = int((spgg["_vaga_k"] == key).sum())
        audit_row, flags = make_audit_row(
            key, last, sp, novo_nome, novo_anterior, novo_id, origem_id,
            novo_id_ant, origem_id_ant, data_depois, regra, False, n_spgg,
        )
        mudou_tel = clean_phone(tel_antes) != clean_phone(novo_tel)
        mudou_cc = vattrs.get("cc_fg") is not None and str(sp.get("CC/FG") or "").strip().upper() != vattrs["cc_fg"]
        mudou_cod = vattrs.get("codigo") is not None and str(sp.get("CÓDIGO") or "").strip() != vattrs["codigo"]
        flags["mudou_tel"] = mudou_tel
        flags["mudou_cc"] = mudou_cc
        flags["mudou_cod"] = mudou_cod
        mudou = any(flags.values())
        audit_row["Alterou alguma coisa"] = "SIM" if mudou else "NÃO"
        audit_row["TELEFONE (antes)"] = tel_antes
        audit_row["TELEFONE (depois)"] = novo_tel
        audit_row["Origem do telefone"] = origem_tel
        audit_row["CC/FG (antes)"] = sp.get("CC/FG")
        audit_row["CC/FG (depois)"] = vattrs.get("cc_fg") or sp.get("CC/FG")
        audit_row["CÓDIGO (antes)"] = sp.get("CÓDIGO")
        audit_row["CÓDIGO (depois)"] = vattrs.get("codigo") or sp.get("CÓDIGO")
        audit_row["ENCARGO (depois)"] = sp.get("ENCARGO")
        vattrs_exist = {k: v for k, v in vattrs.items() if k != "encargo"}
        update_map[key] = {
            "novo_nome": novo_nome,
            "novo_anterior": novo_anterior,
            "nova_data": data_depois,
            "novo_id": novo_id,
            "novo_id_ant": novo_id_ant,
            "novo_tel": novo_tel,
            "vaga_nova": False,
            "mudou": mudou,
            **vattrs_exist,
            **flags,
        }
        audit_rows.append(audit_row)

    audit = pd.DataFrame(audit_rows)
    no_match = pd.DataFrame(no_match_rows)
    return update_map, audit, no_match, id_index


def apply_to_spgg(spgg: pd.DataFrame, update_map: dict) -> pd.DataFrame:
    out = spgg.copy()
    out["_vaga_k"] = out["VAGA_1"].map(vaga_int)
    for i, row in out.iterrows():
        key = row["_vaga_k"]
        info = update_map.get(key)
        if not info or info.get("vaga_nova"):
            continue
        out.at[i, "Nome"] = info["novo_nome"]
        out.at[i, "NOME ANTERIOR"] = info["novo_anterior"]
        if pd.notna(info["nova_data"]):
            out.at[i, "DATA DOE"] = excel_date(info["nova_data"])
        out.at[i, "ID RH"] = info["novo_id"]
        out.at[i, "ID_RH ANT."] = info["novo_id_ant"]
        out.at[i, "TELEFONE"] = info.get("novo_tel") if info.get("novo_tel") is not None else 0
        if info.get("cc_fg"):
            out.at[i, "CC/FG"] = info["cc_fg"]
        if info.get("codigo"):
            out.at[i, "CÓDIGO"] = info["codigo"]
        # ENCARGO pertence à vaga (VAGA_1), não ao nome — não sobrescrever em vaga existente.
        if info.get("fator") is not None:
            out.at[i, "FATOR"] = info["fator"]
        if info.get("nivel") is not None:
            out.at[i, "NÍVEL"] = info["nivel"]
        if info.get("valor") is not None:
            out.at[i, "VALOR"] = info["valor"]

    novas = []
    for key, info in update_map.items():
        if not info.get("vaga_nova"):
            continue
        row = {col: None for col in out.columns if col != "_vaga_k"}
        row["VAGA_1"] = key
        if "VAGA_2" in row:
            row["VAGA_2"] = key
        row["Nome"] = info["novo_nome"]
        row["NOME ANTERIOR"] = info["novo_anterior"]
        row["DATA DOE"] = excel_date(info["nova_data"])
        row["ID RH"] = info["novo_id"]
        row["ID_RH ANT."] = info["novo_id_ant"]
        row["TELEFONE"] = info.get("novo_tel") if info.get("novo_tel") is not None else 0
        row["STATUS"] = "OCUPADO"
        row["ADM"] = "DIRETA"
        if info.get("orgao"):
            row["Órgão"] = info["orgao"]
        if info.get("cc_fg"):
            row["CC/FG"] = info["cc_fg"]
        if info.get("codigo"):
            row["CÓDIGO"] = info["codigo"]
        if info.get("encargo"):
            row["ENCARGO"] = info["encargo"]
        elif info.get("funcao"):
            row["ENCARGO"] = info.get("funcao")
        if info.get("fator") is not None:
            row["FATOR"] = info["fator"]
        if info.get("nivel") is not None:
            row["NÍVEL"] = info["nivel"]
        if info.get("valor") is not None:
            row["VALOR"] = info["valor"]
        if "UNIDADE ADM" in row:
            row["UNIDADE ADM"] = info.get("orgao_doe")
        if info.get("encargo") and re.search(r"diretor|chefe de gabinete|secretario", str(info.get("encargo")), re.I):
            row["CHEFIA/ASS."] = "CHEFIA"
        elif "CHEFIA/ASS." in row:
            row["CHEFIA/ASS."] = "ASS."
        for col in ("VIA PARTIDO", "CARG_INDIC", "INDICAÇÃO", "PARTIDO", "PARTIDO (GRÁFICO)"):
            if col in row:
                row[col] = "-"
        row["_vaga_k"] = key
        novas.append(row)

    if novas:
        out = pd.concat([out, pd.DataFrame(novas)], ignore_index=True)
    return out.drop(columns=["_vaga_k"])


def style_header(ws, n_cols, fill_hex="1F4E79"):
    fill = PatternFill("solid", fgColor=fill_hex)
    font = Font(color="FFFFFF", bold=True, name="Calibri", size=10)
    align = Alignment(vertical="center", wrap_text=True)
    for col in range(1, n_cols + 1):
        cell = ws.cell(1, col)
        cell.fill = fill
        cell.font = font
        cell.alignment = align
    ws.auto_filter.ref = ws.dimensions
    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 30


def autosize(ws, max_width=42):
    for col in ws.columns:
        letter = get_column_letter(col[0].column)
        longest = 0
        for cell in col[:80]:
            if cell.value is None:
                continue
            longest = max(longest, min(len(str(cell.value)), max_width))
        ws.column_dimensions[letter].width = max(12, longest + 2)


def write_df(ws, df: pd.DataFrame, header_hex="1F4E79"):
    if df is None or df.empty:
        ws["A1"] = "Nenhum registro"
        return
    view = df.copy()
    date_like = []
    for col in view.columns:
        if "data" in str(col).lower() or str(col).upper() == "DATA DOE":
            date_like.append(col)
        view[col] = view[col].apply(lambda v: "" if (v is None or (isinstance(v, float) and pd.isna(v)) or (not isinstance(v, (datetime, date)) and pd.isna(v) if not isinstance(v, (list, dict)) else False)) else v)

    for r_idx, row in enumerate(dataframe_to_rows(view, index=False, header=True), 1):
        for c_idx, value in enumerate(row, 1):
            cell = ws.cell(r_idx, c_idx, value)
            cell.font = Font(name="Calibri", size=10)
            cell.alignment = Alignment(vertical="center", wrap_text=False)
            header_name = view.columns[c_idx - 1] if r_idx > 1 else ""
            if r_idx > 1 and header_name in date_like and isinstance(value, (datetime, date)):
                cell.number_format = "DD/MM/YYYY"
            if r_idx > 1 and header_name in {"ID RH", "ID_RH ANT.", "ID RH (antes)", "ID RH (depois)", "ID_RH ANT. (antes)", "ID_RH ANT. (depois)", "VAGA_1", "Vaga n° (chave)"} and isinstance(value, (int, float)) and not isinstance(value, bool):
                cell.number_format = "0"
    style_header(ws, view.shape[1], header_hex)
    autosize(ws)


def paint_audit(ws, df: pd.DataFrame):
    if df is None or df.empty:
        return
    fill_yes = PatternFill("solid", fgColor="FFF2CC")
    fill_no = PatternFill("solid", fgColor="E2EFDA")
    fill_exit = PatternFill("solid", fgColor="FCE4D6")
    fill_new = PatternFill("solid", fgColor="DDEBF7")
    cols = list(df.columns)

    def col(name):
        return cols.index(name) + 1 if name in cols else None

    i_alt = col("Alterou alguma coisa")
    i_ato = col("Ato DOE usado")
    i_nome_d = col("Nome SPGG (depois)")
    i_ant_d = col("NOME ANTERIOR (depois)")
    i_data_d = col("DATA DOE SPGG (depois)")
    i_id_d = col("ID RH (depois)")
    i_origem_id = col("Origem do ID RH")
    for r in range(2, df.shape[0] + 2):
        alt = str(ws.cell(r, i_alt).value) if i_alt else ""
        ato = str(ws.cell(r, i_ato).value or "") if i_ato else ""
        origem = str(ws.cell(r, i_origem_id).value or "") if i_origem_id else ""
        if alt == "SIM":
            for i in (i_nome_d, i_ant_d, i_data_d, i_id_d):
                if i:
                    ws.cell(r, i).fill = fill_yes
        elif alt == "NÃO" and i_nome_d:
            ws.cell(r, i_nome_d).fill = fill_no
        if ato in EXIT_ATOS and i_nome_d:
            ws.cell(r, i_nome_d).fill = fill_exit
        if "nome novo" in origem and i_id_d:
            ws.cell(r, i_id_d).fill = fill_new


def paint_spgg_changes(ws, update_map: dict):
    headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]

    def idx(name):
        return headers.index(name) + 1 if name in headers else None

    i_vaga = idx("VAGA_1")
    targets = {
        "Nome": "mudou_nome",
        "NOME ANTERIOR": "mudou_ant",
        "DATA DOE": "mudou_data",
        "ID RH": "mudou_id",
        "ID_RH ANT.": "mudou_id_ant",
        "TELEFONE": "mudou_tel",
        "CC/FG": "mudou_cc",
        "CÓDIGO": "mudou_cod",
        "ENCARGO": "encargo",
    }
    fill = PatternFill("solid", fgColor="FFF2CC")
    fill_new = PatternFill("solid", fgColor="C6EFCE")
    for r in range(2, ws.max_row + 1):
        key = vaga_int(ws.cell(r, i_vaga).value)
        info = update_map.get(key)
        if not info or not info["mudou"]:
            continue
        if info.get("vaga_nova"):
            for c in range(1, len(headers) + 1):
                ws.cell(r, c).fill = fill_new
            continue
        for col_name, flag in targets.items():
            i = idx(col_name)
            if i and info.get(flag):
                ws.cell(r, i).fill = fill
                if col_name == "DATA DOE":
                    ws.cell(r, i).number_format = "DD/MM/YYYY"


def write_readme(ws, stats: dict):
    ws["A1"] = "TESTE v5 — conferência (carga segura para a SPGG online)"
    ws["A1"].font = Font(name="Calibri", size=16, bold=True, color="1F4E79")
    ws.merge_cells("A1:G1")

    lines = [
        "",
        "Arquivo só para conferência. A SPGG original (local e a da nuvem) NÃO foi alterada.",
        "Versão 5: pensada para você revisar e, depois, colar na original online sem quebrar PROCV.",
        "",
        "COMO USAR ESTE ARQUIVO",
        "1. Abra a aba Conferencia — é a mais fácil de filtrar (Evandro, Bruna, vaga nova, ID em branco).",
        "2. A aba CARGA_PROCV tem SÓ as colunas que a SPGG online aceita gravar. Essa é a que entra em PROCV na nuvem.",
        "3. SPGG_TESTE mostra o resultado visual (nome preenchido) só para conferir. NÃO cole essa aba inteira na original.",
        "",
        "O QUE NÃO COLAR NA ORIGINAL ONLINE",
        "Nome, TELEFONE, VIA PARTIDO, CARG_INDIC, INDICAÇÃO, PARTIDO, NOME ANTERIOR, NÍVEL, VALOR — na nuvem isso é fórmula do Banco de Talentos / da vaga.",
        "Se colar texto em cima, o PROCV da original quebra e o dado fica congelado.",
        "",
        "O QUE PODE IR PARA A NUVEM (aba CARGA_PROCV)",
        "VAGA_1 (chave), ID RH, ID_RH ANT., DATA DOE, STATUS, CC/FG, CÓDIGO, FATOR, ENCARGO, e, se for vaga nova: Órgão, ADM, CHEFIA/ASS., UNIDADE ADM.",
        "Pessoa já conhecida: vai o ID RH que já existia na SPGG baixada.",
        "Pessoa nova: ID RH fica vazio até cadastrar no Banco de Talentos (TB_R_RH).",
        "",
        "REGRAS DE OCUPAÇÃO (iguais às da v4)",
        "• Titular = último NOMEIA/DESIGNA/EXONERA/DISPENSA. RETIFICA não troca pessoa (caso Ana/Evandro).",
        "• Vaga que não existe + NOMEIA/DESIGNA = linha nova (caso Bruna).",
        "• CC/FG e CÓDIGO saem da função do ato de ocupação (CCS-13 → CC).",
        "• ENCARGO permanece da VAGA_1 (não acompanha o nome). Só é preenchido em vaga nova.",
        "",
        "REGRA DE CRUZAMENTO",
        "Chave: Vaga n° (DOE)  =  VAGA_1 (SPGG).",
        "O titular de cada vaga é o último ATO DE OCUPAÇÃO (NOMEIA, DESIGNA, COLOCA, EXONERA, DISPENSA, CESSA), não o último ato qualquer.",
        "",
        "COMO FOI ESCOLHIDO O ÚLTIMO ATO DE OCUPAÇÃO",
        "1. Considera só linhas de ato (ignora títulos ATOS DO GOVERNADOR / SECRETARIA).",
        "2. Se a coluna Data DOE da linha estiver vazia, usa a data da linha de seção 'DATA DO DOE: dd/mm/aaaa' imediatamente acima.",
        "3. APOSTILA é correção menor — não é movimentação. RETIFICA só ocupa titular se confirmar nomeação/designação ou FG/CC. RELOTA/PRORROGA/ALTERA não trocam titular.",
        "4. TORNA/DECLARA SEM EFEITO anula o último NOMEIA/EXONERA etc. da mesma pessoa naquela vaga.",
        "5. Entre os atos de ocupação que sobram, vale a data mais recente.",
        "6. Se no mesmo dia há EXONERA + NOMEIA, prevalece a entrada.",
        "7. Número de vaga com barra (PROA) não entra no cruzamento.",
        "",
        "O QUE FOI COPIADO / ATUALIZADO NA SPGG",
        "• Nome (titular atual)",
        "• NOME ANTERIOR",
        "• DATA DOE  — data do último ato de ocupação",
        "• ID RH     — acompanha o Nome",
        "• ID_RH ANT. — acompanha o NOME ANTERIOR",
        "Em vaga nova: também STATUS=OCUPADO, CC/FG (se der para ler da função), ENCARGO e UNIDADE ADM (texto do DOE). Partido/indicação ficam '-' para não inventar regra da SPGG.",
        "",
        "COMO O ID RH ACOMPANHA O NOME",
        "• Se a pessoa já existia na SPGG, o ID vai com ela — inclusive para a vaga nova (caso Bruna, ID 10177).",
        "• Se a pessoa só aparece no DOE, ID RH fica em branco.",
        "",
        "ABAS DESTE ARQUIVO",
        "1. LEIA-ME — estas regras.",
        "2. Resumo — totais.",
        "3. Conferencia — nome que entra / que sai, IDs, data, CC/FG, se é vaga nova. Use esta para revisar.",
        "4. CARGA_PROCV — só colunas graváveis, chave VAGA_1. Base para colar/PROCV na original online.",
        "5. Auditoria — detalhe técnico de cada campo.",
        "6. SPGG_TESTE — TUDO POR VAGA visual (não colar na nuvem).",
        "7. Vagas_novas — primeiro provimento.",
        "8. DOE_ultimo_ato / DOE_datas_preenchidas / DOE_atos_descartados / DOE_sem_match / DOE_sem_vaga.",
        "",
        "NÚMEROS DESTA RODADA",
        f"Atos reais no DOE (sem títulos): {stats['doe_atos']}",
        f"Atos com Data DOE vazia preenchida pela seção: {stats['datas_preenchidas']}",
        f"Vagas com ato de ocupação válido: {stats['doe_vagas']}",
        f"Atos que não definiram titular (lotação, mais antigos, anulados): {stats['descartados']}",
        f"Vagas aplicadas (existentes + novas): {stats['match']}",
        f"  — vagas novas criadas: {stats['vagas_novas']}",
        f"Vagas que mudaram algum campo: {stats['alteradas']}",
        f"  — ID RH acompanhou nome já existente: {stats['id_acompanhou']}",
        f"  — ID RH em branco (nome novo): {stats['id_novo']}",
        f"Vagas só com retificação/lotação (titular da SPGG mantido): {stats['so_lotacao']}",
        f"Vaga inexistente sem nomeação (não criada): {stats['sem_match']}",
        f"Atos DOE sem Vaga n° válida: {stats['sem_vaga']}",
        "",
        f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}",
    ]
    titles = {
        "COMO USAR ESTE ARQUIVO",
        "O QUE NÃO COLAR NA ORIGINAL ONLINE",
        "O QUE PODE IR PARA A NUVEM (aba CARGA_PROCV)",
        "REGRAS DE OCUPAÇÃO (iguais às da v4)",
        "REGRA DE CRUZAMENTO",
        "COMO FOI ESCOLHIDO O ÚLTIMO ATO DE OCUPAÇÃO",
        "O QUE FOI COPIADO / ATUALIZADO NA SPGG",
        "COMO O ID RH ACOMPANHA O NOME",
        "ABAS DESTE ARQUIVO",
        "NÚMEROS DESTA RODADA",
    }
    for i, line in enumerate(lines, start=2):
        ws[f"A{i}"] = line
        ws[f"A{i}"].font = Font(name="Calibri", size=11, bold=line in titles, color="1F4E79" if line in titles else "000000")
    ws.column_dimensions["A"].width = 160
    ws.row_dimensions[1].height = 22


def main():
    print("Lendo DOE e preenchendo datas de seção...")
    doe = load_doe()
    print("Selecionando último ato de ocupação por vaga...")
    selected, discarded, sem_vaga, so_lotacao = select_last_acts(doe)

    print("Lendo SPGG...")
    spgg = pd.read_excel(SPGG_FILE, sheet_name="TUDO POR VAGA")

    print("Cruzando nomes, datas, ID RH e vagas novas...")
    update_map, audit, no_match, _id_index = build_updates(selected, spgg)

    atos_count = doe[doe["vaga_k"].notna()].groupby("vaga_k").size()
    if not audit.empty:
        audit["Atos DOE nesta vaga"] = audit["VAGA_1"].map(lambda k: int(atos_count.get(k, 1)))
        sort_cols = [c for c in ("Vaga nova", "Alterou alguma coisa", "VAGA_1") if c in audit.columns]
        audit = audit.sort_values(sort_cols, ascending=[False, False, True][: len(sort_cols)])

    print("Aplicando na cópia da SPGG...")
    spgg_teste = apply_to_spgg(spgg, update_map)

    datas_preenchidas = doe[doe["origem_data"].astype(str).str.startswith("seção")].copy()
    id_novo = int((audit["Origem do ID RH"] == "em branco — nome novo (não existia na SPGG)").sum()) if not audit.empty else 0
    id_acompanhou = int(audit["Origem do ID RH"].astype(str).str.contains("acompanhou", na=False).sum()) if not audit.empty else 0
    n_novas = int((audit["Vaga nova"] == "SIM").sum()) if not audit.empty and "Vaga nova" in audit.columns else 0

    stats = {
        "doe_atos": int(len(doe)),
        "datas_preenchidas": int(len(datas_preenchidas)),
        "doe_vagas": int(selected.shape[0]),
        "descartados": int(len(discarded)),
        "match": int(len(audit)),
        "alteradas": int((audit["Alterou alguma coisa"] == "SIM").sum()) if not audit.empty else 0,
        "iguais": int((audit["Alterou alguma coisa"] == "NÃO").sum()) if not audit.empty else 0,
        "mudou_data": int((audit["Mudou data"] == "SIM").sum()) if not audit.empty else 0,
        "id_acompanhou": id_acompanhou,
        "id_novo": id_novo,
        "vagas_novas": n_novas,
        "so_lotacao": int(so_lotacao["vaga_k"].nunique()) if not so_lotacao.empty else 0,
        "sem_match": int(len(no_match)),
        "sem_vaga": int(len(sem_vaga)),
    }

    resumo = pd.DataFrame(
        [
            ["Atos reais no DOE", stats["doe_atos"]],
            ["Atos com Data DOE vazia preenchida pela seção DATA DO DOE", stats["datas_preenchidas"]],
            ["Vagas com último ato de ocupação", stats["doe_vagas"]],
            ["Atos que não definiram titular (lotação, antigos, anulados)", stats["descartados"]],
            ["Vagas aplicadas na SPGG (existentes + novas)", stats["match"]],
            ["  vagas novas criadas (primeiro provimento)", stats["vagas_novas"]],
            ["Vagas que mudaram nome, data e/ou ID", stats["alteradas"]],
            ["  delas, ID RH acompanhou nome já existente", stats["id_acompanhou"]],
            ["  delas, ID RH em branco (nome novo na SPGG)", stats["id_novo"]],
            ["Vagas só com retificação/lotação (titular mantido)", stats["so_lotacao"]],
            ["Vaga inexistente sem nomeação (não criada)", stats["sem_match"]],
            ["Atos DOE sem Vaga n° válida", stats["sem_vaga"]],
            ["Linhas da SPGG original", len(spgg)],
            ["Linhas da SPGG no teste (com vagas novas)", len(spgg_teste)],
        ],
        columns=["Indicador", "Quantidade"],
    )

    doe_last_view = selected[
        [
            "vaga_k",
            "Nome",
            "Vaga Anterior",
            "Data DOE",
            "data_p",
            "origem_data",
            "Indicação",
            "Indicação nova",
            "Proa",
            "Vaga n°",
            "Função",
            "Órgão / Secretaria",
            "linha_doe",
        ]
    ].copy()
    doe_last_view = doe_last_view.rename(
        columns={
            "vaga_k": "Vaga n° (chave)",
            "linha_doe": "Linha na planilha DOE",
            "Data DOE": "Data DOE na célula",
            "data_p": "Data DOE usada",
            "origem_data": "Origem da data",
        }
    )
    doe_last_view["Data DOE na célula"] = doe_last_view["Data DOE na célula"].map(fmt_date)
    doe_last_view["Data DOE usada"] = doe_last_view["Data DOE usada"].map(fmt_date)
    doe_last_view["Atos na vaga"] = doe_last_view["Vaga n° (chave)"].map(lambda k: int(atos_count.get(k, 1)))
    doe_last_view = doe_last_view.sort_values("Vaga n° (chave)")

    if not datas_preenchidas.empty:
        fill_view = datas_preenchidas[
            ["vaga_k", "Nome", "Indicação", "Data DOE", "data_p", "origem_data", "Vaga n°", "linha_doe"]
        ].copy()
        fill_view = fill_view.rename(
            columns={
                "vaga_k": "Vaga n° (chave)",
                "Data DOE": "Data DOE na célula (vazia)",
                "data_p": "Data preenchida da seção",
                "origem_data": "Seção usada",
                "linha_doe": "Linha na planilha DOE",
            }
        )
        fill_view["Data DOE na célula (vazia)"] = fill_view["Data DOE na célula (vazia)"].map(fmt_date)
        fill_view["Data preenchida da seção"] = fill_view["Data preenchida da seção"].map(fmt_date)
        fill_view = fill_view.sort_values(["Data preenchida da seção", "Linha na planilha DOE"])
    else:
        fill_view = pd.DataFrame()

    if not discarded.empty:
        disc_view = discarded[
            ["vaga_k", "Nome", "Vaga Anterior", "data_p", "origem_data", "Indicação", "Vaga n°", "motivo_descarte", "linha_doe"]
        ].copy()
        disc_view = disc_view.rename(
            columns={
                "vaga_k": "Vaga n° (chave)",
                "data_p": "Data DOE usada",
                "origem_data": "Origem da data",
                "linha_doe": "Linha na planilha DOE",
                "motivo_descarte": "Por que saiu",
            }
        )
        disc_view["Data DOE usada"] = disc_view["Data DOE usada"].map(fmt_date)
        disc_view = disc_view.sort_values(["Vaga n° (chave)", "Linha na planilha DOE"])
    else:
        disc_view = pd.DataFrame()

    if not no_match.empty:
        miss_view = no_match[
            ["vaga_k", "Nome", "Vaga Anterior", "data_p", "origem_data", "Indicação", "Vaga n°", "Função", "Órgão / Secretaria", "motivo", "linha_doe"]
        ].copy()
        miss_view = miss_view.rename(
            columns={
                "vaga_k": "Vaga n° (chave)",
                "data_p": "Data DOE usada",
                "origem_data": "Origem da data",
                "linha_doe": "Linha na planilha DOE",
            }
        )
        miss_view["Data DOE usada"] = miss_view["Data DOE usada"].map(fmt_date)
        miss_view = miss_view.sort_values("Vaga n° (chave)")
    else:
        miss_view = pd.DataFrame()

    if not sem_vaga.empty:
        sv_view = sem_vaga[
            ["Nome", "Data DOE", "data_p", "origem_data", "Indicação", "Vaga Anterior", "Proa", "Vaga n°", "Função", "Órgão / Secretaria", "linha_doe"]
        ].copy()
        sv_view["Data DOE"] = sv_view["Data DOE"].map(fmt_date)
        sv_view["data_p"] = sv_view["data_p"].map(fmt_date)
        sv_view = sv_view.rename(columns={"data_p": "Data DOE usada", "origem_data": "Origem da data", "linha_doe": "Linha na planilha DOE"})
    else:
        sv_view = pd.DataFrame()

    print("Escrevendo", OUT_FILE, "...")
    wb = Workbook()
    ws_readme = wb.active
    ws_readme.title = "LEIA-ME"
    write_readme(ws_readme, stats)

    ws_resumo = wb.create_sheet("Resumo")
    write_df(ws_resumo, resumo, "375623")

    conferencia = pd.DataFrame()
    carga = pd.DataFrame()
    if not audit.empty:
        conferencia = pd.DataFrame(
            {
                "VAGA_1": audit["VAGA_1"],
                "Vaga nova": audit.get("Vaga nova", "NÃO"),
                "Nome que entra": audit.get("Nome SPGG (depois)"),
                "ID RH": audit.get("ID RH (depois)"),
                "Origem do ID": audit.get("Origem do ID RH"),
                "Nome que sai": audit.get("NOME ANTERIOR (depois)"),
                "ID_RH ANT.": audit.get("ID_RH ANT. (depois)"),
                "DATA DOE": audit.get("DATA DOE SPGG (depois)"),
                "CC/FG (antes)": audit.get("CC/FG (antes)"),
                "CC/FG (depois)": audit.get("CC/FG (depois)"),
                "CÓDIGO (antes)": audit.get("CÓDIGO (antes)"),
                "CÓDIGO (depois)": audit.get("CÓDIGO (depois)"),
                "ENCARGO (depois)": audit.get("ENCARGO (depois)"),
                "Ato DOE": audit.get("Ato DOE usado"),
                "Órgão SPGG": audit.get("Órgão SPGG"),
                "Cole na original?": audit.apply(
                    lambda r: "CRIAR LINHA (vaga nova)"
                    if str(r.get("Vaga nova")) == "SIM"
                    else "ATUALIZAR pela VAGA_1",
                    axis=1,
                ),
            }
        )
        conferencia = conferencia.sort_values(["Vaga nova", "VAGA_1"], ascending=[False, True])

        def status_carga(row):
            ato = str(row.get("Ato DOE usado") or "").upper()
            if ato in EXIT_ATOS:
                return "LIVRE"
            if ato in ENTRY_ATOS:
                return "OCUPADO"
            return row.get("STATUS SPGG")

        carga = pd.DataFrame(
            {
                "VAGA_1": audit["VAGA_1"],
                "Ação": audit["Vaga nova"].map(lambda x: "CRIAR" if x == "SIM" else "ATUALIZAR")
                if "Vaga nova" in audit.columns
                else "ATUALIZAR",
                "ID RH": audit.get("ID RH (depois)"),
                "ID_RH ANT.": audit.get("ID_RH ANT. (depois)"),
                "DATA DOE": audit.get("DATA DOE SPGG (depois)"),
                "STATUS": audit.apply(status_carga, axis=1),
                "CC/FG": audit.get("CC/FG (depois)"),
                "CÓDIGO": audit.get("CÓDIGO (depois)"),
                "ENCARGO": audit.get("ENCARGO (depois)"),
                "Órgão (só vaga nova)": audit.apply(
                    lambda r: r.get("Órgão DOE") if str(r.get("Vaga nova")) == "SIM" else "",
                    axis=1,
                ),
                "Nome (NÃO COLAR — só conferência)": audit.get("Nome SPGG (depois)"),
                "Nome que sai (NÃO COLAR — só conferência)": audit.get("NOME ANTERIOR (depois)"),
            }
        )
        carga = carga.sort_values(["Ação", "VAGA_1"], ascending=[False, True])

    ws_conf = wb.create_sheet("Conferencia")
    write_df(ws_conf, conferencia, "2E75B6")

    ws_carga = wb.create_sheet("CARGA_PROCV")
    write_df(ws_carga, carga, "375623")

    ws_audit = wb.create_sheet("Auditoria")
    write_df(ws_audit, audit, "C65911")
    paint_audit(ws_audit, audit)

    ws_spgg = wb.create_sheet("SPGG_TESTE")
    write_df(ws_spgg, spgg_teste, "1F4E79")
    paint_spgg_changes(ws_spgg, update_map)

    if not audit.empty and "Vaga nova" in audit.columns:
        novas_view = audit[audit["Vaga nova"] == "SIM"].copy()
        keep_novas = [
            c
            for c in (
                "VAGA_1",
                "Nome SPGG (depois)",
                "ID RH (depois)",
                "TELEFONE (depois)",
                "CC/FG (depois)",
                "CÓDIGO (depois)",
                "ENCARGO (depois)",
                "Origem do ID RH",
                "NOME ANTERIOR (depois)",
                "DATA DOE SPGG (depois)",
                "Ato DOE usado",
                "Função DOE",
                "Órgão DOE",
                "Regra aplicada",
            )
            if c in novas_view.columns
        ]
        novas_view = novas_view[keep_novas] if keep_novas else novas_view
    else:
        novas_view = pd.DataFrame()

    ws_novas = wb.create_sheet("Vagas_novas")
    write_df(ws_novas, novas_view, "548235")

    ws_last = wb.create_sheet("DOE_ultimo_ato")
    write_df(ws_last, doe_last_view, "2F5496")

    ws_fill = wb.create_sheet("DOE_datas_preenchidas")
    write_df(ws_fill, fill_view, "2E75B6")

    ws_disc = wb.create_sheet("DOE_atos_descartados")
    write_df(ws_disc, disc_view, "7F6000")

    ws_miss = wb.create_sheet("DOE_sem_match")
    write_df(ws_miss, miss_view, "C00000")

    ws_sv = wb.create_sheet("DOE_sem_vaga")
    write_df(ws_sv, sv_view, "833C0C")

    wb.save(OUT_FILE)
    print("OK:", OUT_FILE)
    print(stats)


if __name__ == "__main__":
    main()
