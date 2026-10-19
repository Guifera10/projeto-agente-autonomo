#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Orquestrador autônomo DOE → extração FG → SPGG.

Para cada dia desde a última pasta em `resultados/` até hoje:
consulta o DOE tipo 1, baixa a última atualização do dia (`nroEdicao`
quando há extra), confere tamanho/páginas, extrai FG e aplica na SPGG.

A base de cada atualização é o último `SPGG_PARA_ENVIAR_NUVEM.xlsx` de
uma pasta **anterior** à data do DOE. Extra do mesmo dia reprocessa a
pasta do dia em cima dessa base (a extra é o diário atualizado).

O resultado vai para `resultados/AAAA-MM-DD/` — um por dia. Se a API
ainda não tiver o DOE de hoje, a vigia consulta de 10 em 10 min até
gravar a pasta; com o resultado pronto, encerra (horários normais 07:00
e 19:00). Extra nova no ciclo das 19:00 reprocessa a pasta. Extração com
0 atos grava a pasta e **não** marca sucesso definitivo. Não envia nada à nuvem.

Uso:
  python orquestrador_doe.py --so-detectar
  python orquestrador_doe.py --so-baixar
  python orquestrador_doe.py
  python orquestrador_doe.py --vigia
  python orquestrador_doe.py --watch --intervalo 15
"""

from __future__ import annotations

import argparse
import json
import shutil
import socket
import subprocess
import sys
import time
import urllib.error
import urllib.request
from datetime import date, datetime, timedelta
from pathlib import Path
from urllib.parse import urlparse

API = "https://doe-backend.pro.rs.gov.br"
TIPO_DOE = 1
UA = "OrquestradorDOE/1.0"
TENTATIVAS_REDE = 6
TENTATIVAS_PDF = 3
PDF_MIN_BYTES = 10_000
ESPERA_API_S = 12 * 60
INTERVALO_VIGIA_MIN = 10
HORA_SEM_DOE = 19
HORA_FIM_VIGIA = 23

AQUI = Path(__file__).resolve().parent
PASTA_RESULTADOS = AQUI / "resultados"
DATA_INICIO_PADRAO = date(2026, 9, 11)


def log(msg: str) -> None:
    agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{agora}] {msg}", flush=True)


def pasta_resultados() -> Path:
    PASTA_RESULTADOS.mkdir(parents=True, exist_ok=True)
    return PASTA_RESULTADOS


def pasta_do_dia(item: dict | None = None) -> Path:
    """Pasta de resultados da edição: resultados/AAAA-MM-DD/."""
    base = pasta_resultados()
    dia = str((item or {}).get("data") or "")[:10]
    if len(dia) == 10 and dia[4] == "-" and dia[7] == "-":
        dest = base / dia
        dest.mkdir(parents=True, exist_ok=True)
        return dest
    return base


def pasta_tem_resultado(dia: date) -> bool:
    return (pasta_resultados() / dia.isoformat() / "ultimo_resultado.txt").exists()


def registrar_dia_sem_doe(dia: date) -> None:
    """Garante resultados/AAAA-MM-DD/ mesmo sem edição tipo 1 (fim de semana, feriado)."""
    if pasta_tem_resultado(dia):
        return
    dest = pasta_resultados() / dia.isoformat()
    dest.mkdir(parents=True, exist_ok=True)
    base_spgg, _base_dia = achar_spgg_anterior(dia)
    if base_spgg:
        copiar_seguro(base_spgg, dest / "SPGG_PARA_ENVIAR_NUVEM.xlsx")
        chec = base_spgg.parent / "CHECAGEM_atualizacao_SPGG.xlsx"
        if chec.exists():
            copiar_seguro(chec, dest / chec.name)
    copiar_resultado(estado_path(), "estado_doe.json", pasta_dest=dest)
    gravar_resumo(
        [
            f"ciclo: {datetime.now().isoformat(timespec='seconds')}",
            "modo: sem DOE tipo 1",
            f"data: {dia.isoformat()}",
            "edicao: nenhuma",
            f"base_spgg: {base_spgg or ''}",
            f"pasta_resultados: {dest}",
        ],
        pasta_dest=dest,
    )
    log(f"Pasta do dia sem DOE: {dest}")


def arquivo_em_uso(exc: BaseException) -> bool:
    if isinstance(exc, PermissionError):
        return True
    return getattr(exc, "winerror", None) in {5, 32}


def copiar_seguro(origem: Path, destino: Path) -> None:
    destino.parent.mkdir(parents=True, exist_ok=True)
    try:
        shutil.copy2(origem, destino)
    except OSError as exc:
        if arquivo_em_uso(exc):
            raise RuntimeError(
                f"Não consegui gravar {destino.name}: o Excel está com o arquivo aberto. "
                f"Feche {destino} e rode de novo."
            ) from exc
        raise


def copiar_resultado(
    origem: Path | str | None,
    destino_nome: str | None = None,
    pasta_dest: Path | None = None,
) -> Path | None:
    if not origem:
        return None
    origem = Path(origem)
    if not origem.exists():
        return None
    dest = (pasta_dest or pasta_resultados()) / (destino_nome or origem.name)
    copiar_seguro(origem, dest)
    log(f"Resultado em {dest}")
    return dest


def gravar_resumo(linhas: list[str], pasta_dest: Path | None = None) -> Path:
    texto = "\n".join(linhas) + "\n"
    dest = (pasta_dest or pasta_resultados()) / "ultimo_resultado.txt"
    dest.write_text(texto, encoding="utf-8")
    raiz = pasta_resultados() / "ultimo_resultado.txt"
    if dest.resolve() != raiz.resolve():
        raiz.write_text(texto, encoding="utf-8")
    log(f"Resumo: {dest}")
    return dest


def trazer_saidas(item: dict, pasta_atualizacao: Path | None) -> None:
    """Copia PDF, extrato FG deste ciclo, SPGG e checagem para resultados/AAAA-MM-DD/."""
    dest = pasta_do_dia(item)
    copiar_resultado(item.get("pdf"), pasta_dest=dest)
    extrato = item.get("extrato")
    if extrato:
        origem = Path(extrato)
        copiar_resultado(origem, "FG_Cargos_Comissao.xlsx", pasta_dest=dest)
        datado = item.get("extrato_datado")
        if datado:
            copiar_resultado(datado, pasta_dest=dest)
        elif origem.name != "FG_Cargos_Comissao.xlsx":
            copiar_resultado(origem, pasta_dest=dest)
    if item.get("spgg") and pasta_atualizacao:
        copiar_resultado(item.get("spgg"), "SPGG_PARA_ENVIAR_NUVEM.xlsx", pasta_dest=dest)
        copiar_resultado(
            pasta_atualizacao / "CHECAGEM_atualizacao_SPGG.xlsx",
            pasta_dest=dest,
        )
    copiar_resultado(estado_path(), "estado_doe.json", pasta_dest=dest)


def _backoff(tentativa: int) -> None:
    espera = min(60, 5 * (2 ** tentativa))
    log(f"Nova tentativa em {espera}s...")
    time.sleep(espera)


def _host_api() -> tuple[str, int]:
    parsed = urlparse(API)
    host = parsed.hostname or "doe-backend.pro.rs.gov.br"
    port = parsed.port or (443 if (parsed.scheme or "https") == "https" else 80)
    return host, port


def acordar_rede() -> bool:
    """Abre TCP no host da API para acordar Wi-Fi / 802.1X após ocioso."""
    host, port = _host_api()
    try:
        with socket.create_connection((host, port), timeout=8):
            pass
        return True
    except OSError as exc:
        log(f"Sem rota para {host}:{port} ({exc})")
        return False


def aguardar_api(orcamento_s: int = ESPERA_API_S) -> None:
    """Espera a API responder. Cobre NIC dormindo e instabilidade matinal da PROCERGS."""
    url = f"{API}/public/diarios/consultar/?tipoDiario={TIPO_DOE}&tipoRetorno=2"
    inicio = time.monotonic()
    tentativa = 0
    while True:
        acordar_rede()
        try:
            http_json(url, timeout=20, tentativas=1)
            if tentativa:
                log("API do DOE voltou a responder.")
            else:
                log("API do DOE acessível.")
            return
        except (urllib.error.URLError, TimeoutError, OSError) as exc:
            decorrido = int(time.monotonic() - inicio)
            if decorrido >= orcamento_s:
                log(f"API do DOE inacessível após {decorrido}s.")
                raise
            espera = min(60, 5 * (2 ** min(tentativa, 3)))
            log(
                f"API indisponível ({exc}). "
                f"Nova tentativa em {espera}s (já {decorrido}s de {orcamento_s}s)."
            )
            time.sleep(espera)
            tentativa += 1


def http_json(url: str, timeout: int = 45, tentativas: int = TENTATIVAS_REDE):
    ultimo: BaseException | None = None
    for i in range(tentativas):
        try:
            req = urllib.request.Request(
                url, headers={"User-Agent": UA, "Accept": "application/json"}
            )
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                raw = resp.read().decode("utf-8").strip()
            if not raw:
                raise ValueError("resposta vazia da API")
            return json.loads(raw)
        except urllib.error.HTTPError as exc:
            if 400 <= exc.code < 500:
                raise ValueError(f"HTTP {exc.code} em {url}") from exc
            ultimo = exc
            log(f"Falha de rede HTTP {exc.code} ({exc})")
        except json.JSONDecodeError:
            raise
        except ValueError:
            raise
        except (urllib.error.URLError, TimeoutError, OSError) as exc:
            ultimo = exc
            log(f"Falha de rede ({exc})")
        if i + 1 < tentativas:
            _backoff(i)
    assert ultimo is not None
    raise ultimo


def http_download(
    url: str,
    destino: Path,
    timeout: int = 300,
    tentativas: int = TENTATIVAS_REDE,
) -> None:
    destino.parent.mkdir(parents=True, exist_ok=True)
    ultimo: BaseException | None = None
    for i in range(tentativas):
        tmp = destino.with_suffix(destino.suffix + ".part")
        try:
            req = urllib.request.Request(
                url, headers={"User-Agent": UA, "Accept": "application/pdf"}
            )
            with urllib.request.urlopen(req, timeout=timeout) as resp, tmp.open("wb") as out:
                shutil.copyfileobj(resp, out)
            tmp.replace(destino)
            return
        except urllib.error.HTTPError as exc:
            tmp.unlink(missing_ok=True)
            if 400 <= exc.code < 500:
                raise
            ultimo = exc
            log(f"Download HTTP {exc.code} ({exc})")
        except (urllib.error.URLError, TimeoutError, OSError) as exc:
            tmp.unlink(missing_ok=True)
            ultimo = exc
            log(f"Download falhou ({exc})")
        if i + 1 < tentativas:
            _backoff(i)
    assert ultimo is not None
    raise ultimo


def achar_pasta_agente(cli: Path | None) -> Path:
    if cli:
        return cli.resolve()
    candidatos = [
        AQUI / "AgenteDOE",
        AQUI.parent / "AgenteDOE",
        AQUI,
    ]
    for pasta in candidatos:
        if (pasta / "agente_fg_comissao.py").exists():
            return pasta.resolve()
    raise FileNotFoundError(
        "Não achei agente_fg_comissao.py. Use --pasta-agente."
    )


def achar_pasta_atualizacao(cli: Path | None) -> Path:
    if cli:
        return cli.resolve()
    candidatos = [
        AQUI / "Atualizacao",
        AQUI / "Atualização DOE",
        AQUI,
        AQUI.parent / "Atualizacao",
        AQUI.parent / "Atualização DOE",
        AQUI.parent,
    ]
    for pasta in candidatos:
        if (pasta / "gerar_spgg_nuvem_limpo.py").exists():
            return pasta.resolve()
    raise FileNotFoundError(
        "Não achei gerar_spgg_nuvem_limpo.py. Use --pasta-atualizacao."
    )


def estado_path(_pasta: Path | None = None) -> Path:
    return AQUI / "estado_doe.json"


def data_edicao(item: dict) -> date | None:
    raw = str(item.get("data") or "")[:10]
    try:
        return date.fromisoformat(raw)
    except ValueError:
        return None


def achar_spgg_anterior(antes_de: date) -> tuple[Path | None, date | None]:
    """Último SPGG_PARA_ENVIAR_NUVEM.xlsx em pasta estritamente anterior a `antes_de`."""
    melhor_spgg: Path | None = None
    melhor_dia: date | None = None
    raiz = pasta_resultados()
    if not raiz.exists():
        return None, None
    for pasta in raiz.iterdir():
        if not pasta.is_dir():
            continue
        try:
            dia = date.fromisoformat(pasta.name)
        except ValueError:
            continue
        if dia >= antes_de:
            continue
        spgg = pasta / "SPGG_PARA_ENVIAR_NUVEM.xlsx"
        if not spgg.exists():
            continue
        if melhor_dia is None or dia > melhor_dia:
            melhor_dia = dia
            melhor_spgg = spgg
    return melhor_spgg, melhor_dia


def achar_ultima_spgg() -> tuple[Path | None, date | None]:
    """Último SPGG_PARA_ENVIAR_NUVEM.xlsx em resultados/AAAA-MM-DD/."""
    return achar_spgg_anterior(date.max)


def data_minima(piso: date, ultima_dia: date | None) -> date:
    """Início da varredura: última pasta (para pegar extra do mesmo dia) ou o piso."""
    if ultima_dia is None:
        return piso
    return ultima_dia if ultima_dia >= piso else piso


def carregar_estado(pasta: Path) -> dict:
    caminho = estado_path(pasta)
    if not caminho.exists():
        return {"processados": []}
    try:
        return json.loads(caminho.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return {"processados": []}


def salvar_estado(pasta: Path, estado: dict) -> None:
    estado["atualizado_em"] = datetime.now().isoformat(timespec="seconds")
    estado_path(pasta).write_text(
        json.dumps(estado, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def chave_edicao(item: dict) -> str:
    return f"{item['id']}|{item.get('data')}|{item.get('numero')}|{item.get('tipo')}"


def ja_processado(estado: dict, item: dict) -> bool:
    return chave_edicao(item) in set(estado.get("processados") or [])


def marcar_processado(estado: dict, item: dict) -> None:
    lista = list(estado.get("processados") or [])
    k = chave_edicao(item)
    if k not in lista:
        lista.append(k)
    estado["processados"] = lista[-200:]
    estado["ultimo"] = {
        "id": item["id"],
        "data": item.get("data"),
        "numero": item.get("numero"),
        "tipo": item.get("tipo"),
        "pdf": item.get("pdf"),
        "extrato": item.get("extrato"),
        "nro_edicao_pdf": item.get("nro_edicao_pdf"),
    }


def consultar_ultima() -> dict:
    return http_json(
        f"{API}/public/diarios/consultar/?tipoDiario={TIPO_DOE}&tipoRetorno=2"
    )


def consultar_dia(data_iso: str) -> dict:
    return http_json(
        f"{API}/public/diarios/consultar/?tipoDiario={TIPO_DOE}&data={data_iso}&tipoRetorno=2"
    )


def _item_ordinaria(detalhe: dict, data_iso: str) -> dict:
    ident = int(detalhe["id"])
    return {
        "id": ident,
        "data": detalhe.get("dataString") or data_iso,
        "data_br": detalhe.get("dataPublicacao"),
        "numero": detalhe.get("numeroEdicao"),
        "tipo": "ORDINARIA",
        "paginas": detalhe.get("qtdPaginas"),
        "id_download": ident,
        "nro_edicao_pdf": None,
    }


def _item_extra(extra: dict, detalhe: dict, data_iso: str) -> dict:
    nro = extra.get("nroEdicao")
    return {
        "id": int(extra["id"]),
        "data": data_iso,
        "data_br": extra.get("dataPublicacao") or detalhe.get("dataPublicacao"),
        "numero": nro,
        "tipo": f"EXTRA_{nro}",
        "paginas": extra.get("qtdPaginas"),
        "id_download": int(detalhe["id"]),
        "nro_edicao_pdf": int(nro) if nro is not None else None,
    }


def edicao_do_dia(data_iso: str) -> dict | None:
    """DOE tipo 1 do dia: última extra (nroEdicao) ou a ordinária."""
    try:
        detalhe = consultar_dia(data_iso)
    except (ValueError, json.JSONDecodeError) as exc:
        log(f"{data_iso}: sem DOE tipo 1 ({exc})")
        return None
    if not detalhe or not detalhe.get("id"):
        log(f"{data_iso}: sem DOE tipo 1")
        return None
    ordinaria = _item_ordinaria(detalhe, data_iso)
    extras = sorted(
        detalhe.get("edicoes") or [],
        key=lambda e: int(e.get("nroEdicao") or 0),
    )
    if not extras:
        return ordinaria
    escolhida = _item_extra(extras[-1], detalhe, data_iso)
    log(
        f"{data_iso}: última atualização = {escolhida['tipo']} "
        f"(nroEdicao={escolhida.get('nro_edicao_pdf')}, {escolhida.get('paginas')} pág.)"
    )
    return escolhida


def montar_itens_periodo(inicio: date, fim: date) -> tuple[list[dict], list[date]]:
    """Uma edição por dia (última atualização) de inicio até fim, inclusive."""
    itens: list[dict] = []
    sem_doe: list[date] = []
    dia = inicio
    while dia <= fim:
        iso = dia.isoformat()
        item = edicao_do_dia(iso)
        if item:
            itens.append(item)
        else:
            sem_doe.append(dia)
        dia += timedelta(days=1)
    try:
        ultima = consultar_ultima()
        iso = str(ultima.get("dataString") or "")[:10]
        if iso and not any(i.get("data") == iso for i in itens):
            d_ult = date.fromisoformat(iso)
            if inicio <= d_ult <= fim:
                item = edicao_do_dia(iso)
                if item:
                    itens.append(item)
                    itens.sort(key=lambda i: str(i.get("data") or ""))
                    sem_doe = [d for d in sem_doe if d != d_ult]
                    log(f"Incluída última edição da API: {iso}")
    except (ValueError, json.JSONDecodeError, urllib.error.URLError, TimeoutError, OSError) as exc:
        log(f"Não conferi consultar_ultima: {exc}")
    return itens, sem_doe


def montar_itens(incluir_extras: bool) -> list[dict]:
    """Compatibilidade: última ordinária do dia (e extra se pedido). Prefira montar_itens_periodo."""
    ultima = consultar_ultima()
    data_iso = ultima.get("dataString") or ""
    itens = [_item_ordinaria(ultima, data_iso)]
    if incluir_extras and data_iso:
        detalhe = consultar_dia(data_iso)
        for extra in detalhe.get("edicoes") or []:
            itens.append(_item_extra(extra, detalhe, data_iso))
    return itens


def contar_paginas(pdf: Path) -> int | None:
    try:
        import pdfplumber

        with pdfplumber.open(str(pdf)) as doc:
            return len(doc.pages)
    except Exception as exc:
        log(f"Não consegui contar páginas de {pdf.name}: {exc}")
        return None


def pdf_inteiro(caminho: Path, item: dict, exigir_paginas: bool = True) -> bool:
    if not caminho.exists() or caminho.stat().st_size < PDF_MIN_BYTES:
        return False
    esperado = item.get("paginas")
    if not esperado:
        return True
    pags = contar_paginas(caminho)
    if pags is None:
        return not exigir_paginas
    # Extra via nroEdicao devolve o diário completo; qtdPaginas da extra
    # é só o caderno extra. Incompleto = menos páginas do que a API informa.
    if int(pags) < int(esperado):
        log(f"{caminho.name}: {pags} pág. no arquivo, API informa {esperado}")
        return False
    return True


def url_pdf(item: dict) -> str:
    ident = item.get("id_download") or item["id"]
    url = f"{API}/public/diarios/download/{ident}/pdf"
    nro = item.get("nro_edicao_pdf")
    if nro:
        url += f"?nroEdicao={nro}"
    return url


def baixar_pdf(item: dict, pasta_pdf: Path) -> Path:
    pasta_pdf.mkdir(parents=True, exist_ok=True)
    ident = item.get("id_download") or item["id"]
    nome = f"DOE_{item['data']}_{item['tipo']}_{ident}.pdf"
    destino = pasta_pdf / nome
    url = url_pdf(item)
    if destino.exists() and pdf_inteiro(destino, item):
        log(f"PDF já em disco e conferido: {destino.name}")
        return destino
    if destino.exists():
        log(f"PDF em disco incompleto ou desatualizado: {destino.name}. Rebaixando.")
        destino.unlink(missing_ok=True)

    ultimo_erro = "download falhou"
    for tentativa in range(1, TENTATIVAS_PDF + 1):
        log(f"Baixando {url} ({tentativa}/{TENTATIVAS_PDF})")
        http_download(url, destino)
        tam = destino.stat().st_size
        if tam < PDF_MIN_BYTES:
            ultimo_erro = f"PDF suspeito ({tam} bytes): {destino}"
            log(ultimo_erro)
            destino.unlink(missing_ok=True)
            continue
        if pdf_inteiro(destino, item):
            log(f"PDF salvo: {destino.name} ({tam} bytes)")
            return destino
        ultimo_erro = f"PDF não conferiu páginas com a API ({destino.name})"
        log(f"{ultimo_erro}. Tentativa {tentativa}/{TENTATIVAS_PDF}.")
        destino.unlink(missing_ok=True)
    raise RuntimeError(ultimo_erro)


def extrair_fg(pdf: Path, pasta_agente: Path) -> tuple[Path | None, Path | None]:
    script = pasta_agente / "agente_fg_comissao.py"
    antes = {p.resolve() for p in pasta_agente.glob("FG_Cargos_Comissao*.xlsx")}
    log(f"Extraindo atos de {pdf.name}")
    proc = subprocess.run(
        [sys.executable, str(script), "--pdf", str(pdf)],
        cwd=str(pasta_agente),
        check=False,
    )
    if proc.returncode != 0:
        raise RuntimeError(f"agente_fg_comissao.py saiu com código {proc.returncode}")
    depois = [
        p
        for p in pasta_agente.glob("FG_Cargos_Comissao*.xlsx")
        if p.resolve() not in antes
    ]
    if not depois:
        log("Extração sem arquivo novo (0 atos). Não reutiliza extrato de outro dia.")
        return None, None
    gerado = max(depois, key=lambda p: p.stat().st_mtime)
    estavel = pasta_agente / "FG_Cargos_Comissao.xlsx"
    if gerado.resolve() != estavel.resolve():
        copiar_seguro(gerado, estavel)
    log(f"Extrato: {gerado.name} → {estavel.name}")
    return estavel, gerado


def atualizar_spgg(
    extrato: Path,
    pasta_atualizacao: Path,
    base_spgg: Path | None,
) -> Path:
    destino = pasta_atualizacao / "SPGG_PARA_ENVIAR_NUVEM.xlsx"
    if base_spgg and base_spgg.exists():
        if base_spgg.resolve() != destino.resolve():
            copiar_seguro(base_spgg, destino)
        log(f"Base SPGG (pasta anterior): {base_spgg}")
    if str(pasta_atualizacao) not in sys.path:
        sys.path.insert(0, str(pasta_atualizacao))
    import gerar_spgg_nuvem_limpo as nuvem
    import gerar_teste_atualizacao as g

    g.DOE_CANDIDATOS = [extrato]
    log(f"Atualizando SPGG com {extrato}")
    try:
        nuvem.incremental_lote(extrato)
    except OSError as exc:
        if arquivo_em_uso(exc):
            raise RuntimeError(
                f"Excel está com SPGG_PARA_ENVIAR_NUVEM.xlsx aberto em {pasta_atualizacao}. "
                "Feche o arquivo e rode de novo."
            ) from exc
        raise
    return nuvem.OUT


def processar_item(
    item: dict,
    pasta_agente: Path,
    pasta_atualizacao: Path | None,
    so_detectar: bool,
    so_baixar: bool,
    com_spgg: bool,
    base_spgg: Path | None = None,
) -> dict:
    nro = item.get("nro_edicao_pdf")
    extra_txt = f", nroEdicao {nro}" if nro else ""
    log(
        f"Edição {item['tipo']} nº {item['numero']} "
        f"de {item.get('data_br') or item['data']} "
        f"(id {item['id']}, download {item.get('id_download') or item['id']}"
        f"{extra_txt}, {item.get('paginas')} pág.)"
    )
    if so_detectar:
        return item
    pdf = baixar_pdf(item, pasta_agente / "pdfs")
    item["pdf"] = str(pdf)
    if so_baixar:
        return item
    estavel, datado = extrair_fg(pdf, pasta_agente)
    if estavel is None:
        item["extrato"] = ""
        item["extrato_datado"] = ""
        log("Sem atos FG neste DOE; SPGG não alterada.")
        return item
    item["extrato"] = str(estavel)
    item["extrato_datado"] = str(datado) if datado else ""
    if com_spgg:
        if not pasta_atualizacao:
            raise FileNotFoundError("Atualização da SPGG exige pasta de atualização.")
        saida = atualizar_spgg(estavel, pasta_atualizacao, base_spgg)
        item["spgg"] = str(saida)
        log(f"SPGG gerada: {saida}")
    else:
        log("SPGG não atualizada (--sem-spgg).")
    return item


def _resumo_linhas(processado: dict, args, base_spgg: Path | None, pasta_dia: Path) -> list[str]:
    return [
        f"ciclo: {datetime.now().isoformat(timespec='seconds')}",
        f"modo: {'DOE → FG → SPGG' if args.com_spgg else 'DOE → FG'}",
        f"edicao: {processado.get('tipo')} nº {processado.get('numero')}",
        f"data: {processado.get('data_br') or processado.get('data')}",
        f"id: {processado.get('id')}",
        f"nro_edicao_pdf: {processado.get('nro_edicao_pdf') or ''}",
        f"paginas: {processado.get('paginas')}",
        f"pdf: {processado.get('pdf', '')}",
        f"extrato: {processado.get('extrato', '')}",
        f"extrato_datado: {processado.get('extrato_datado', '')}",
        f"spgg: {processado.get('spgg', '')}",
        f"base_spgg: {base_spgg or ''}",
        f"pasta_resultados: {pasta_dia}",
    ]


def uma_passada(args) -> int:
    pasta_agente = achar_pasta_agente(args.pasta_agente)
    pasta_atualizacao = (
        achar_pasta_atualizacao(args.pasta_atualizacao) if args.com_spgg else None
    )
    estado = carregar_estado(AQUI)
    _, ultima_dia = achar_ultima_spgg()
    hoje = date.today()
    inicio = data_minima(args.a_partir_de, ultima_dia)

    log(f"Agente: {pasta_agente}")
    log(f"Modo: {'DOE → FG → SPGG' if args.com_spgg else 'DOE → FG (sem SPGG)'}")
    if ultima_dia:
        log(f"Última pasta com SPGG: {ultima_dia.isoformat()}")
        log(
            f"Varredura {inicio.isoformat()} → {hoje.isoformat()} "
            "(inclui o último dia, para extra posterior)"
        )
    else:
        log(f"Sem SPGG anterior em resultados; varredura {inicio.isoformat()} → {hoje.isoformat()}")
    if pasta_atualizacao:
        log(f"Atualização: {pasta_atualizacao}")

    try:
        aguardar_api()
        itens, sem_doe = montar_itens_periodo(inicio, hoje)
    except (urllib.error.URLError, TimeoutError, OSError) as exc:
        log(f"Falha ao consultar API do DOE: {exc}")
        return 2

    if not args.so_detectar:
        for dia in sem_doe:
            if dia < hoje:
                registrar_dia_sem_doe(dia)

    novos = []
    for item in itens:
        if args.forcar or not ja_processado(estado, item):
            novos.append(item)
        else:
            log(f"Já processado: {chave_edicao(item)}")

    if not novos:
        if not args.so_detectar and not pasta_tem_resultado(hoje) and hoje in sem_doe:
            log(
                f"DOE de {hoje.isoformat()} ainda não saiu. "
                f"Vigia consulta de novo em {INTERVALO_VIGIA_MIN} min."
            )
        else:
            log("Nenhuma edição nova.")
        return 0

    erros = 0
    for item in novos:
        dia = data_edicao(item) or hoje
        base_spgg, base_dia = achar_spgg_anterior(dia)
        if base_spgg:
            log(f"Base SPGG para {dia.isoformat()}: {base_dia} → {base_spgg}")
        try:
            processado = processar_item(
                item,
                pasta_agente,
                pasta_atualizacao,
                so_detectar=args.so_detectar,
                so_baixar=args.so_baixar,
                com_spgg=args.com_spgg,
                base_spgg=base_spgg,
            )
            if args.so_detectar:
                continue
            extrato_ok = bool(processado.get("extrato"))
            if extrato_ok:
                marcar_processado(estado, item)
                salvar_estado(AQUI, estado)
            else:
                log(
                    "Extração vazia: grava a pasta do dia e não marca sucesso definitivo "
                    "(o próximo ciclo tenta de novo)."
                )
            trazer_saidas(processado, pasta_atualizacao)
            pasta_dia = pasta_do_dia(processado)
            gravar_resumo(
                _resumo_linhas(processado, args, base_spgg, pasta_dia),
                pasta_dest=pasta_dia,
            )
        except (urllib.error.URLError, TimeoutError) as exc:
            log(f"Erro de rede na edição {item['id']}: {exc}")
            log("Interrompendo a varredura para não furar a cadeia da SPGG.")
            erros += 1
            break
        except Exception as exc:
            log(f"Erro na edição {item['id']}: {exc}")
            log("Interrompendo a varredura para não furar a cadeia da SPGG.")
            erros += 1
            break
    return 1 if erros else 0


def vigiar(args) -> int:
    """De 10 em 10 min só enquanto o dia não tiver pasta; com resultado, encerra."""
    intervalo = max(1, args.intervalo)
    log(
        f"Modo vigia: sem resultado de hoje, consulta a cada {intervalo} min. "
        f"Com a pasta gravada, encerra (próxima sessão no horário normal)."
    )
    while True:
        try:
            uma_passada(args)
        except KeyboardInterrupt:
            log("Encerrado.")
            return 0
        except Exception as exc:
            log(f"Falha inesperada no ciclo: {exc}")

        agora = datetime.now()
        hoje = agora.date()
        if not pasta_tem_resultado(hoje) and agora.hour >= HORA_SEM_DOE:
            try:
                item = edicao_do_dia(hoje.isoformat())
            except (urllib.error.URLError, TimeoutError, OSError, ValueError, json.JSONDecodeError):
                item = None
            if item is None:
                registrar_dia_sem_doe(hoje)

        if pasta_tem_resultado(hoje):
            log("Resultado do dia gravado. Vigia encerrada; próxima sessão no horário normal.")
            return 0

        if agora.hour >= HORA_FIM_VIGIA:
            registrar_dia_sem_doe(hoje)
            log("Fim da vigia do dia sem DOE. Próxima sessão: 07:00.")
            return 0

        log(f"Sem resultado de {hoje.isoformat()}. Nova consulta em {intervalo} min.")
        time.sleep(intervalo * 60)


def main() -> int:
    parser = argparse.ArgumentParser(description="Orquestrador autônomo DOE → FG → SPGG")
    parser.add_argument("--pasta-agente", type=Path, default=None)
    parser.add_argument("--pasta-atualizacao", type=Path, default=None)
    parser.add_argument(
        "--vigia",
        action="store_true",
        help="Consulta a cada 10 min só até gravar o resultado do dia; depois encerra",
    )
    parser.add_argument("--watch", action="store_true", help="Fica em loop")
    parser.add_argument("--intervalo", type=int, default=15, help="Minutos entre checagens")
    parser.add_argument(
        "--extras",
        action="store_true",
        help="Mantido por compatibilidade: o ciclo já usa a última atualização do dia.",
    )
    parser.add_argument("--forcar", action="store_true", help="Ignora estado_doe.json")
    parser.add_argument("--so-detectar", action="store_true", help="Só consulta a API")
    parser.add_argument("--so-baixar", action="store_true", help="Para depois do download do PDF")
    parser.add_argument(
        "--sem-spgg",
        action="store_true",
        help="Não atualiza a SPGG (só PDF + extração FG)",
    )
    parser.add_argument(
        "--a-partir-de",
        type=date.fromisoformat,
        default=DATA_INICIO_PADRAO,
        help="Ignora edições anteriores a esta data (YYYY-MM-DD)",
    )
    args = parser.parse_args()
    args.com_spgg = not args.sem_spgg

    if args.vigia:
        if args.intervalo == 15:
            args.intervalo = INTERVALO_VIGIA_MIN
        return vigiar(args)

    if not args.watch:
        return uma_passada(args)

    log(f"Modo watch a cada {args.intervalo} min")
    while True:
        try:
            uma_passada(args)
        except KeyboardInterrupt:
            log("Encerrado.")
            return 0
        except Exception as exc:
            log(f"Falha inesperada no ciclo: {exc}")
        time.sleep(max(1, args.intervalo) * 60)


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\nCancelado.")
        sys.exit(1)
