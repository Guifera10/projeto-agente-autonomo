#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Reprocessa DOE 30/09/2026–06/10/2026 (ordinária + extras) numa pasta só.

Base SPGG: resultados/2026-09-29/SPGG_PARA_ENVIAR_NUVEM.xlsx
Saída:     resultados/2026-10-06/  (para o orquestrador continuar depois desta data)
"""
from __future__ import annotations

import json
import shutil
import sys
from datetime import date, datetime, timedelta
from pathlib import Path

import orquestrador_doe as o

AQUI = Path(__file__).resolve().parent
PASTA_AGENTE = AQUI / "AgenteDOE"
PASTA_PDF = PASTA_AGENTE / "pdfs"
PASTA_ATUAL = AQUI / "Atualização DOE"
BASE_SPGG = AQUI / "resultados" / "2026-09-29" / "SPGG_PARA_ENVIAR_NUVEM.xlsx"
BASE_CHEC = AQUI / "resultados" / "2026-09-29" / "CHECAGEM_atualizacao_SPGG.xlsx"
INICIO = date(2026, 9, 30)
FIM = date(2026, 10, 6)


def edicoes_do_periodo() -> tuple[list[dict], list[dict]]:
    """Uma edição por dia (último extra, se houver) + lista para o estado."""
    dias: list[dict] = []
    estado_itens: list[dict] = []
    dia = INICIO
    while dia <= FIM:
        iso = dia.isoformat()
        try:
            detalhe = o.consultar_dia(iso)
        except Exception as exc:
            o.log(f"{iso}: sem edição ({exc})")
            dia += timedelta(days=1)
            continue
        if not detalhe:
            o.log(f"{iso}: sem edição")
            dia += timedelta(days=1)
            continue
        ordinaria = {
            "id": int(detalhe["id"]),
            "data": detalhe.get("dataString") or iso,
            "data_br": detalhe.get("dataPublicacao"),
            "numero": detalhe.get("numeroEdicao"),
            "tipo": "ORDINARIA",
            "paginas": detalhe.get("qtdPaginas"),
            "id_download": int(detalhe["id"]),
            "nro_edicao_pdf": None,
        }
        estado_itens.append(dict(ordinaria))
        extras = sorted(
            detalhe.get("edicoes") or [],
            key=lambda e: int(e.get("nroEdicao") or 0),
        )
        escolhida = ordinaria
        for extra in extras:
            nro = extra.get("nroEdicao")
            rec = {
                "id": int(extra["id"]),
                "data": iso,
                "data_br": extra.get("dataPublicacao") or detalhe.get("dataPublicacao"),
                "numero": nro,
                "tipo": f"EXTRA_{nro}",
                "paginas": extra.get("qtdPaginas"),
                "id_download": int(detalhe["id"]),
                "nro_edicao_pdf": int(nro) if nro is not None else None,
            }
            estado_itens.append(dict(rec))
            escolhida = rec
        dias.append(escolhida)
        dia += timedelta(days=1)
    return dias, estado_itens


def baixar_ultima(item: dict, pasta_pdf: Path) -> Path:
    """Baixa a última atualização do dia (nroEdicao do extra). Sempre rebaixa."""
    pasta_pdf.mkdir(parents=True, exist_ok=True)
    nro = item.get("nro_edicao_pdf")
    nome = f"DOE_{item['data']}_{item['tipo']}_{item['id_download']}.pdf"
    destino = pasta_pdf / nome
    url = f"{o.API}/public/diarios/download/{item['id_download']}/pdf"
    if nro:
        url += f"?nroEdicao={nro}"
    o.log(f"Baixando última atualização: {url}")
    o.http_download(url, destino)
    tam = destino.stat().st_size
    if tam < 10_000:
        raise RuntimeError(f"PDF suspeito ({tam} bytes): {destino}")
    o.log(f"PDF salvo: {destino.name} ({tam} bytes)")
    return destino


def main() -> int:
    if not BASE_SPGG.exists():
        raise FileNotFoundError(f"Base 29/09 não encontrada: {BASE_SPGG}")

    o.log("Consultando API (última atualização de cada dia, 30/09 a 06/10)...")
    itens, estado_itens = edicoes_do_periodo()
    for it in itens:
        o.log(
            f"  {it['data']} {it['tipo']} nº {it['numero']} "
            f"download id {it['id_download']} nroEdicao={it.get('nro_edicao_pdf')}"
        )
    if not itens:
        o.log("Nenhuma edição no período.")
        return 2

    pdfs: list[Path] = []
    for it in itens:
        pdf = baixar_ultima(it, PASTA_PDF)
        it["pdf"] = str(pdf)
        pdfs.append(pdf)

    o.log(f"Extraindo FG de {len(pdfs)} PDFs (última atualização de cada dia)...")
    script = PASTA_AGENTE / "agente_fg_comissao.py"
    antes = {p.resolve() for p in PASTA_AGENTE.glob("FG_Cargos_Comissao*.xlsx")}
    cmd = [sys.executable, str(script), "--pdf", *[str(p) for p in pdfs]]
    proc = __import__("subprocess").run(cmd, cwd=str(PASTA_AGENTE), check=False)
    if proc.returncode != 0:
        raise RuntimeError(f"agente_fg_comissao.py saiu com código {proc.returncode}")
    novos = [
        p
        for p in PASTA_AGENTE.glob("FG_Cargos_Comissao*.xlsx")
        if p.resolve() not in antes
    ]
    if not novos:
        raise RuntimeError("Extração não gerou Excel novo.")
    gerado = max(novos, key=lambda p: p.stat().st_mtime)
    estavel = PASTA_AGENTE / "FG_Cargos_Comissao.xlsx"
    if gerado.resolve() != estavel.resolve():
        shutil.copy2(gerado, estavel)
    o.log(f"Extrato combinado: {gerado.name}")

    dest_spgg = PASTA_ATUAL / "SPGG_PARA_ENVIAR_NUVEM.xlsx"
    dest_chec = PASTA_ATUAL / "CHECAGEM_atualizacao_SPGG.xlsx"
    shutil.copy2(BASE_SPGG, dest_spgg)
    if BASE_CHEC.exists():
        shutil.copy2(BASE_CHEC, dest_chec)
    o.log(f"Base SPGG restaurada de 29/09: {BASE_SPGG.name}")

    if str(PASTA_ATUAL) not in sys.path:
        sys.path.insert(0, str(PASTA_ATUAL))
    import gerar_spgg_nuvem_limpo as nuvem
    import gerar_teste_atualizacao as g

    g.DOE_CANDIDATOS = [estavel]
    o.log(f"Atualizando SPGG com {estavel.name}")
    nuvem.incremental_lote(estavel)
    o.log(f"SPGG gerada: {nuvem.OUT}")

    pasta_dia = AQUI / "resultados" / FIM.isoformat()
    pasta_dia.mkdir(parents=True, exist_ok=True)
    for pdf in pdfs:
        shutil.copy2(pdf, pasta_dia / pdf.name)
    shutil.copy2(estavel, pasta_dia / "FG_Cargos_Comissao.xlsx")
    shutil.copy2(gerado, pasta_dia / gerado.name)
    shutil.copy2(nuvem.OUT, pasta_dia / "SPGG_PARA_ENVIAR_NUVEM.xlsx")
    chec = PASTA_ATUAL / "CHECAGEM_atualizacao_SPGG.xlsx"
    if chec.exists():
        shutil.copy2(chec, pasta_dia / chec.name)

    estado = o.carregar_estado(AQUI)
    for it in estado_itens:
        o.marcar_processado(estado, it)
    estado["ultimo"]["pdf"] = str(pdfs[-1])
    estado["ultimo"]["extrato"] = str(estavel)
    o.salvar_estado(AQUI, estado)
    shutil.copy2(o.estado_path(), pasta_dia / "estado_doe.json")

    o.gravar_resumo(
        [
            f"ciclo: {datetime.now().isoformat(timespec='seconds')}",
            "modo: lote 30/09–06/10 (ordinária + extras) → FG → SPGG",
            f"base_spgg: {BASE_SPGG}",
            f"edicoes: {len(itens)}",
            "lista: " + "; ".join(f"{i['data']} {i['tipo']} id {i['id']}" for i in itens),
            f"extrato: {gerado}",
            f"spgg: {nuvem.OUT}",
            f"pasta_resultados: {pasta_dia}",
        ],
        pasta_dest=pasta_dia,
    )
    o.log(f"Pasta única de resultados: {pasta_dia}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
