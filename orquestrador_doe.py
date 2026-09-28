#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Orquestrador autônomo DOE → extração FG → SPGG.

Fluxo: consulta a API do DOE, baixa o PDF, extrai FG/cargos e aplica o
lote na SPGG. A base de cada atualização é sempre o último
`SPGG_PARA_ENVIAR_NUVEM.xlsx` gerado em `resultados/AAAA-MM-DD/`.
Só entram edições com data posterior à dessa pasta. O resultado do
ciclo vai para a pasta do dia novo.

Uso:
  python orquestrador_doe.py --so-detectar
  python orquestrador_doe.py --so-baixar
  python orquestrador_doe.py
  python orquestrador_doe.py --watch --intervalo 15
"""

from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
import time
import urllib.error
import urllib.request
from datetime import date, datetime, timedelta
from pathlib import Path

API = "https://doe-backend.pro.rs.gov.br"
TIPO_DOE = 1
UA = "OrquestradorDOE/1.0"

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
    dest.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(origem, dest)
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
    """Copia PDF, extrato FG, SPGG e checagem para resultados/AAAA-MM-DD/."""
    dest = pasta_do_dia(item)
    copiar_resultado(item.get("pdf"), pasta_dest=dest)
    extrato = item.get("extrato")
    if extrato:
        origem = Path(extrato)
        copiar_resultado(origem, "FG_Cargos_Comissao.xlsx", pasta_dest=dest)
        datados = sorted(
            origem.parent.glob("FG_Cargos_Comissao_*.xlsx"),
            key=lambda p: p.stat().st_mtime,
        )
        if datados:
            copiar_resultado(datados[-1], pasta_dest=dest)
    if item.get("spgg") and pasta_atualizacao:
        copiar_resultado(item.get("spgg"), "SPGG_PARA_ENVIAR_NUVEM.xlsx", pasta_dest=dest)
        copiar_resultado(
            pasta_atualizacao / "CHECAGEM_atualizacao_SPGG.xlsx",
            pasta_dest=dest,
        )
    copiar_resultado(estado_path(), "estado_doe.json", pasta_dest=dest)


def http_json(url: str, timeout: int = 45) -> dict | list:
    req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept": "application/json"})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return json.loads(resp.read().decode("utf-8"))


def http_download(url: str, destino: Path, timeout: int = 300) -> None:
    destino.parent.mkdir(parents=True, exist_ok=True)
    tmp = destino.with_suffix(destino.suffix + ".part")
    req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept": "application/pdf"})
    with urllib.request.urlopen(req, timeout=timeout) as resp, tmp.open("wb") as out:
        shutil.copyfileobj(resp, out)
    tmp.replace(destino)


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


def achar_ultima_spgg() -> tuple[Path | None, date | None]:
    """Último SPGG_PARA_ENVIAR_NUVEM.xlsx em resultados/AAAA-MM-DD/."""
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
        spgg = pasta / "SPGG_PARA_ENVIAR_NUVEM.xlsx"
        if not spgg.exists():
            continue
        if melhor_dia is None or dia > melhor_dia:
            melhor_dia = dia
            melhor_spgg = spgg
    return melhor_spgg, melhor_dia


def data_minima(piso: date, ultima_dia: date | None) -> date:
    """Próxima edição válida: posterior à última pasta de resultado."""
    if ultima_dia is None:
        return piso
    seguinte = ultima_dia + timedelta(days=1)
    return seguinte if seguinte > piso else piso


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
    }


def consultar_ultima() -> dict:
    return http_json(
        f"{API}/public/diarios/consultar/?tipoDiario={TIPO_DOE}&tipoRetorno=2"
    )


def consultar_dia(data_iso: str) -> dict:
    return http_json(
        f"{API}/public/diarios/consultar/?tipoDiario={TIPO_DOE}&data={data_iso}&tipoRetorno=2"
    )


def montar_itens(incluir_extras: bool) -> list[dict]:
    ultima = consultar_ultima()
    data_iso = ultima.get("dataString") or ""
    numero = ultima.get("numeroEdicao")
    itens = [
        {
            "id": int(ultima["id"]),
            "data": data_iso,
            "data_br": ultima.get("dataPublicacao"),
            "numero": numero,
            "tipo": "ORDINARIA",
            "paginas": ultima.get("qtdPaginas"),
        }
    ]
    if incluir_extras:
        detalhe = consultar_dia(data_iso) if data_iso else ultima
        for extra in detalhe.get("edicoes") or []:
            nro = extra.get("nroEdicao")
            itens.append(
                {
                    "id": int(extra["id"]),
                    "data": data_iso,
                    "data_br": extra.get("dataPublicacao") or ultima.get("dataPublicacao"),
                    "numero": nro,
                    "tipo": f"EXTRA_{nro}",
                    "paginas": extra.get("qtdPaginas"),
                }
            )
    return itens


def baixar_pdf(item: dict, pasta_pdf: Path) -> Path:
    pasta_pdf.mkdir(parents=True, exist_ok=True)
    nome = f"DOE_{item['data']}_{item['tipo']}_{item['id']}.pdf"
    destino = pasta_pdf / nome
    url = f"{API}/public/diarios/download/{item['id']}/pdf"
    if destino.exists() and destino.stat().st_size > 50_000:
        log(f"PDF já estava em disco: {destino.name}")
        return destino
    log(f"Baixando {url}")
    http_download(url, destino)
    tam = destino.stat().st_size
    if tam < 10_000:
        raise RuntimeError(f"PDF suspeito ({tam} bytes): {destino}")
    log(f"PDF salvo: {destino.name} ({tam} bytes)")
    return destino


def extrair_fg(pdf: Path, pasta_agente: Path) -> Path | None:
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
        return None
    gerado = max(depois, key=lambda p: p.stat().st_mtime)
    estavel = pasta_agente / "FG_Cargos_Comissao.xlsx"
    if gerado.resolve() != estavel.resolve():
        shutil.copy2(gerado, estavel)
    log(f"Extrato: {gerado.name} → {estavel.name}")
    return estavel


def atualizar_spgg(
    extrato: Path,
    pasta_atualizacao: Path,
    base_spgg: Path | None,
) -> Path:
    destino = pasta_atualizacao / "SPGG_PARA_ENVIAR_NUVEM.xlsx"
    if base_spgg and base_spgg.exists():
        if base_spgg.resolve() != destino.resolve():
            shutil.copy2(base_spgg, destino)
        log(f"Base SPGG (último gerado): {base_spgg}")
    if str(pasta_atualizacao) not in sys.path:
        sys.path.insert(0, str(pasta_atualizacao))
    import gerar_spgg_nuvem_limpo as nuvem
    import gerar_teste_atualizacao as g

    g.DOE_CANDIDATOS = [extrato]
    log(f"Atualizando SPGG com {extrato}")
    nuvem.incremental_lote(extrato)
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
    log(
        f"Edição {item['tipo']} nº {item['numero']} "
        f"de {item.get('data_br') or item['data']} "
        f"(id {item['id']}, {item.get('paginas')} pág.)"
    )
    if so_detectar:
        return item
    pdf = baixar_pdf(item, pasta_agente / "pdfs")
    item["pdf"] = str(pdf)
    if so_baixar:
        return item
    extrato = extrair_fg(pdf, pasta_agente)
    if extrato is None:
        item["extrato"] = ""
        log("Sem atos FG neste DOE; SPGG não alterada.")
        return item
    item["extrato"] = str(extrato)
    if com_spgg:
        if not pasta_atualizacao:
            raise FileNotFoundError("Atualização da SPGG exige pasta de atualização.")
        saida = atualizar_spgg(extrato, pasta_atualizacao, base_spgg)
        item["spgg"] = str(saida)
        log(f"SPGG gerada: {saida}")
    else:
        log("SPGG não atualizada (--sem-spgg).")
    return item


def uma_passada(args) -> int:
    pasta_agente = achar_pasta_agente(args.pasta_agente)
    pasta_atualizacao = (
        achar_pasta_atualizacao(args.pasta_atualizacao) if args.com_spgg else None
    )
    estado = carregar_estado(AQUI)
    base_spgg, ultima_dia = achar_ultima_spgg()
    piso = data_minima(args.a_partir_de, ultima_dia)

    log(f"Agente: {pasta_agente}")
    log(f"Modo: {'DOE → FG → SPGG' if args.com_spgg else 'DOE → FG (sem SPGG)'}")
    if base_spgg and ultima_dia:
        log(f"Última SPGG: {base_spgg}")
        log(f"Só edições posteriores a {ultima_dia.isoformat()} (a partir de {piso.isoformat()})")
    else:
        log(f"Sem SPGG anterior em resultados; vale a partir de {piso.isoformat()}")
    if pasta_atualizacao:
        log(f"Atualização: {pasta_atualizacao}")

    try:
        itens = montar_itens(args.extras)
    except Exception as exc:
        log(f"Falha ao consultar API do DOE: {exc}")
        return 2

    novos = []
    for item in itens:
        dia = data_edicao(item)
        if dia is not None and dia < piso:
            log(
                f"Ignorado (não posterior a última atualização; piso {piso.isoformat()}): "
                f"{chave_edicao(item)}"
            )
            continue
        if args.forcar or not ja_processado(estado, item):
            novos.append(item)
        else:
            log(f"Já processado: {chave_edicao(item)}")

    if not novos:
        log("Nenhuma edição nova.")
        return 0

    erros = 0
    for item in novos:
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
            if not args.so_detectar:
                marcar_processado(estado, item)
                salvar_estado(AQUI, estado)
            trazer_saidas(processado, pasta_atualizacao)
            pasta_dia = pasta_do_dia(processado)
            spgg_dia = pasta_dia / "SPGG_PARA_ENVIAR_NUVEM.xlsx"
            if spgg_dia.exists():
                base_spgg = spgg_dia
            gravar_resumo(
                [
                    f"ciclo: {datetime.now().isoformat(timespec='seconds')}",
                    f"modo: {'DOE → FG → SPGG' if args.com_spgg else 'DOE → FG'}",
                    f"edicao: {processado.get('tipo')} nº {processado.get('numero')}",
                    f"data: {processado.get('data_br') or processado.get('data')}",
                    f"id: {processado.get('id')}",
                    f"paginas: {processado.get('paginas')}",
                    f"pdf: {processado.get('pdf', '')}",
                    f"extrato: {processado.get('extrato', '')}",
                    f"spgg: {processado.get('spgg', '')}",
                    f"base_spgg: {base_spgg or ''}",
                    f"pasta_resultados: {pasta_dia}",
                ],
                pasta_dest=pasta_dia,
            )
        except (urllib.error.URLError, TimeoutError) as exc:
            log(f"Erro de rede na edição {item['id']}: {exc}")
            erros += 1
        except Exception as exc:
            log(f"Erro na edição {item['id']}: {exc}")
            erros += 1
    return 1 if erros else 0


def main() -> int:
    parser = argparse.ArgumentParser(description="Orquestrador autônomo DOE → FG → SPGG")
    parser.add_argument("--pasta-agente", type=Path, default=None)
    parser.add_argument("--pasta-atualizacao", type=Path, default=None)
    parser.add_argument("--watch", action="store_true", help="Fica em loop")
    parser.add_argument("--intervalo", type=int, default=15, help="Minutos entre checagens")
    parser.add_argument("--extras", action="store_true", help="Processa também extras do dia")
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
