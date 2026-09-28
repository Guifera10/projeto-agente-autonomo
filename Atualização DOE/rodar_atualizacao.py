#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Lançador offline: extrai o DOE (AgenteDOE) e atualiza a SPGG + checagem.

Uso:
  python rodar_atualizacao.py
  python rodar_atualizacao.py "FG_Cargos_Comissao.xlsx"
"""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

PASTA = Path(__file__).resolve().parent
AGENTE = PASTA.parent / "AgenteDOE"
if str(PASTA) not in sys.path:
    sys.path.insert(0, str(PASTA))


def pausar(mensagem="Pressione Enter para fechar..."):
    try:
        input(f"\n{mensagem}")
    except EOFError:
        pass


def verificar_dependencias(modulos) -> bool:
    faltando = []
    for modulo in modulos:
        try:
            __import__(modulo)
        except ImportError:
            faltando.append(modulo)
    if not faltando:
        return True
    print("Faltam bibliotecas Python:")
    for modulo in faltando:
        print(f"  - {modulo}")
    print("\nCom internet, instale uma vez:")
    print("  python -m pip install " + " ".join(faltando))
    print("\nDepois disso funciona sem internet.")
    return False


def listar_extratos() -> list[Path]:
    vistos = set()
    arquivos = []
    for pasta in (AGENTE, PASTA):
        if not pasta.is_dir():
            continue
        for arq in sorted(pasta.glob("FG_Cargos*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True):
            chave = (arq.name, arq.stat().st_size)
            if chave in vistos:
                continue
            vistos.add(chave)
            arquivos.append(arq)
    return arquivos


def escolher_extrato(arquivos: list[Path]) -> Path | None:
    print("Extratos FG_Cargos encontrados (mais recente primeiro):\n")
    for i, arq in enumerate(arquivos, 1):
        print(f"  [{i}] {arq}")
    print("\nDigite o número, o caminho completo, ou Enter para o [1].")
    print("Para cancelar: 0\n")
    escolha = input("Sua escolha: ").strip()
    if escolha in {"0", "sair", "cancelar"}:
        return None
    if not escolha:
        return arquivos[0]
    if escolha.isdigit():
        n = int(escolha)
        if 1 <= n <= len(arquivos):
            return arquivos[n - 1]
        print("Número inválido.")
        return None
    caminho = Path(escolha.strip('"'))
    if caminho.exists():
        return caminho
    print("Arquivo não encontrado:", caminho)
    return None


def atualizar_spgg(fonte: Path):
    import gerar_teste_atualizacao as g
    import gerar_spgg_nuvem_limpo as nuvem

    if not (PASTA / "TB_R_CARGOS_SPGG.xlsx").exists():
        print("Falta TB_R_CARGOS_SPGG.xlsx nesta pasta (cópia da original da nuvem).")
        return False
    g.DOE_CANDIDATOS = [fonte]
    print("\nFonte única:", fonte)
    print("Base SPGG:", PASTA / "TB_R_CARGOS_SPGG.xlsx")
    nuvem.main()
    print("\nPronto.")
    print("  Nuvem:   ", nuvem.OUT)
    print("  Checagem:", nuvem.OUT_CHECAGEM)
    return True


def extrair_pdfs():
    lancador = AGENTE / "rodar_agente.py"
    if not lancador.exists():
        print("Não achei", lancador)
        return
    subprocess.run([sys.executable, str(lancador)], check=False)


def menu():
    print("=" * 60)
    print("  Atualização SPGG × DOE  (offline)")
    print("=" * 60)
    print("  [1] Extrair atos dos PDFs (AgenteDOE)")
    print("  [2] Gerar SPGG da nuvem + checagem (a partir do FG_Cargos)")
    print("  [3] Os dois, nesta ordem")
    print("  [0] Sair")
    print()
    return input("Sua escolha: ").strip()


def main():
    if len(sys.argv) > 1:
        fonte = Path(sys.argv[1].strip('"'))
        if not fonte.exists():
            print("Arquivo não encontrado:", fonte)
            sys.exit(1)
        if not verificar_dependencias(("pandas", "openpyxl")):
            sys.exit(1)
        ok = atualizar_spgg(fonte)
        sys.exit(0 if ok else 1)

    while True:
        escolha = menu()
        if escolha in {"0", "sair"}:
            return
        if escolha == "1":
            extrair_pdfs()
            continue
        if escolha in {"2", "3"}:
            if escolha == "3":
                extrair_pdfs()
            if not verificar_dependencias(("pandas", "openpyxl")):
                pausar()
                return
            arquivos = listar_extratos()
            if not arquivos:
                print(f"Nenhum FG_Cargos*.xlsx em {AGENTE} nem em {PASTA}.")
                pausar()
                return
            fonte = escolher_extrato(arquivos)
            if fonte is None:
                continue
            atualizar_spgg(fonte)
            pausar()
            return
        print("Opção inválida.")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nCancelado.")
        sys.exit(1)
