#!/usr/bin/env python3
"""Lançador offline do Agente FG / Cargos em Comissão."""

import os
import subprocess
import sys
from pathlib import Path

PASTA = Path(__file__).resolve().parent
AGENTE = PASTA / "agente_fg_comissao.py"


def pausar(mensagem="Pressione Enter para fechar a janela..."):
    try:
        input(f"\n{mensagem}")
    except EOFError:
        pass


def verificar_dependencias() -> bool:
    faltando = []
    for modulo in ("pdfplumber", "openpyxl"):
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
    print("  python -m pip install pdfplumber openpyxl")
    print("\nDepois disso o agente funciona sem internet.")
    return False


def listar_pdfs():
    return sorted(set(PASTA.glob("*.pdf")) | set(PASTA.glob("*.PDF")))


def escolher_pdfs(pdfs):
    print("PDFs encontrados na pasta:\n")
    for i, pdf in enumerate(pdfs, 1):
        print(f"  [{i}] {pdf.name}")

    print("\nDigite o número do PDF (um, vários ou all)")
    print("Exemplos:  1   |   1 3 5   |   all")
    print("Para cancelar:  0  ou  sair\n")

    while True:
        escolha = input("Sua escolha: ").strip().lower()
        if escolha in {"0", "sair", "cancelar", "voltar", "exit", "quit"}:
            return []
        if escolha == "all":
            return list(pdfs)

        partes = escolha.replace(",", " ").split()
        if not partes:
            print("Digite um número, vários números ou all.")
            continue

        selecionados = []
        try:
            for parte in partes:
                num = int(parte)
                if 1 <= num <= len(pdfs):
                    selecionados.append(pdfs[num - 1])
                else:
                    print(f"  [AVISO] Número {num} inválido, ignorado.")
        except ValueError:
            print("Use apenas números ou all.")
            continue

        if selecionados:
            vistos = set()
            unicos = []
            for pdf in selecionados:
                if pdf not in vistos:
                    vistos.add(pdf)
                    unicos.append(pdf)
            return unicos
        print("Nenhum arquivo válido. Tente novamente.")


def main():
    os.chdir(PASTA)

    print("=" * 62)
    print("   AGENTE FG / CARGOS EM COMISSÃO")
    print("   Executador offline")
    print("=" * 62)
    print()
    print(f"Pasta: {PASTA}")
    print()

    if not AGENTE.exists():
        print("ERRO: agente_fg_comissao.py não está nesta pasta.")
        pausar()
        sys.exit(1)

    if not verificar_dependencias():
        pausar()
        sys.exit(1)

    pdfs = listar_pdfs()
    if not pdfs:
        print("Nenhum PDF encontrado nesta pasta.")
        print("Coloque o Diário Oficial (.pdf) junto destes scripts e rode de novo.")
        pausar()
        sys.exit(1)

    selecionados = escolher_pdfs(pdfs)
    if not selecionados:
        print("\nOperação cancelada.")
        pausar()
        sys.exit(0)

    print(f"\nProcessando {len(selecionados)} arquivo(s):", flush=True)
    for pdf in selecionados:
        print(f"  - {pdf.name}", flush=True)
    print("-" * 62, flush=True)

    comando = [sys.executable, str(AGENTE), "--pdf"] + [str(p) for p in selecionados]
    try:
        resultado = subprocess.run(comando, check=False)
    except Exception as e:
        print(f"\nERRO ao iniciar o agente: {e}")
        pausar()
        sys.exit(1)

    if resultado.returncode != 0:
        print("\nO agente terminou com erro.")
        pausar()
        sys.exit(resultado.returncode)

    print("\nConcluído. O Excel ficou nesta pasta (FG_Cargos_Comissao_....xlsx).")
    print("=" * 62)
    pausar()


if __name__ == "__main__":
    main()
