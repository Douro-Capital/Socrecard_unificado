"""
orchestrador.py  —  Script Mãe  —  Douro Capital
=======================================================
Executa o pipeline de scorecard de crédito em três etapas sequenciais:

  1. Ranking CP         → tenta atualizar dados da ComDinheiro.
                          Se a cota da API estiver esgotada (ou qualquer
                          outra falha), emite um AVISO e continua com os
                          dados já existentes no Excel.
  2. Scorecard Bancos   → baixa dados do BCB e grava Watch List Bancos.xlsm
  3. Gerar Scorecard    → lê as planilhas e gera scorecard.html

Uso:
  python orchestrador.py
=======================================================
"""

import subprocess
import sys
import time
import os
from datetime import datetime

# ─────────────────────────────────────────────────────────────
# CONFIGURAÇÃO  — todos os scripts devem estar nesta mesma pasta
# ─────────────────────────────────────────────────────────────
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# Cada etapa: (label, arquivo, abort_on_failure)
# abort_on_failure=False → falha vira aviso e o pipeline continua
PIPELINE = [
    ("1º — Ranking CP",       "Ranking CP.py",        False),  # falha = aviso, continua
    ("2º — Scorecard Bancos", "Scorecard Bancos.py",  True),   # falha = pipeline para
    ("3º — Gerar Scorecard",  "gerar_scorecard 2.py", True),   # falha = pipeline para
]

# ─────────────────────────────────────────────────────────────
# HELPERS DE FORMATAÇÃO
# ─────────────────────────────────────────────────────────────

WIDTH = 62

def _line(char: str = "─") -> str:
    return char * WIDTH

def banner(text: str, char: str = "═") -> None:
    print(f"\n{_line(char)}")
    print(f"  {text}")
    print(_line(char))

def _fmt_time(seconds: float) -> str:
    if seconds < 60:
        return f"{seconds:.1f}s"
    m, s = divmod(int(seconds), 60)
    return f"{m}m {s:02d}s"


# ─────────────────────────────────────────────────────────────
# EXECUÇÃO DE CADA ETAPA
# ─────────────────────────────────────────────────────────────

def run_step(label: str, filename: str) -> float:
    """
    Executa um script Python como subprocesso isolado.
    Retorna o tempo de execução em segundos.
    Levanta RuntimeError se o script terminar com código != 0.
    """
    script_path = os.path.join(SCRIPT_DIR, filename)

    if not os.path.exists(script_path):
        raise FileNotFoundError(
            f"Arquivo não encontrado: {script_path}\n"
            f"  Verifique se o nome do script está correto."
        )

    banner(f"INICIANDO: {label}")
    print(f"  Arquivo : {filename}")
    print(f"  Início  : {datetime.now().strftime('%H:%M:%S')}\n")
    print(_line("·"))

    t0 = time.perf_counter()

    result = subprocess.run(
        [sys.executable, script_path],
        cwd=SCRIPT_DIR,
    )

    elapsed = time.perf_counter() - t0
    print(_line("·"))

    if result.returncode != 0:
        raise RuntimeError(
            f"'{label}' encerrou com código de saída {result.returncode}.\n"
            f"  Verifique os logs acima para identificar o erro."
        )

    print(f"\n  ✔  {label} concluído em {_fmt_time(elapsed)}")
    return elapsed


# ─────────────────────────────────────────────────────────────
# PIPELINE PRINCIPAL
# ─────────────────────────────────────────────────────────────

def main() -> None:
    banner("DOURO CAPITAL — PIPELINE DE SCORECARD DE CRÉDITO", "═")
    print(f"  Início   : {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    print(f"  Etapas   : {len(PIPELINE)}")
    print(f"  Diretório: {SCRIPT_DIR}")

    t_total = time.perf_counter()
    resultados: list[tuple[str, str, float]] = []

    for label, filename, abort_on_failure in PIPELINE:
        try:
            elapsed = run_step(label, filename)
            resultados.append((label, "✔  OK", elapsed))

        except FileNotFoundError as exc:
            elapsed = 0.0
            if abort_on_failure:
                resultados.append((label, "✘  ARQUIVO NÃO ENCONTRADO", elapsed))
                _abort(label, exc, resultados)
            else:
                resultados.append((label, "⚠  AVISO (continuou)", elapsed))
                _warn(label, exc)

        except RuntimeError as exc:
            elapsed = 0.0
            if abort_on_failure:
                resultados.append((label, "✘  FALHA NA EXECUÇÃO", elapsed))
                _abort(label, exc, resultados)
            else:
                resultados.append((label, "⚠  AVISO (continuou)", elapsed))
                _warn(label, exc)

        except Exception as exc:
            elapsed = 0.0
            if abort_on_failure:
                resultados.append((label, f"✘  ERRO: {type(exc).__name__}", elapsed))
                _abort(label, exc, resultados)
            else:
                resultados.append((label, "⚠  AVISO (continuou)", elapsed))
                _warn(label, exc)

    # ── Resumo final ─────────────────────────────────────────
    elapsed_total = time.perf_counter() - t_total
    banner("RESUMO DO PIPELINE", "═")

    for label, status, tempo in resultados:
        tempo_str = _fmt_time(tempo) if tempo > 0 else "—"
        print(f"  {status}  {label:<30}  ({tempo_str})")

    print(_line())
    print(f"  Tempo total : {_fmt_time(elapsed_total)}")
    print(f"  Fim         : {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    print(f"\n  Abra 'scorecard.html' no navegador para ver o resultado.\n")


def _warn(label: str, exc: Exception) -> None:
    """Exibe aviso de falha não-crítica e continua o pipeline."""
    print(f"\n  ⚠  {label} falhou mas o pipeline continuará com dados existentes.")
    print(f"     Motivo: {exc}\n")


def _abort(label: str, exc: Exception, resultados: list) -> None:
    """Exibe o resumo parcial e encerra o pipeline."""
    banner("PIPELINE INTERROMPIDO", "!")
    print(f"  Etapa  : {label}")
    print(f"  Motivo : {exc}")
    print()

    executadas = {r[0] for r in resultados}
    pendentes  = [lbl for lbl, _, _ in PIPELINE if lbl not in executadas]
    if pendentes:
        print("  Etapas NÃO executadas:")
        for p in pendentes:
            print(f"    · {p}")
    print()

    banner("RESUMO PARCIAL", "─")
    for lbl, status, tempo in resultados:
        tempo_str = _fmt_time(tempo) if tempo > 0 else "—"
        print(f"  {status}  {lbl:<30}  ({tempo_str})")
    print()

    sys.exit(1)


if __name__ == "__main__":
    main()
