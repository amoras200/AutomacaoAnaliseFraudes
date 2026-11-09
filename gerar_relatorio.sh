#!/usr/bin/env bash
# Gera o relatorio de chargeback. Sem argumentos, abre janelas para escolher a
# planilha e o periodo. Com argumentos, repassa ao script:
#   ./gerar_relatorio.sh planilha.xlsx --inicio 2026-03 --fim 2026-06
cd "$(dirname "$0")" || exit 1

esperar() { [ -t 0 ] && read -r -p "Pressione Enter para fechar..." _; }

if ! command -v python3 >/dev/null 2>&1; then
    echo "[ERRO] python3 nao encontrado. Instale o Python 3 (ex.: sudo apt install python3 python3-venv python3-tk)."
    esperar; exit 1
fi

if [ ! -d venv ]; then
    echo "Preparando o ambiente pela primeira vez (precisa de internet)..."
    python3 -m venv venv || {
        echo "[ERRO] Nao foi possivel criar o ambiente virtual. No Debian/Ubuntu: sudo apt install python3-venv"
        esperar; exit 1
    }
fi

# shellcheck disable=SC1091
source venv/bin/activate

if ! python -c "import pandas, openpyxl, matplotlib" >/dev/null 2>&1; then
    pip install -r requirements.txt || {
        echo "[ERRO] Falha ao instalar as bibliotecas. Verifique a internet."
        esperar; exit 1
    }
fi

python relatorio_chargeback.py "$@" || { esperar; exit 1; }
