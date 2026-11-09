#!/usr/bin/env python3
"""
Relatório de Chargeback: (1) valor por status, mês a mês; (2) frequência por
casa de show.

Uso (terminal, Linux ou Windows):
    python relatorio_chargeback.py planilha.xlsx
    python relatorio_chargeback.py planilha.xlsx --inicio 2026-03 --fim 2026-06
    python relatorio_chargeback.py planilha.xlsx --top 15
    python relatorio_chargeback.py planilha.xlsx --coluna-data "Data Transação"

Uso (sem terminal, ex.: duplo clique): abre janelas para escolher o arquivo
e informar o período.

Saída: pasta "relatorios/" ao lado da planilha, com um .xlsx (tabelas) e um
.pdf (gráfico por mês, gráfico por casa de show e tabela do período).
"""
import argparse
import sys
from datetime import date
from pathlib import Path

import matplotlib

matplotlib.use("Agg")  # sem janela; funciona em qualquer máquina
import matplotlib.pyplot as plt
import pandas as pd
from matplotlib.backends.backend_pdf import PdfPages

# ----------------------------------------------------------------------------
# Configuração: nomes das colunas e das abas. Se a planilha mudar, mexa aqui.
# ----------------------------------------------------------------------------
COL_VALOR = "Valor Pré-Chargeback (R$)"
COL_STATUS = "Status Do Pedido"
COL_LOCAL = "Local"
COL_DATA_PADRAO = "Data Recepção"  # mês em que o chargeback chegou
ABA_MODELO = "Total Geral - Chargeback.{ano}"

# Corrige variações de digitação encontradas na planilha
CORRECOES_STATUS = {
    "Fraude Recuperada Produtor": "Fraude - Recuperada Produtor",
    "Fraude Recuperada Ticket360": "Fraude - Recuperada Ticket360",
}
CORRECOES_LOCAL = {
    "Campo do Rive": "Campo do River",
    "istrito Anhembi": "Distrito Anhembi",
}

# Cores por significado: azul/verde = resolvido a favor, cinza = em andamento
# ou neutro, amarelo = atenção, laranja/vermelho = perda.
# A ORDEM abaixo também define a ordem de empilhamento, da legenda e da tabela
# (do "melhor" para o "pior").
CORES_STATUS = {
    "Fraude - Evitada":              "#1F77B4",  # azul
    "Fraude - Recuperada Ticket360": "#2E9E4F",  # verde
    "Fraude - Recuperada Produtor":  "#8FD694",  # verde claro
    "Fraude - Em Recuperação":       "#9AA0A6",  # cinza (em processamento)
    "Pedido Rejeitado":              "#5F6368",  # cinza escuro (neutro)
    "Pago":                          "#C9CCD1",  # cinza claro (neutro)
    "Cancelado Cliente":             "#F2C500",  # amarelo
    "Fraude - Tomada Produtor":      "#E8590C",  # laranja escuro (perda)
    "Fraude - Tomada Ticket360":     "#D62728",  # vermelho (maior preocupação)
}
COR_OUTROS = "#7F7F7F"  # status novo que ainda não está na lista acima
ORDEM_STATUS = list(CORES_STATUS)

MESES_PT = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun",
            "Jul", "Ago", "Set", "Out", "Nov", "Dez"]


def cor(status: str) -> str:
    return CORES_STATUS.get(status, COR_OUTROS)


def ordenar_status(status) -> list:
    """Ordena pela lista ORDEM_STATUS; status desconhecidos vão para o fim."""
    return sorted(status, key=lambda s: (ORDEM_STATUS.index(s)
                                         if s in ORDEM_STATUS else 99, s))


# ----------------------------------------------------------------------------
# Leitura e tratamento
# ----------------------------------------------------------------------------
def limpar_texto(s: pd.Series) -> pd.Series:
    """Tira espaços/tabs/quebras de linha sobrando."""
    return s.astype("string").str.replace(r"\s+", " ", regex=True).str.strip()


def ler_planilha(caminho: Path, inicio: pd.Period, fim: pd.Period,
                 col_data: str) -> pd.DataFrame:
    """Lê as abas dos anos do período e devolve só as colunas necessárias."""
    xls = pd.ExcelFile(caminho)
    partes = []
    for ano in range(inicio.year, fim.year + 1):
        aba = ABA_MODELO.format(ano=ano)
        if aba not in xls.sheet_names:
            print(f"Aviso: aba '{aba}' não existe na planilha; ignorando {ano}.")
            continue
        df = xls.parse(aba)
        df.columns = [str(c).strip() for c in df.columns]
        faltando = [c for c in (COL_VALOR, COL_STATUS, col_data)
                    if c not in df.columns]
        if faltando:
            sys.exit(f"Erro: aba '{aba}' não tem a(s) coluna(s): {faltando}")
        if COL_LOCAL not in df.columns:
            print(f"Aviso: aba '{aba}' não tem a coluna '{COL_LOCAL}'.")
            df[COL_LOCAL] = pd.NA
        partes.append(df[[col_data, COL_STATUS, COL_VALOR, COL_LOCAL]])

    if not partes:
        sys.exit("Erro: nenhuma aba do período foi encontrada.")

    df = pd.concat(partes, ignore_index=True)
    df.columns = ["data", "status", "valor", "local"]
    df["data"] = pd.to_datetime(df["data"], errors="coerce", dayfirst=True)
    df["valor"] = pd.to_numeric(df["valor"], errors="coerce")
    df["status"] = limpar_texto(df["status"]).replace(CORRECOES_STATUS)

    ruins = df["data"].isna() | df["valor"].isna() | df["status"].isna()
    if ruins.any():
        print(f"Aviso: {ruins.sum()} linha(s) ignoradas (data, valor ou status vazio/inválido).")
    df = df[~ruins].copy()

    # Local: limpa espaços, corrige typos e descarta valores que na verdade
    # são status digitados na coluna errada.
    df["local"] = limpar_texto(df["local"]).replace(CORRECOES_LOCAL)
    status_conhecidos = set(ORDEM_STATUS) | set(df["status"].unique())
    df.loc[df["local"].isin(status_conhecidos), "local"] = pd.NA

    df["mes"] = df["data"].dt.to_period("M")
    return df[(df["mes"] >= inicio) & (df["mes"] <= fim)]


def montar_tabela(df: pd.DataFrame, inicio: pd.Period, fim: pd.Period) -> pd.DataFrame:
    """Linhas = status, colunas = meses, valores = soma. Meses sem dado = 0."""
    tabela = df.pivot_table(index="status", columns="mes", values="valor",
                            aggfunc="sum", fill_value=0.0)
    todos_meses = pd.period_range(inicio, fim, freq="M")
    tabela = tabela.reindex(columns=todos_meses, fill_value=0.0)
    tabela = tabela.reindex(ordenar_status(tabela.index))
    tabela["Total"] = tabela.sum(axis=1)
    tabela.loc["TOTAL GERAL"] = tabela.sum()
    tabela.columns = [c if c == "Total" else f"{MESES_PT[c.month - 1]}/{c.year}"
                      for c in tabela.columns]
    tabela.index.name = "Status do Pedido"
    return tabela


def montar_tabela_casas(df: pd.DataFrame) -> pd.DataFrame:
    """Linhas = casas de show; colunas = nº de ocorrências por status,
    total de ocorrências e valor total (R$). Ordenada por frequência."""
    d = df.dropna(subset=["local"])
    sem_local = len(df) - len(d)
    if sem_local:
        print(f"Aviso: {sem_local} linha(s) sem 'Local' válido ficaram fora do gráfico de casas.")
    freq = d.pivot_table(index="local", columns="status", values="valor",
                         aggfunc="count", fill_value=0)
    freq = freq[ordenar_status(freq.columns)]
    freq["Total de ocorrências"] = freq.sum(axis=1)
    freq["Valor total (R$)"] = d.groupby("local")["valor"].sum()
    freq = freq.sort_values(["Total de ocorrências", "Valor total (R$)"],
                            ascending=False)
    freq.index.name = "Casa de show"
    return freq


# ----------------------------------------------------------------------------
# Saídas
# ----------------------------------------------------------------------------
def brl(v: float) -> str:
    return "R$ " + f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def brl_curto(v: float) -> str:
    if v >= 1_000_000:
        return f"R$ {v / 1_000_000:.1f} mi".replace(".", ",")
    if v >= 1_000:
        return f"R$ {v / 1_000:.1f} mil".replace(".", ",")
    return f"R$ {v:.0f}"


def salvar_excel(tabela: pd.DataFrame, casas: pd.DataFrame, destino: Path) -> None:
    with pd.ExcelWriter(destino, engine="openpyxl") as xw:
        tabela.to_excel(xw, sheet_name="Chargeback por status")
        ws = xw.sheets["Chargeback por status"]
        ws.column_dimensions["A"].width = 34
        for col in "BCDEFGHIJKLMNO":
            ws.column_dimensions[col].width = 15
        for row in ws.iter_rows(min_row=2, min_col=2):
            for c in row:
                c.number_format = '"R$" #,##0.00'

        casas.to_excel(xw, sheet_name="Frequência por casa")
        ws2 = xw.sheets["Frequência por casa"]
        ws2.column_dimensions["A"].width = 38
        for i in range(2, casas.shape[1] + 2):
            ws2.cell(row=1, column=i).alignment = ws2.cell(row=1, column=i).alignment.copy(wrap_text=True)
            ws2.column_dimensions[ws2.cell(row=1, column=i).column_letter].width = 18
        ult = casas.shape[1] + 1  # coluna do valor total
        for row in ws2.iter_rows(min_row=2, min_col=ult, max_col=ult):
            for c in row:
                c.number_format = '"R$" #,##0.00'


def grafico_mensal(tabela: pd.DataFrame, titulo: str):
    dados = tabela.drop(index="TOTAL GERAL")
    meses = [c for c in tabela.columns if c != "Total"]
    fig, ax = plt.subplots(figsize=(11.7, 8.3))  # A4 paisagem
    dados[meses].T.plot(kind="bar", stacked=True, ax=ax, width=0.75,
                        color=[cor(s) for s in dados.index])
    ax.set_title(titulo, fontsize=14, weight="bold")
    ax.set_ylabel("Valor (R$)")
    ax.set_xlabel("")
    ax.yaxis.set_major_formatter(lambda x, _: f"{x:,.0f}".replace(",", "."))
    plt.setp(ax.get_xticklabels(), rotation=45, ha="right")
    ax.legend(title="Status", fontsize=8, loc="upper left",
              bbox_to_anchor=(1.01, 1))
    fig.tight_layout()
    return fig


def grafico_casas(casas: pd.DataFrame, titulo: str, top: int):
    status_cols = [c for c in casas.columns
                   if c not in ("Total de ocorrências", "Valor total (R$)")]
    topo = casas.head(top).iloc[::-1]  # inverte p/ a maior ficar no topo
    fig, ax = plt.subplots(figsize=(11.7, 8.3))
    esquerda = pd.Series(0, index=topo.index)
    for s in status_cols:
        ax.barh(topo.index, topo[s], left=esquerda, color=cor(s), label=s,
                height=0.7)
        esquerda += topo[s]
    # rótulo no fim de cada barra: nº de ocorrências + valor total
    margem = topo["Total de ocorrências"].max() * 0.01
    for y, (n, v) in enumerate(zip(topo["Total de ocorrências"],
                                   topo["Valor total (R$)"])):
        ax.text(n + margem, y, f"{int(n)}  ·  {brl_curto(v)}",
                va="center", fontsize=8)
    ax.set_xlim(0, topo["Total de ocorrências"].max() * 1.18)
    ax.set_title(titulo, fontsize=14, weight="bold")
    ax.set_xlabel("Nº de chargebacks recebidos (rótulo: quantidade · valor total)")
    handles, labels = ax.get_legend_handles_labels()
    ax.legend(handles, labels, title="Status", fontsize=8, loc="lower right")
    ax.spines[["top", "right"]].set_visible(False)
    fig.tight_layout()
    return fig


def tabela_pdf(tabela: pd.DataFrame, titulo: str):
    fig, ax = plt.subplots(figsize=(11.7, 8.3))
    ax.axis("off")
    ax.set_title(titulo + " — valores", fontsize=14, weight="bold")
    texto = tabela.map(brl).reset_index()
    t = ax.table(cellText=texto.values, colLabels=texto.columns,
                 loc="upper center", cellLoc="right")
    t.auto_set_font_size(False)
    t.set_fontsize(6 if len(texto.columns) > 8 else 8)
    t.auto_set_column_width(range(len(texto.columns)))
    return fig


def salvar_pdf(tabela, casas, periodo: str, destino: Path, top: int) -> None:
    with PdfPages(destino) as pdf:
        for fig in (
            grafico_mensal(tabela, f"Chargeback por status — {periodo}"),
            grafico_casas(casas, f"Chargebacks por casa de show (top {top}) — {periodo}", top),
            tabela_pdf(tabela, f"Chargeback por status — {periodo}"),
        ):
            pdf.savefig(fig)
            plt.close(fig)


# ----------------------------------------------------------------------------
# Entrada: terminal ou janelas
# ----------------------------------------------------------------------------
def pedir_por_janelas():
    """Modo duplo-clique: escolhe arquivo e período em janelas simples."""
    import tkinter as tk
    from tkinter import filedialog, messagebox, simpledialog

    root = tk.Tk()
    root.withdraw()
    arq = filedialog.askopenfilename(
        title="Escolha a planilha de chargeback",
        filetypes=[("Excel", "*.xlsx *.xlsm")])
    if not arq:
        sys.exit(0)
    hoje = date.today()
    ini = simpledialog.askstring("Período", "Mês inicial (AAAA-MM):",
                                 initialvalue=f"{hoje.year}-01")
    fim = simpledialog.askstring("Período", "Mês final (AAAA-MM):",
                                 initialvalue=f"{hoje.year}-{hoje.month:02d}")
    if not ini or not fim:
        sys.exit(0)
    return arq, ini, fim, root, messagebox


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("planilha", nargs="?", help="caminho do .xlsx")
    ap.add_argument("--inicio", help="mês inicial AAAA-MM (padrão: janeiro do ano atual)")
    ap.add_argument("--fim", help="mês final AAAA-MM (padrão: mês atual)")
    ap.add_argument("--top", type=int, default=10,
                    help="quantas casas de show aparecem no gráfico (padrão: 10)")
    ap.add_argument("--coluna-data", default=COL_DATA_PADRAO,
                    help=f"coluna usada p/ definir o mês (padrão: '{COL_DATA_PADRAO}')")
    args = ap.parse_args()

    gui = args.planilha is None
    if gui:
        arq, ini, fim, root, msgbox = pedir_por_janelas()
    else:
        hoje = date.today()
        arq = args.planilha
        ini = args.inicio or f"{hoje.year}-01"
        fim = args.fim or f"{hoje.year}-{hoje.month:02d}"

    try:
        inicio, fim_p = pd.Period(ini, "M"), pd.Period(fim, "M")
        if fim_p < inicio:
            raise ValueError("o mês final é anterior ao inicial")
        caminho = Path(arq)
        df = ler_planilha(caminho, inicio, fim_p, args.coluna_data)
        if df.empty:
            raise ValueError("nenhum chargeback encontrado nesse período")

        tabela = montar_tabela(df, inicio, fim_p)
        casas = montar_tabela_casas(df)
        if casas.empty:
            raise ValueError("nenhuma linha com 'Local' válido nesse período")
        pasta = caminho.parent / "relatorios"
        pasta.mkdir(exist_ok=True)
        base = f"chargeback_{inicio}_{fim_p}"
        salvar_excel(tabela, casas, pasta / f"{base}.xlsx")
        salvar_pdf(tabela, casas, f"{inicio} a {fim_p}", pasta / f"{base}.pdf",
                   args.top)
    except Exception as e:  # noqa: BLE001 — mensagem amigável p/ usuário final
        if gui:
            msgbox.showerror("Erro", str(e))
            sys.exit(1)
        sys.exit(f"Erro: {e}")

    msg = f"Pronto! Arquivos salvos em:\n{pasta}"
    if gui:
        msgbox.showinfo("Concluído", msg)
    else:
        print(tabela.map(brl).to_string())
        print(f"\nTop {args.top} casas por frequência:")
        print(casas.head(args.top)[["Total de ocorrências", "Valor total (R$)"]]
              .assign(**{"Valor total (R$)": lambda x: x["Valor total (R$)"].map(brl)})
              .to_string())
        print("\n" + msg)


if __name__ == "__main__":
    main()
