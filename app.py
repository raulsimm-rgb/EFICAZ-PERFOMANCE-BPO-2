# =============================================================================
#  PAINEL BPO FINANCEIRO  -  arquivo único
#
#  COMO ABRIR: dê dois cliques neste arquivo (ou, no terminal: python app.py).
#  Na primeira vez ele instala o que falta (Streamlit, pandas, openpyxl,
#  reportlab, xlrd, plotly) e
#  depois abre o painel no navegador. Deixe a janela preta aberta enquanto
#  usa o painel; fechar a janela encerra o painel.
#
#  Os dados ficam no arquivo "bpo.db", criado nesta mesma pasta. Faça backup dele.
# =============================================================================
import os
import sys


def _rodando_dentro_do_streamlit():
    try:
        from streamlit.runtime import exists
        return exists()
    except Exception:
        return False


def _abrir_painel():
    import importlib.util
    import re
    import subprocess
    import webbrowser
    from importlib.metadata import version

    print("=" * 60)
    print("  PAINEL BPO FINANCEIRO")
    print("=" * 60)

    # 1. Instala o que estiver faltando (só na primeira vez)
    instalar = [p for p in ("streamlit", "pandas", "openpyxl", "reportlab", "xlrd", "plotly") if importlib.util.find_spec(p) is None]
    if "streamlit" not in instalar:
        principal, secundaria = (int(x) for x in version("streamlit").split(".")[:2])
        if (principal, secundaria) < (1, 37):
            instalar.append("streamlit")  # versão antiga: atualiza
    if instalar:
        print(f"\nInstalando: {', '.join(instalar)} (só acontece uma vez, aguarde)...\n")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--upgrade", *instalar])

    # 2. Inicia o painel e abre o navegador
    print("\nIniciando o painel... o navegador vai abrir sozinho.")
    print("NÃO FECHE ESTA JANELA enquanto estiver usando o painel.\n")
    processo = subprocess.Popen(
        [sys.executable, "-m", "streamlit", "run", os.path.abspath(__file__),
         "--server.headless=true", "--browser.gatherUsageStats=false",
         # Tema com as cores da Eficaz Performance BPO (marinho + ciano)
         "--theme.base=light", "--theme.primaryColor=#102850", "--theme.backgroundColor=#F4F7FB",
         "--theme.secondaryBackgroundColor=#EAF0F7", "--theme.textColor=#0F1E36", "--theme.font=sans serif"],
        stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, encoding="utf-8", errors="replace",
        env={**os.environ, "PYTHONUNBUFFERED": "1"},
    )
    navegador_aberto = False
    for linha in processo.stdout:
        print(linha, end="")
        porta = re.search(r"URL:\s*https?://[^\s:/]+:(\d+)", linha)
        if porta and not navegador_aberto:
            endereco = f"http://localhost:{porta.group(1)}"
            webbrowser.open(endereco)
            navegador_aberto = True
            print(f"\n>>> Painel aberto em {endereco}  (se não abriu, copie esse endereço no navegador)\n")
    processo.wait()


if __name__ == "__main__" and not _rodando_dentro_do_streamlit():
    try:
        _abrir_painel()
    except KeyboardInterrupt:
        pass
    except Exception as erro:
        print(f"\nERRO: {erro}")
        print("Mande uma foto desta tela para o suporte.")
        input("\nPressione Enter para fechar...")
    sys.exit()


# =============================================================================
#  BIBLIOTECAS
# =============================================================================
import base64
import io
import json
import math
import re
import sqlite3
import traceback
import unicodedata
from datetime import datetime

import pandas as pd
import streamlit as st


# =============================================================================
#  IDENTIDADE VISUAL – EFICAZ PERFORMANCE BPO
#  Para trocar o logo, coloque um arquivo "logo.png" (ou .jpg) na mesma pasta
#  deste app: ele passa a ser usado na barra lateral e no PDF.
# =============================================================================
MARCA_NOME = "Eficaz Performance BPO"
COR_MARINHO = "#102850"          # cor primária / estrutural (do logo)
COR_MARINHO_ESCURO = "#0A1C3D"
COR_MARINHO_MEDIO = "#1E3A6E"
COR_CIANO = "#18D0E8"            # cor de destaque / acento (do logo)
COR_CIANO_ESCURO = "#0E9FB5"     # ciano legível sobre fundo claro (linhas, hover)
COR_CIANO_SUAVE = "#E4F7FA"
COR_FUNDO = "#F4F7FB"
COR_FUNDO_LATERAL = "#EAF0F7"
COR_TEXTO = "#0F1E36"
COR_TEXTO_SECUNDARIO = "#4A5A70"
COR_BORDA = "#D8E1EC"

LOGO_COMPLETO_B64 = (
    "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"
    "/wJsPZd7b38HAAAAAElFTkSuQmCC"
)
LOGO_ICONE_B64 = (
    "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAADAFBMVEUAAAAQKVIcOm4a1OgWt8wQobYOnrMo1+j9//8A//8ZMVZV"
    "VVVVqqoAf38AVVUhPW1///8pRG0nuMoAAP8wSGsAAFUkOVkkQW0yR2Z/f38hNlktSHAdM1U8VHU6U3Y41eFs5+gzSGYspbRQaIZH"
    "2eIkorKN+/cAqqpR2+MAPz9CVnLX//9Nx9IoO1kYLU9GWnVQYnpB1+IYw9czuspZdZA6u8lJxc9Z3uQ2qLVjdYxphZYAAD07VndV"
    "qv9Ru8Rnx86MprG0//8+1uA9rLlWa4lFsLtb4eZ3lapq1txl5+p49PIwQl0rQF01UXNIY4VV//9v3eSJ5usAAH8AVaoAf/84VYFC"
    "u8hXx8xb1NnJ4+oQrcMsPVkzRF4pQF41SGksydhIYX1LtLxFt8NofpR6jKZ96uqUmbIeMlA/Pz85OXIsPmE/f389wc421N5AT2lN"
    "U3NBsbtDws9b0dVhZn1/f/94prGXt8ezs8+pxsa80ujf3/8AJEgAmcwaz+Ae4PIkJEgqP2EmPGIlPmk/WYA9WoEzZpk/v78/ucYz"
    "zMwy099CVnZGXoFBXodVVapfdIpJr79Aq7VAvstG1t5O1N5haX9gbolieZGImZmRtracxMSD9/e2kZGqqr+jtrayzOXGxsb/v/8A"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAxjZUYAAABAHRSTlMA/v7+"
    "/v7++QMB8wMDAgP8As72AZADsPFwAs+w1FJuzzBQ0DCQ7hcDbwRJCFKT1TUsq/7QJbN0VKsbGAQ5A00wEA24lRxvURQvRiNwko1J"
    "AysXAgMCO4w5Sgn+clqzPPVFV3IoEx0QvAQEiwSs0UgOhpA9FwIPEAoNCwgHBf7/B3i432mEBQSZBeVwQYQDGGSDo5esIiVBDw4N"
    "IQcMDgoJBAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA1Flh2QAAB8dJREFUeNq1WedfU0kUvYQEXoIhpJAGCYYYSqiCILCogCBKUVFXxd7X"
    "tbu67rq999777p+6d+aVqe8lodwPyI+8d87MvefcuRMBaorgsRbYxgjuwB/h7aMg8Gd3AnRvD0VLEIxDHY2NuwiFsfXwpwCOzDXW"
    "19cjxRLAseBWrx527ibw9SZFFyZs6yiC+xB+lw1vUvT0bRkFkc4fAjyl6OgxzLJvweq7FHiTYu/mKQj8bE+jDt+k2KQtcHlGT4cL"
    "PKWYO7IJW+B7xl4veEqxe6O2aAkDHFLh1XRRW3QHN+ArJnx+vWrKTFvsq4XC6NbDz2FRoUstOtqiFuepvmKyDLZg4W/qPqTOq9ZX"
    "GuE33rcRgi7b6zhkQLylGvjZns/rPXOQcaWAChQmvK6KN8UqEokdn2vU2sLDeS7CpyeA1Z6DzU3eKvawRUscYO99l1cy9JF4M/nZ"
    "xGzuthyNLUjHP7Jbv+lT1qZvAywm8gZP0dejc55es25lC1rwiLqYT/l80aOkzrziNLaYlcDjsKTVNuvHCN+Xj/pIRNqKAK0Z3jON"
    "Wlvw1T1z+R31mVlnpwQtF434rDApwo7rl35QVzc2C0GuLff7R95WhW9JJ9MKUI76+Igk9gA0x11tMXKu1CcR+P0j+jqFEX64zSdH"
    "KjHpVJu2Rs4WI+cCgU6VwKGgSrPggwhfVOEpRX6GUWC+bM2O+AMBFwJKwXsl3guwJxHxuUSU0yzaiNrChEcCQ0vg919G4b91zPw7"
    "+mpx1RWelCKaMxzNfoYV/9EfaKhA4I+Ns/zPpFM+74hEy1YHxNUcmGiw4F1TJFAY+Q8UvESbvKMU0WzmJIEPcCHtYL9fiNFx3Pn1"
    "qFY5R6MKxeqr2FwPlnj8QIknQKPFRIbYaD+xlrh6slSsupETKVJt09B1WoQPnBN2gK3i8JhMMYYUed69X5ruNZtGhMEPgbEgwQdK"
    "C2sgt7tXdBT/5VNWLa84LY5Q/J6PUIpUNAfwbWfA1o75T+nRTbVdnwDoH1UojsOTRMpSYxN7Gn+dSpDOivD7CbwjHvyldPCwtl2T"
    "A6R/1C9RnF6CoURahLeOnqlE3oZv4LIzcQDgpH6C+QQpxqVNjMROY1MN3lXP74/xx5kJCs8RdO5Hf+pPZdIWikQPMXkX45oTyqC+"
    "KpnwDkPnVey7v8nb5dMaaRuGrk9litFxaaoisxEVvoDfuYDSzobaC5jRsMrwLJ+KUDPNw/ExhaLfoFMHOylN+AAjKC0YMH0xVIeR"
    "LWvGUSbt1OqfGs2OnrVGHjo8XSs5tTXxS9/1wXSSwmOEksMCehjmedOmUh/ijVvR7D1U38mnRGzXOh3gANkEKvMAzCdseEqxYgj5"
    "n5a6S3Rep9l7O8leLWXaqUd4oswyD08YRIIXpa6Te3iXaHZU1uzjJQ7equ3EGczBjofpdoGg3Z3AavG0lKotGhh8A/EtEf6JFvPw"
    "4DeRdSPgDilCsfZYorChzQ10XnXUhTb65r32kAsBq0HkK/zkZU7vh7+PqQQW/AJ3UyZGnWljBLyKmoUUpfLrzItkrhJauZ0iInze"
    "fq3EqFyOsuChIqF5konkLKNw4B+tcQ2kWYJXCHLyUc4mXKmVm/hiS0b4X9OSTLOi0Yop3VHemhFa+WWLQGrJuFVD0ija4KLUKtTx"
    "KtJGxs8gRzEeGyE1oPDNGWdxAAUZvi6bfk3pRkVlIImsIkVvnHW4tfFYjAo//JSDxwYqL18DT0bcIXXmSUyxalOKa1eBtVUCP52U"
    "Vx+6M0ldIUfmhDKkU80+4yi67S8fqfBfwG2r8GSw79XfMr/GNz7SDNF9jCKzz5YOOZX3rIQU/F8A3g3r8ckbRMvqEJ1TDn2N8K0Y"
    "XF6nn+rg7W4VUidc3hbmw6+ndfC0j6pDCI3n3BvqEM3dyuhsl87WuUd7QXNkSlbRUxBbIIuhCl+KZFk+MpUFva/awtZHTnk4NKBS"
    "GEE+pQXZibiCo23aOX1Io8xJuHVe/qshFPiSlEPz05zGFqp0VjB3KPBJcReDIkFBMPo/tJBNdJiRKeTVJ4ew+nGIk8vuQFUEUXai"
    "IcV6PuVBECLFtMSVeRMvRaFqdkAVwBR5Q+yzPHw2x55EaxjnB6tKEVKQrFpdVL4pc/AFZihiwsKgUAPwIEBdDNPCOZdxJihWqecO"
    "PNlE+SXJbJ4EhGLRaSnhVu60sD+eEj4eVrQr7eCSasZQ+m8Hg2SgbE7HtvDtlhzXwktnMjG/pre0c19FNNnXV1KiYapMu7bDA5pe"
    "MVBUupGuv/Aysa6vpvCZdG7o4KXp3Xr9X10DTl4RuujMF2VBmevLOvgh0F1xyB1KQ2FbFZxvNNlt2RCVydoYhN2OtPk7IX0+7ZLe"
    "bvKaVlA8Fww3eEsRRU1Oec06ypSFT2P5Df0VU6AYSuo0+xd70+2huuUZ/XGsDkg6XTuaRYdf1yrzSRXwbFrT2CJnfbE1uaKDZ82l"
    "cjS7HLtEszBz3kX4ttSqCu3ETGbmYnpQL/zWcI3/D4W7uJWuqybaH4hDUy0Uk4lQRfgL4CF876DfvyS94SsKvwKFm+KtSN+qUpne"
    "mi27UAzccB3Ua6XQ2WLgJ3b92WToNFuz8CtSCON68tIGhF9Rs4s2RfbBxpVZwRYDSDFY2BZ469L388oFY1PCryyoTQvfOzK9tcL/"
    "D9XgBSuNQFxjAAAAAElFTkSuQmCC"
)


def logo_bytes(versao: str = "completo") -> bytes:
    """PNG do logo. Um 'logo.png'/'logo.jpg' na pasta do app tem prioridade sobre o embutido."""
    if versao == "completo":
        pasta = os.path.dirname(os.path.abspath(__file__))
        for nome in ("logo.png", "logo.jpg", "logo.jpeg"):
            caminho = os.path.join(pasta, nome)
            if os.path.exists(caminho):
                try:
                    with open(caminho, "rb") as arquivo:
                        return arquivo.read()
                except OSError:
                    break
    return base64.b64decode(LOGO_COMPLETO_B64 if versao == "completo" else LOGO_ICONE_B64)


def logo_data_uri(versao: str = "completo") -> str:
    dados = logo_bytes(versao)
    tipo = "image/jpeg" if dados[:3] == b"\xff\xd8\xff" else "image/png"
    return f"data:{tipo};base64,{base64.b64encode(dados).decode()}"


# =============================================================================
#  1. CÁLCULOS (DRE, caixa, ponto de equilíbrio, formatação em R$)
# =============================================================================
# Linhas para onde cada categoria do ERP pode ser classificada.
# A ordem aqui é a ordem em que aparecem nas listas de escolha.
LINHAS = {
    "receita_bruta": "Receita Operacional Bruta",
    "impostos_vendas": "Deduções / Impostos sobre Vendas",
    "custos_variaveis": "CMV / CSV (custo da mercadoria ou do serviço)",
    "despesas_variaveis": "Despesas Comerciais Variáveis (comissões, fretes de venda, taxas de cartão)",
    "custos_fixos": "Custos e Despesas Fixas",
    "depreciacao": "Depreciação e Amortização",
    "receitas_financeiras": "Receitas Financeiras",
    "despesas_financeiras": "Despesas Financeiras (juros, tarifas, multas)",
    "impostos_lucro": "Impostos sobre o Lucro (IR/CSLL)",
    "amortizacao_emprestimos": "Amortização de Empréstimos (principal) – só caixa",
    "investimentos": "Investimentos / Compra de Ativos – só caixa",
    "distribuicao_lucros": "Distribuição de Lucros / Retiradas – só caixa",
    "ignorar": "Ignorar (transferência entre contas, aporte etc.)",
}
LINHAS_POR_ROTULO = {rotulo: chave for chave, rotulo in LINHAS.items()}
LINHAS_VALOR = [chave for chave in LINHAS if chave != "ignorar"]


# Linhas de natureza credora (somam). Todas as outras são devedoras (subtraem do resultado/caixa).
LINHAS_CREDORAS = ("receita_bruta", "receitas_financeiras")


def _numero_seguro(x) -> float:
    """Converte para float; vazio, NaN ou infinito viram 0 (NaN passaria por 'x or 0')."""
    try:
        n = float(x)
    except (TypeError, ValueError):
        return 0.0
    return n if math.isfinite(n) else 0.0


def calcular(valores: dict, funcionarios: int = 0) -> dict:
    """Recebe os totais por linha (chaves de LINHAS) e devolve a DRE completa.

    Cada valor vem na natureza da linha: receitas positivas = entrada; custos e despesas
    positivos = saída. Valor negativo = saldo invertido (ex.: estornos maiores que a despesa).
    Todos os percentuais usam a Receita Bruta como base.
    """
    v = {chave: _numero_seguro(valores.get(chave, 0)) for chave in LINHAS_VALOR}
    r = dict(v)
    r["receita_liquida"] = v["receita_bruta"] - v["impostos_vendas"]
    # Lucro Bruto: só o custo direto (mercadoria/serviço). Margem de Contribuição: também as
    # despesas que nascem de cada venda (comissões, fretes de venda, taxas de cartão).
    r["lucro_bruto"] = r["receita_liquida"] - v["custos_variaveis"]
    r["margem_contribuicao"] = r["lucro_bruto"] - v["despesas_variaveis"]
    r["ebitda"] = r["margem_contribuicao"] - v["custos_fixos"]
    r["ebit"] = r["ebitda"] - v["depreciacao"]
    r["resultado_financeiro"] = v["receitas_financeiras"] - v["despesas_financeiras"]
    r["lair"] = r["ebit"] + r["resultado_financeiro"]
    r["lucro_liquido"] = r["lair"] - v["impostos_lucro"]
    r["geracao_caixa"] = (
        r["lucro_liquido"] + v["depreciacao"]
        - v["amortizacao_emprestimos"] - v["investimentos"] - v["distribuicao_lucros"]
    )

    bruta = v["receita_bruta"]
    r["perc_mc"] = (r["margem_contribuicao"] / bruta * 100) if bruta > 0 else 0.0
    r["margem_bruta"] = (r["lucro_bruto"] / bruta * 100) if bruta > 0 else 0.0
    r["margem_liquida"] = (r["lucro_liquido"] / bruta * 100) if bruta > 0 else 0.0
    # Ponto de equilíbrio em faturamento bruto (comparável à Receita Bruta).
    # Só existe com faturamento, margem de contribuição positiva e custos fixos positivos.
    if bruta > 0 and r["perc_mc"] > 0 and v["custos_fixos"] > 0:
        r["peq"], r["peq_situacao"] = v["custos_fixos"] / (r["perc_mc"] / 100), "ok"
    elif bruta <= 0:
        r["peq"], r["peq_situacao"] = 0.0, "sem_faturamento"
    elif r["perc_mc"] <= 0:
        r["peq"], r["peq_situacao"] = 0.0, "margem_negativa"
    else:
        r["peq"], r["peq_situacao"] = 0.0, "sem_custo_fixo"
    funcionarios = int(_numero_seguro(funcionarios))
    r["funcionarios"] = funcionarios
    r["receita_por_funcionario"] = (bruta / funcionarios) if funcionarios > 0 else 0.0
    return r


def verificar_cascata(r: dict) -> list:
    """Confere cada degrau da DRE e do caixa. Devolve a lista de inconsistências (vazia = tudo certo)."""
    regras = [
        ("Receita Líquida", r["receita_liquida"], r["receita_bruta"] - r["impostos_vendas"]),
        ("Lucro Bruto", r["lucro_bruto"], r["receita_liquida"] - r["custos_variaveis"]),
        ("Margem de Contribuição", r["margem_contribuicao"], r["lucro_bruto"] - r["despesas_variaveis"]),
        ("EBITDA", r["ebitda"], r["margem_contribuicao"] - r["custos_fixos"]),
        ("EBIT", r["ebit"], r["ebitda"] - r["depreciacao"]),
        ("Resultado Financeiro", r["resultado_financeiro"], r["receitas_financeiras"] - r["despesas_financeiras"]),
        ("LAIR", r["lair"], r["ebit"] + r["resultado_financeiro"]),
        ("Lucro Líquido", r["lucro_liquido"], r["lair"] - r["impostos_lucro"]),
        ("Geração de Caixa", r["geracao_caixa"], r["lucro_liquido"] + r["depreciacao"]
         - r["amortizacao_emprestimos"] - r["investimentos"] - r["distribuicao_lucros"]),
        # Visão direta: todas as entradas menos todas as saídas (exceto depreciação, que não é caixa)
        ("Geração de Caixa (soma direta)", r["geracao_caixa"],
         sum(r[k] for k in LINHAS_CREDORAS)
         - sum(r[k] for k in LINHAS_VALOR if k not in LINHAS_CREDORAS and k != "depreciacao")),
    ]
    problemas = [f"{nome}: {brl(atual)} ≠ {brl(esperado)}" for nome, atual, esperado in regras
                 if not math.isfinite(atual) or abs(atual - esperado) > 0.01]
    problemas += [f"{k}: valor inválido" for k, v in r.items() if isinstance(v, float) and not math.isfinite(v)]
    return problemas


def saldos_invertidos(r: dict) -> list:
    """Linhas com saldo contrário à sua natureza (ex.: despesa credora por estornos > lançamentos)."""
    return [k for k in LINHAS_VALOR if r.get(k, 0) < -0.005]


def sobre_bruta(valor: float, r: dict) -> float:
    bruta = r["receita_bruta"]
    return (valor / bruta * 100) if bruta > 0 else 0.0


# (rótulo, chave do resultado)
ESTRUTURA_DRE = [
    ("1. Receita Operacional Bruta", "receita_bruta"),
    ("(-) Deduções de Vendas e Impostos", "impostos_vendas"),
    ("(=) Receita Operacional Líquida", "receita_liquida"),
    ("(-) CMV / CSV (custo da mercadoria ou do serviço)", "custos_variaveis"),
    ("(=) Lucro Bruto", "lucro_bruto"),
    ("(-) Despesas Comerciais Variáveis (comissões, fretes, taxas de cartão)", "despesas_variaveis"),
    ("(=) Margem de Contribuição", "margem_contribuicao"),
    ("(-) Custos e Despesas Fixas", "custos_fixos"),
    ("(=) EBITDA (Resultado Operacional antes da Depreciação)", "ebitda"),
    ("(-) Depreciação e Amortização", "depreciacao"),
    ("(=) EBIT (Resultado Operacional)", "ebit"),
    ("(+/-) Resultado Financeiro (receitas - despesas financeiras)", "resultado_financeiro"),
    ("(=) Lucro Antes do IR/CSLL", "lair"),
    ("(-) Impostos sobre o Lucro (IR/CSLL)", "impostos_lucro"),
    ("(=) Lucro Líquido do Período", "lucro_liquido"),
]

ESTRUTURA_CAIXA = [
    ("Lucro Líquido do Período", "lucro_liquido"),
    ("(+) Depreciação e Amortização (não sai do caixa)", "depreciacao"),
    ("(-) Amortização de Empréstimos (principal)", "amortizacao_emprestimos"),
    ("(-) Investimentos / Compra de Ativos", "investimentos"),
    ("(-) Distribuição de Lucros / Retiradas", "distribuicao_lucros"),
    ("(=) Geração de Caixa do Mês", "geracao_caixa"),
]


def brl(valor: float) -> str:
    """1234.5 -> 'R$ 1.234,50'."""
    texto = f"{abs(valor):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"-R$ {texto}" if valor < 0 else f"R$ {texto}"


def pct(valor: float) -> str:
    """12.345 -> '12,3%'."""
    return f"{valor:.1f}%".replace(".", ",")


MESES = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"]


def nome_mes(competencia: str) -> str:
    """'2026-09' -> 'set/2026'."""
    ano, mes = competencia.split("-")
    return f"{MESES[int(mes) - 1]}/{ano}"

# =============================================================================
#  2. IMPORTAÇÃO DO RELATÓRIO DO ERP
#     Feita para aceitar o layout de qualquer ERP: acha a tabela no meio de
#     logotipos e filtros, adivinha as colunas por apelidos, entende número
#     brasileiro e americano, sinal, parênteses, D/C e coluna de tipo.
# =============================================================================
class ErroImportacao(Exception):
    """Erro com mensagem pronta para mostrar ao operador (sem tela vermelha)."""


def _sem_acento(texto) -> str:
    texto = unicodedata.normalize("NFKD", str(texto))
    return "".join(c for c in texto if not unicodedata.combining(c)).lower().strip()


def _tem_termo(texto_normalizado: str, termos) -> bool:
    """Procura termos no início de palavras: 'vl' acha 'Vl. Líquido', mas 'iss' não acha 'comissão'."""
    return any(re.search(r"(?<![a-z0-9])" + re.escape(t), texto_normalizado) for t in termos)


# Apelidos de cada coluna, em ordem de preferência. Termos de "evitar" descartam
# colunas parecidas que não servem (ex.: "Conta bancária" não é categoria).
APELIDOS = {
    "data": {
        "termos": ["data de compet", "competencia", "compet", "dt compet", "mes de compet", "data de pagamento",
                   "pagamento", "dt pagto", "pagto", "baixa", "liquidacao", "quitacao", "data do movimento",
                   "movimento", "vencimento", "venc", "emissao", "data de lancamento", "data", "dt", "posting date",
                   "transaction date", "date", "due date"],
        "evitar": ["usuario", "alteracao", "atualiz", "criacao", "cadastro", "inclusao", "tipo"],
    },
    "categoria": {
        "termos": ["categoria", "plano de conta", "plano conta", "classificacao", "classif", "natureza financeira",
                   "conta financeira", "conta gerencial", "conta contabil", "rubrica", "grupo", "centro de resultado",
                   "natureza", "category", "account", "conta"],
        "evitar": ["bancaria", "banco", "corrente", "agencia", "caixa/banco", "conta movimento", "cod"],
    },
    "valor": {
        "termos": ["valor liquido", "vl liquido", "valor pago", "valor recebido", "valor realizado", "valor do lancamento",
                   "valor", "vlr", "vl", "montante", "quantia", "importe", "amount", "value", "total", "r$"],
        "evitar": ["saldo", "acumulado", "desconto", "juros", "multa", "original", "previsto", "bruto", "%", "percent"],
    },
    "descricao": {
        "termos": ["descricao", "descri", "historico", "hist", "observacao", "observ", "complemento", "memo",
                   "favorecido", "fornecedor", "cliente", "razao social", "nome", "detalhe", "description", "memo", "payee"],
        "evitar": ["categoria"],
    },
    "tipo": {
        "termos": ["tipo de lancamento", "tipo do lancamento", "tipo", "d/c", "c/d", "debito/credito",
                   "credito/debito", "operacao", "sinal", "e/s", "entrada/saida", "natureza do lancamento", "movimento",
                   "type", "dr/cr"],
        "evitar": ["documento", "pagamento", "cobranca", "tipo de conta", "data"],
    },
    "entrada": {
        "termos": ["entrada", "entradas", "credito", "creditos", "recebimento", "recebido", "receita", "credit",
                   "inflow"],
        "evitar": ["data", "tipo", "saldo", "/"],
    },
    "saida": {
        "termos": ["saida", "saidas", "debito", "debitos", "pagamento", "pago", "despesa", "debit", "outflow"],
        "evitar": ["data", "tipo", "saldo", "/", "forma"],
    },
}
# Palavras que, juntas numa mesma linha, indicam o cabeçalho da tabela.
GRUPOS_CABECALHO = {
    "data": ["data", "dt", "venc", "compet", "emissao", "pagto", "baixa", "liquidacao", "date"],
    "valor": ["valor", "vlr", "vl", "montante", "total", "debito", "credito", "entrada", "saida", "r$", "quantia",
              "amount", "value", "debit", "credit"],
    "categoria": ["categoria", "plano de conta", "classifica", "natureza", "conta", "rubrica", "grupo", "category",
                  "account"],
    "descricao": ["descri", "hist", "observ", "complemento", "favorecido", "fornecedor", "cliente", "razao", "memo",
                  "payee"],
    "tipo": ["tipo", "d/c", "operacao", "sinal", "type"],
}


def adivinhar_coluna(colunas, papel, ja_usadas=()):
    """Escolhe a coluna mais provável para um papel (data, categoria, valor...)."""
    regra = APELIDOS[papel]
    for respeitar_evitar in (True, False):  # 2ª passada aceita, p. ex., "Valor Bruto" se for o único valor
        for termo in regra["termos"]:
            for coluna in colunas:
                if coluna in ja_usadas:
                    continue
                nome = _sem_acento(coluna)
                if _tem_termo(nome, [termo]) and not (respeitar_evitar and _tem_termo(nome, regra["evitar"])):
                    return coluna
    return None


def _pontuar_cabecalho(celulas) -> int:
    """Quantos grupos diferentes (data, valor, categoria...) aparecem na linha."""
    textos = [_sem_acento(c) for c in celulas if isinstance(c, str) and c.strip()]
    if len(textos) < 2:
        return 0
    return sum(any(_tem_termo(t, termos) for t in textos) for termos in GRUPOS_CABECALHO.values())


def _celulas_texto(celulas) -> int:
    return sum(isinstance(c, str) and bool(c.strip()) for c in celulas)


def _achar_cabecalho(bruto: pd.DataFrame, limite: int = 80):
    """Varre as primeiras linhas e devolve (posição da linha, pontuação).
    Ignora logotipos, títulos e filtros que os ERPs colocam antes da tabela."""
    melhor, melhor_chave, melhor_pontos = None, (0, 0), 0
    for i in range(min(len(bruto), limite)):
        celulas = bruto.iloc[i].tolist()
        pontos = _pontuar_cabecalho(celulas)
        # Empate: vence a linha com mais células de texto (linhas de filtro do relatório têm poucas).
        chave = (pontos, _celulas_texto(celulas))
        if pontos and chave > melhor_chave:
            melhor, melhor_chave, melhor_pontos = i, chave, pontos
    if melhor is None:
        # Sem palavras conhecidas: usa a primeira linha só de textos com 3+ células.
        for i in range(min(len(bruto), limite)):
            valores = [c for c in bruto.iloc[i].tolist() if pd.notna(c) and str(c).strip()]
            if len(valores) >= 3 and all(isinstance(c, str) for c in valores):
                return i, 0
        return 0, 0
    return melhor, melhor_pontos


def _decodificar(conteudo: bytes) -> str:
    for codificacao in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            return conteudo.decode(codificacao)
        except UnicodeDecodeError:
            continue
    raise ErroImportacao("Não consegui ler o texto do arquivo (codificação desconhecida). "
                         "Abra no Excel e salve como .xlsx.")


def _ler_texto(conteudo: bytes) -> pd.DataFrame:
    texto = _decodificar(conteudo)
    linhas = [l for l in texto.splitlines() if l.strip()]
    if not linhas:
        raise ErroImportacao("O arquivo está vazio.")
    # Separador: o que aparece no maior número de linhas (";" no padrão brasileiro).
    separador = max([";", "\t", "|", ","], key=lambda s: (sum(s in l for l in linhas), s == ";"))
    largura = max(l.count(separador) for l in linhas) + 1
    try:
        return pd.read_csv(io.StringIO(texto), header=None, names=range(largura), sep=separador,
                           dtype=str, skip_blank_lines=True, engine="python", quotechar='"')
    except Exception as erro:  # noqa: BLE001
        raise ErroImportacao(f"O arquivo de texto está fora do padrão e não pôde ser lido ({erro}). "
                             "Abra no Excel e salve como .xlsx.") from erro


def _ler_planilhas(nome: str, conteudo: bytes) -> dict:
    """Devolve {nome da aba: tabela bruta sem cabeçalho}."""
    if not conteudo:
        raise ErroImportacao("O arquivo está vazio.")
    inicio = conteudo[:200].lstrip().lower()
    if nome.endswith((".xls", ".xlsx", ".xlsm")) and inicio.startswith((b"<", b"<?xml", b"<html", b"<table")):
        raise ErroImportacao(
            "Este arquivo tem extensão de Excel, mas por dentro é uma página da internet (alguns ERPs exportam assim). "
            "Abra no Excel e use 'Salvar como' → 'Pasta de Trabalho do Excel (.xlsx)', depois envie de novo.")
    if nome.endswith((".xlsx", ".xlsm")):
        try:
            return pd.read_excel(io.BytesIO(conteudo), header=None, dtype=object, sheet_name=None)
        except Exception as erro:  # noqa: BLE001
            raise ErroImportacao(
                "Não consegui abrir este Excel. Ele pode estar corrompido ou protegido por senha. "
                "Abra no Excel, remova a senha se houver, e salve de novo como .xlsx.") from erro
    if nome.endswith(".xls"):
        try:
            return pd.read_excel(io.BytesIO(conteudo), header=None, dtype=object, sheet_name=None)
        except ImportError as erro:
            raise ErroImportacao(
                "Para ler arquivos .xls (Excel antigo) o painel precisa do complemento 'xlrd'. Feche a janela "
                "preta e abra o painel de novo, que ele instala sozinho. Ou salve o arquivo como .xlsx.") from erro
        except Exception as erro:  # noqa: BLE001
            raise ErroImportacao("Não consegui abrir este .xls. Abra no Excel e salve como .xlsx.") from erro
    if nome.endswith((".csv", ".txt")):
        return {"Arquivo": _ler_texto(conteudo)}
    raise ErroImportacao("Formato não suportado. Envie .xlsx, .xls ou .csv exportado do ERP.")


def _nomes_unicos(cabecalho) -> list:
    """Cabeçalho vazio vira 'Coluna N'; nomes repetidos ganham (2), (3)..."""
    nomes, vistos = [], {}
    for j, c in enumerate(cabecalho):
        if isinstance(c, float) and c.is_integer():
            c = int(c)
        nome = str(c).strip() if pd.notna(c) and str(c).strip() else f"Coluna {j + 1}"
        nome = re.sub(r"\s+", " ", nome)
        if nome in vistos:
            vistos[nome] += 1
            nome = f"{nome} ({vistos[nome]})"
        else:
            vistos[nome] = 1
        nomes.append(nome)
    return nomes


def ler_arquivo(nome: str, conteudo: bytes, aba: str = None, linha_cabecalho: int = None):
    """Lê o relatório do ERP. Devolve (tabela, info).
    info: aba usada, abas disponíveis, linha do cabeçalho (como no Excel) e se a detecção é confiável.
    linha_cabecalho (opcional) força a linha do cabeçalho, contada como no Excel (1 = primeira)."""
    nome = nome.lower()
    planilhas = _ler_planilhas(nome, conteudo)
    planilhas = {str(k): v.dropna(how="all").dropna(axis=1, how="all") for k, v in planilhas.items()}
    planilhas = {k: v for k, v in planilhas.items() if not v.empty}
    if not planilhas:
        raise ErroImportacao("O arquivo não tem nenhuma linha preenchida.")

    # Escolhe a aba com o cabeçalho mais convincente (ERPs às vezes exportam capa + dados).
    if aba is None or aba not in planilhas:
        aba = max(planilhas, key=lambda k: (_achar_cabecalho(planilhas[k])[1], len(planilhas[k])))
    bruto = planilhas[aba]

    if linha_cabecalho:
        posicoes = [i for i, idx in enumerate(bruto.index) if idx + 1 >= linha_cabecalho]
        if not posicoes:
            raise ErroImportacao(f"A aba '{aba}' não tem linha {linha_cabecalho}.")
        pos, pontos = posicoes[0], _pontuar_cabecalho(bruto.iloc[posicoes[0]].tolist())
    else:
        pos, pontos = _achar_cabecalho(bruto)

    tabela = bruto.iloc[pos + 1:].copy()
    tabela.columns = _nomes_unicos(bruto.iloc[pos].tolist())
    tabela = tabela.dropna(how="all")
    # Remove colunas totalmente vazias e repetições do cabeçalho no meio (quebra de página do relatório).
    tabela = tabela.loc[:, tabela.notna().any()]
    if tabela.empty:
        raise ErroImportacao("Achei o cabeçalho, mas nenhuma linha de dados abaixo dele.")
    primeira = tabela.columns[0]
    tabela = tabela[tabela[primeira].astype(str).str.strip() != primeira]
    linha_excel = int(bruto.index[pos]) + 1
    info = {"aba": aba, "abas": list(planilhas), "linha_cabecalho": linha_excel, "confiavel": pontos >= 2}
    return tabela.reset_index(drop=True), info


# ----------------------------------------------------------------- números
def _limpar_texto_numero(x) -> str:
    s = str(x).strip().replace(" ", " ").replace("−", "-").replace("–", "-")
    return s


def detectar_formato_numero(serie: pd.Series) -> str:
    """'br' (1.234,56) ou 'us' (1,234.56), olhando a coluna inteira."""
    br = us = 0
    for x in serie.dropna().head(1000):
        if isinstance(x, (int, float)):
            continue
        t = re.sub(r"[^\d,.]", "", _limpar_texto_numero(x))
        if "," in t and "." in t:
            if t.rfind(",") > t.rfind("."):
                br += 1
            else:
                us += 1
        elif re.search(r",\d{1,2}$", t):
            br += 1
        elif re.search(r"\.\d{1,2}$", t):
            us += 1
    return "us" if us > br else "br"


def _converter_numero(x, formato: str):
    if x is None or isinstance(x, bool):
        return None
    if isinstance(x, (int, float)):
        return None if pd.isna(x) else float(x)
    s = _limpar_texto_numero(x)
    if not s or s in ("-", "--", "R$", "R$ -"):
        return None
    maiusculo = s.upper()
    negativo = (
        s.startswith("-") or s.endswith("-") or "R$ -" in maiusculo or "R$-" in maiusculo
        or (s.startswith("(") and s.endswith(")"))
        or bool(re.search(r"(^|[\s\d])D$", maiusculo)) or maiusculo.startswith("D ")  # 500,00 D (débito)
    )
    t = re.sub(r"[^\d,.]", "", s)
    if not re.search(r"\d", t):
        return None
    milhar, decimal = (".", ",") if formato == "br" else (",", ".")
    if milhar in t and decimal in t:
        t = t.replace(milhar, "").replace(decimal, ".")
    elif decimal in t:
        if t.count(decimal) > 1:  # 1,234,567 num arquivo "br": na verdade é separador de milhar
            t = t.replace(decimal, "")
        else:
            t = t.replace(decimal, ".")
    elif milhar in t:
        if t.count(milhar) > 1 or re.fullmatch(r"\d{1,3}(" + re.escape(milhar) + r"\d{3})+", t):
            t = t.replace(milhar, "")
        else:  # ex.: "1234.5" num arquivo "br" — só pode ser decimal
            t = t.replace(milhar, ".")
    try:
        n = float(t)
    except ValueError:
        return None
    return -abs(n) if negativo else n


def para_numero(serie: pd.Series, formato: str = None) -> pd.Series:
    """Converte qualquer escrita de valor: -500 · (500) · 500- · R$ 1.234,56 · 1,234.56 · 500,00 D."""
    formato = formato or detectar_formato_numero(serie)
    return serie.map(lambda x: _converter_numero(x, formato)).astype(float)


# ------------------------------------------------------------------- datas
def _detectar_ordem_data(serie: pd.Series) -> bool:
    """True se as datas com barra vêm com o dia primeiro (padrão brasileiro)."""
    dia_primeiro = mes_primeiro = 0
    for x in serie.dropna().head(1000):
        m = re.match(r"\s*(\d{1,2})[/.-](\d{1,2})[/.-]\d{2,4}", str(x))
        if m:
            a, b = int(m[1]), int(m[2])
            dia_primeiro += a > 12
            mes_primeiro += b > 12
    return not (mes_primeiro > dia_primeiro)


def _converter_data(x, dia_primeiro: bool):
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return pd.NaT
    if isinstance(x, pd.Timestamp):
        return x
    if hasattr(x, "year"):  # datetime/date vindos do Excel
        return pd.Timestamp(x)
    if isinstance(x, (int, float)):
        n = float(x)
        if 20000 <= n <= 80000:  # número serial do Excel (1954 a 2119)
            return pd.Timestamp("1899-12-30") + pd.Timedelta(days=n)
        x = str(int(n)) if n.is_integer() else str(x)
    s = str(x).strip()
    if not s:
        return pd.NaT
    if re.fullmatch(r"\d{8}", s):  # 20260715
        return pd.to_datetime(s, format="%Y%m%d", errors="coerce")
    if re.fullmatch(r"(19|20)\d{2}(0[1-9]|1[0-2])", s):  # 202607 (ano e mês)
        return pd.Timestamp(f"{s[:4]}-{s[4:]}-01")
    if re.fullmatch(r"\d{1,2}[/.-]\d{4}", s) or re.fullmatch(r"\d{4}[/.-]\d{1,2}", s) \
            or re.fullmatch(r"[A-Za-zçÇ]{3,9}\.?[/ -]*\d{2,4}", s):  # só mês: 07/2026, 2026-07, jul/2026
        comp = _para_competencia(s)
        return pd.Timestamp(f"{comp}-01") if comp else pd.NaT
    if re.match(r"\d{4}-\d{2}-\d{2}", s):  # ISO
        return pd.to_datetime(s[:19], errors="coerce")
    return pd.to_datetime(s, dayfirst=dia_primeiro, errors="coerce")


def para_data(serie: pd.Series) -> pd.Series:
    dia_primeiro = _detectar_ordem_data(serie)
    return serie.map(lambda x: _converter_data(x, dia_primeiro))


# ------------------------------------------------------------ entra/sai
MODO_SINAL = "Valor com sinal: saídas negativas, como -500 ou (500)"
MODO_POSITIVO = "Todos os valores positivos: a categoria define"
MODO_TIPO = "Coluna de tipo: Receita/Despesa, D/C, Entrada/Saída"
MODO_DUAS = "Duas colunas de valor: Entradas e Saídas (ou Crédito e Débito)"
MODOS = [MODO_SINAL, MODO_TIPO, MODO_DUAS, MODO_POSITIVO]


def sentido_do_tipo(tipo):
    """+1 entrada, -1 saída, None se não reconhecer."""
    t = _sem_acento(tipo).strip(" .")
    if not t or t == "nan":
        return None
    if t in ("d", "s", "-", "db", "deb", "dr", "p", "sai", "cp", "pagar", "saida", "debito") or t.startswith(
            ("desp", "deb", "sai", "pag", "custo", "a pagar", "conta a pagar", "contas a pagar", "retir", "sangria",
             "estorno de receita", "compra")):
        return -1
    if t in ("c", "e", "+", "cr", "cred", "r", "ent", "cr.", "receber", "entrada", "credito") or t.startswith(
            ("rec", "cred", "entr", "a receber", "conta a receber", "contas a receber", "venda", "aporte",
             "suprimento", "estorno de despesa")):
        return 1
    if any(p in t for p in ("pagar", "despesa", "debito", "saida")):
        return -1
    if any(p in t for p in ("receber", "receita", "credito", "entrada")):
        return 1
    return None


def sugerir_modo(tabela, col_valor, col_tipo, col_entrada, col_saida) -> str:
    if col_entrada and col_saida and col_entrada != col_saida:
        return MODO_DUAS
    if col_valor is not None and col_valor in tabela:
        numeros = para_numero(tabela[col_valor])
        if (numeros < 0).any():
            return MODO_SINAL
        if col_tipo and col_tipo in tabela:
            preenchido = tabela[col_tipo].notna() & (tabela[col_tipo].astype(str).str.strip() != "")
            sentidos = tabela.loc[preenchido, col_tipo].map(sentido_do_tipo)
            if len(sentidos) and sentidos.notna().mean() >= 0.6 and (sentidos == -1).any():
                return MODO_TIPO
        return MODO_POSITIVO if (numeros > 0).any() else MODO_SINAL
    return MODO_SINAL


def _exemplos(serie, n=4) -> str:
    vistos = [str(v) for v in serie.dropna().astype(str).str.strip().unique()[:n] if v]
    return ", ".join(f"'{v}'" for v in vistos) if vistos else "(vazia)"


def preparar(tabela, col_data, col_categoria, col_valor=None, col_descricao=None, modo=MODO_SINAL,
             col_tipo=None, col_entrada=None, col_saida=None):
    """Padroniza a tabela do ERP. Devolve (lançamentos válidos, linhas descartadas com motivo, avisos).
    Levanta ErroImportacao com mensagem clara se a escolha de colunas não fizer sentido."""
    exigidas = {"Data": col_data, "Categoria": col_categoria}
    if modo == MODO_DUAS:
        exigidas.update({"Entradas / Crédito": col_entrada, "Saídas / Débito": col_saida})
    else:
        exigidas["Valor"] = col_valor
    if modo == MODO_TIPO:
        exigidas["Tipo"] = col_tipo
    faltando = [papel for papel, col in exigidas.items() if not col or col not in tabela.columns]
    if faltando:
        raise ErroImportacao(
            "Falta escolher a(s) coluna(s): " + ", ".join(faltando) + ". Escolha nas listas acima. Se o arquivo "
            "não tiver essa informação, troque o formato dos valores (passo 2) ou exporte do ERP um relatório "
            "que a inclua.")
    usadas = [c for c in exigidas.values()]
    repetidas = {c for c in usadas if usadas.count(c) > 1}
    if repetidas:
        raise ErroImportacao(f"A coluna '{', '.join(repetidas)}' foi escolhida para mais de uma função. "
                             "Cada informação (data, categoria, valor...) precisa de uma coluna diferente.")

    avisos = []
    datas = para_data(tabela[col_data])
    preenchidas = tabela[col_data].notna() & (tabela[col_data].astype(str).str.strip() != "")
    if preenchidas.sum() and datas[preenchidas].isna().mean() > 0.5:
        raise ErroImportacao(
            f"A coluna '{col_data}', escolhida como Data, não tem datas reconhecíveis. "
            f"Exemplos encontrados: {_exemplos(tabela.loc[preenchidas & datas.isna(), col_data])}. "
            "Escolha outra coluna em 'Data que define o mês'.")

    if modo == MODO_DUAS:
        entradas, saidas = para_numero(tabela[col_entrada]), para_numero(tabela[col_saida])
        valores = entradas.fillna(0).abs() - saidas.fillna(0).abs()
        valores[entradas.isna() & saidas.isna()] = float("nan")
        colunas_valor = [col_entrada, col_saida]
    else:
        valores = para_numero(tabela[col_valor])
        colunas_valor = [col_valor]
    for col in colunas_valor:
        tem_texto = tabela[col].notna() & (tabela[col].astype(str).str.strip() != "")
        numeros = para_numero(tabela[col])
        if tem_texto.sum() and numeros[tem_texto].isna().mean() > 0.5:
            raise ErroImportacao(
                f"A coluna '{col}', escolhida como valor, não tem números reconhecíveis. "
                f"Exemplos encontrados: {_exemplos(tabela.loc[tem_texto & numeros.isna(), col])}. "
                "Escolha a coluna certa de valor.")

    motivo = pd.Series("", index=tabela.index)
    if modo == MODO_POSITIVO:
        if (valores < 0).any():
            avisos.append(f"{int((valores < 0).sum())} valor(es) negativo(s) foram tratados como positivos, porque o "
                          "modo escolhido é 'Todos os valores positivos'. Se o arquivo usa sinal, troque o modo.")
        valores = valores.abs()
    elif modo == MODO_TIPO:
        sentidos = tabela[col_tipo].map(sentido_do_tipo)
        # Só reclama de linhas que seriam lançamentos (linhas de total não têm data)
        desconhecidos = sentidos.isna() & valores.notna() & datas.notna()
        if desconhecidos.any():
            nao_lidos = tabela.loc[desconhecidos, col_tipo].dropna().astype(str).str.strip().unique()
            avisos.append(
                f"{int(desconhecidos.sum())} linha(s) com tipo que não reconheci como entrada ou saída "
                f"({', '.join(repr(t) for t in nao_lidos[:6]) or 'vazio'}). Elas ficaram de fora; "
                "confira se a coluna de tipo está certa.")
            motivo[desconhecidos] = "Tipo não reconhecido como entrada ou saída"
        valores = valores.abs() * sentidos
    elif modo == MODO_SINAL and valores.notna().any() and not (valores < 0).any():
        avisos.append("Nenhum valor negativo no arquivo. Se as despesas vêm positivas, escolha o modo "
                      "'Coluna de tipo' ou 'Todos os valores positivos'; senão as despesas viram receita.")

    categorias = tabela[col_categoria].astype(str).str.strip()
    categorias = categorias.where(~categorias.isin(["", "nan", "None", "NaN", "-"]))
    descricoes = tabela[col_descricao].astype(str).str.strip().replace({"nan": ""}) if col_descricao else ""

    motivo[(motivo == "") & datas.isna()] = "Sem data válida"
    motivo[(motivo == "") & categorias.isna()] = "Sem categoria"
    motivo[(motivo == "") & valores.isna()] = "Sem valor numérico"

    df = pd.DataFrame({"data": datas, "categoria": categorias, "descricao": descricoes, "valor": valores})
    validas = motivo == ""
    descartadas = tabela[~validas].copy()
    descartadas.insert(0, "Motivo", motivo[~validas])
    df = df[validas & (df["valor"] != 0)].copy()
    if df.empty:
        raise ErroImportacao(
            "Nenhum lançamento válido com essas colunas. Confira na tabela de linhas ignoradas o motivo de cada uma "
            "e ajuste a escolha das colunas.") if len(descartadas) else ErroImportacao("O arquivo não tem lançamentos.")
    sem_categoria = int((motivo == "Sem categoria").sum())
    if sem_categoria > len(tabela) * 0.3:
        avisos.append(f"{sem_categoria} linhas estão sem categoria e ficaram de fora. Confira se a coluna de categoria "
                      "é a certa ou se o ERP deixou lançamentos sem classificar.")
    df["competencia"] = df["data"].dt.strftime("%Y-%m")
    df["data"] = df["data"].dt.strftime("%Y-%m-%d")
    return df.reset_index(drop=True), descartadas, avisos


# Termos das despesas que nascem de cada venda (vão para "Despesas Comerciais Variáveis").
TERMOS_DESPESAS_VARIAVEIS = ("comiss", "frete", "taxa de cartao", "taxas de cartao", "taxa cartao", "taxa do cartao",
                             "maquininha", "mdr", "adquirente", "gateway", "intermediac", "marketplace",
                             "taxa de venda", "royalt", "bonificac")

# Sugestão inicial de classificação. O BPO sempre revisa no Plano de Contas.
REGRAS_SUGESTAO = [
    ("ignorar", ("transferencia", "aporte", "entre contas", "resgate", "aplicacao financeira")),
    ("impostos_lucro", ("irpj", "csll", "imposto de renda")),
    ("impostos_vendas", ("simples nacional", "icms", "iss", "pis", "cofins",
                         "imposto", "devolu", "desconto concedido", "deduc")),
    ("receitas_financeiras", ("rendimento", "juros recebidos", "desconto obtido", "receita financeira")),
    ("despesas_financeiras", ("juros", "tarifa", "iof", "multa", "despesa financeira", "encargos financeiros")),
    ("amortizacao_emprestimos", ("emprestimo", "financiamento", "amortiza")),
    ("distribuicao_lucros", ("distribuicao", "lucros", "dividendo", "retirada")),
    ("depreciacao", ("deprecia",)),
    ("investimentos", ("imobilizado", "equipamento", "maquina", "veiculo", "reforma", "investimento")),
    # Frete de COMPRA integra o custo da mercadoria; os demais fretes são de venda (despesa comercial).
    ("custos_variaveis", ("frete sobre compra", "frete de compra", "frete compra", "frete s/ compra")),
    ("despesas_variaveis", TERMOS_DESPESAS_VARIAVEIS),
    ("custos_variaveis", ("cmv", "csv", "custo da mercadoria", "custo do servico", "mercadoria", "materia",
                          "insumo", "fornecedor", "embalage", "mao de obra direta", "terceirizacao do servico")),
    ("receita_bruta", ("venda", "receita", "servico prestado", "faturamento", "honorario")),
    ("custos_fixos", ("aluguel", "salari", "folha", "pro-labore", "pro labore", "prolabore", "encargo",
                      "fgts", "inss", "energia", "agua", "internet", "telefon", "contab", "sistema", "software",
                      "condominio", "iptu", "manutenc", "limpeza", "seguro", "vale transporte", "beneficio")),
]


def sugerir_linha(categoria: str, soma_valores: float, sinal_confiavel: bool = True) -> str:
    nome = _sem_acento(categoria)
    for linha, palavras in REGRAS_SUGESTAO:
        # \b no início: "iss" pega "ISS/ISSQN" mas não "comissões"
        if any(re.search(r"\b" + re.escape(p.strip()), nome) for p in palavras):
            return linha
    # Sem sinal no arquivo (modo "todos positivos"), não dá para saber se é entrada: o mais comum é despesa.
    if not sinal_confiavel:
        return "custos_fixos"
    return "receita_bruta" if soma_valores > 0 else "custos_fixos"

# =============================================================================
#  3. BANCO DE DADOS LOCAL (bpo.db)
# =============================================================================
CAMINHO_BANCO = os.environ.get("BPO_BANCO", os.path.join(os.path.dirname(os.path.abspath(__file__)), "bpo.db"))

ESQUEMA = """
CREATE TABLE IF NOT EXISTS clientes (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nome TEXT NOT NULL UNIQUE,
    pacote TEXT
);
CREATE TABLE IF NOT EXISTS lancamentos (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    cliente_id INTEGER NOT NULL REFERENCES clientes(id),
    competencia TEXT NOT NULL,          -- 'AAAA-MM'
    data TEXT NOT NULL,                 -- 'AAAA-MM-DD'
    categoria TEXT NOT NULL,
    descricao TEXT,
    valor REAL NOT NULL,
    arquivo TEXT,
    importado_em TEXT
);
CREATE INDEX IF NOT EXISTS ix_lanc_cliente_mes ON lancamentos(cliente_id, competencia);
CREATE TABLE IF NOT EXISTS mapeamento (
    cliente_id INTEGER NOT NULL REFERENCES clientes(id),
    categoria TEXT NOT NULL,
    linha TEXT NOT NULL,
    PRIMARY KEY (cliente_id, categoria)
);
CREATE TABLE IF NOT EXISTS orcamento (
    cliente_id INTEGER NOT NULL REFERENCES clientes(id),
    competencia TEXT NOT NULL,
    linha TEXT NOT NULL,
    valor REAL NOT NULL,
    PRIMARY KEY (cliente_id, competencia, linha)
);
CREATE TABLE IF NOT EXISTS comentarios_mes (
    cliente_id INTEGER NOT NULL REFERENCES clientes(id),
    competencia TEXT NOT NULL,
    texto TEXT,
    atualizado_em TEXT,
    PRIMARY KEY (cliente_id, competencia)
);
CREATE TABLE IF NOT EXISTS migracoes (
    nome TEXT PRIMARY KEY,
    aplicada_em TEXT,
    detalhe TEXT                         -- JSON com o que foi alterado
);
CREATE TABLE IF NOT EXISTS avisos_lidos (
    cliente_id INTEGER NOT NULL,
    aviso TEXT NOT NULL,
    PRIMARY KEY (cliente_id, aviso)
);
CREATE TABLE IF NOT EXISTS centros_custo (
    cliente_id INTEGER NOT NULL REFERENCES clientes(id),
    nome TEXT NOT NULL,
    ordem INTEGER,
    PRIMARY KEY (cliente_id, nome)
);
CREATE TABLE IF NOT EXISTS layout_importacao (
    cliente_id INTEGER PRIMARY KEY REFERENCES clientes(id),
    config TEXT NOT NULL,                -- colunas escolhidas na última importação (JSON)
    atualizado_em TEXT
);
CREATE TABLE IF NOT EXISTS saldos_iniciais (
    cliente_id INTEGER NOT NULL REFERENCES clientes(id),
    competencia TEXT NOT NULL,           -- saldo em banco no 1º dia deste mês
    valor REAL NOT NULL,
    PRIMARY KEY (cliente_id, competencia)
);
CREATE TABLE IF NOT EXISTS parametros_mes (
    cliente_id INTEGER NOT NULL REFERENCES clientes(id),
    competencia TEXT NOT NULL,
    funcionarios INTEGER,
    PRIMARY KEY (cliente_id, competencia)
);
"""


def conectar(caminho: str = None) -> sqlite3.Connection:
    con = sqlite3.connect(caminho or CAMINHO_BANCO, check_same_thread=False)
    con.executescript(ESQUEMA)
    # Migração: bancos criados antes da natureza da despesa ganham a coluna nova (nada é apagado).
    colunas_mapa = [c[1] for c in con.execute("PRAGMA table_info(mapeamento)").fetchall()]
    if "centro_custo" not in colunas_mapa:
        with con:
            con.execute("ALTER TABLE mapeamento ADD COLUMN centro_custo TEXT")
    # Migração: cada lançamento passa a saber se o sinal do arquivo é confiável.
    colunas_lanc = [c[1] for c in con.execute("PRAGMA table_info(lancamentos)").fetchall()]
    if "sinal_confiavel" not in colunas_lanc:
        with con:
            con.execute("ALTER TABLE lancamentos ADD COLUMN sinal_confiavel INTEGER")
            # Importações antigas: se o arquivo tinha algum valor negativo, o sinal era real;
            # se era tudo positivo, foi importado no modo "todos positivos".
            con.execute("""
                UPDATE lancamentos SET sinal_confiavel = CASE WHEN EXISTS (
                    SELECT 1 FROM lancamentos o
                    WHERE o.cliente_id = lancamentos.cliente_id
                      AND COALESCE(o.arquivo, '') = COALESCE(lancamentos.arquivo, '')
                      AND COALESCE(o.importado_em, '') = COALESCE(lancamentos.importado_em, '')
                      AND o.valor < 0) THEN 1 ELSE 0 END
                WHERE sinal_confiavel IS NULL""")
    _migrar_despesas_variaveis(con)
    return con


MIGRACAO_DESPESAS_VARIAVEIS = "separar_despesas_comerciais_variaveis"


def _migrar_despesas_variaveis(con):
    """Uma única vez: categorias comerciais (comissões, fretes de venda, taxas de cartão) que estavam em
    'Custos Variáveis' passam para 'Despesas Comerciais Variáveis'. Os lançamentos não mudam e a
    Margem de Contribuição continua com o mesmo valor; só aparece o Lucro Bruto no meio do caminho."""
    if con.execute("SELECT 1 FROM migracoes WHERE nome = ?", (MIGRACAO_DESPESAS_VARIAVEIS,)).fetchone():
        return
    frete_de_compra = ("frete sobre compra", "frete de compra", "frete compra", "frete s/ compra")
    movidas = {}
    for cliente_id, categoria in con.execute(
            "SELECT cliente_id, categoria FROM mapeamento WHERE linha = 'custos_variaveis'").fetchall():
        nome = _sem_acento(categoria)
        if _tem_termo(nome, TERMOS_DESPESAS_VARIAVEIS) and not _tem_termo(nome, frete_de_compra):
            movidas.setdefault(str(cliente_id), []).append(categoria)
    with con:
        con.executemany(
            "UPDATE mapeamento SET linha = 'despesas_variaveis' WHERE cliente_id = ? AND categoria = ?",
            [(int(c), cat) for c, cats in movidas.items() for cat in cats])
        con.execute("INSERT INTO migracoes (nome, aplicada_em, detalhe) VALUES (?, ?, ?)",
                    (MIGRACAO_DESPESAS_VARIAVEIS, datetime.now().isoformat(timespec="seconds"),
                     json.dumps(movidas, ensure_ascii=False)))


def categorias_movidas_na_migracao(con, cliente_id: int) -> list:
    """Categorias deste cliente movidas pela migração, se o aviso ainda não foi dispensado."""
    if con.execute("SELECT 1 FROM avisos_lidos WHERE cliente_id = ? AND aviso = ?",
                   (cliente_id, MIGRACAO_DESPESAS_VARIAVEIS)).fetchone():
        return []
    linha = con.execute("SELECT detalhe FROM migracoes WHERE nome = ?", (MIGRACAO_DESPESAS_VARIAVEIS,)).fetchone()
    try:
        return json.loads(linha[0]).get(str(cliente_id), []) if linha else []
    except (TypeError, ValueError):
        return []


def dispensar_aviso(con, cliente_id: int, aviso: str):
    with con:
        con.execute("INSERT OR IGNORE INTO avisos_lidos (cliente_id, aviso) VALUES (?, ?)", (cliente_id, aviso))


# Valor de cada lançamento na natureza da linha da DRE (positivo = receita para linhas credoras,
# positivo = despesa/saída para as demais). Lançamento a lançamento, ANTES de somar:
#  - arquivo com sinal: entrada (+) soma na receita; saída (-) vira despesa positiva; estorno (+) numa
#    despesa reduz a despesa; devolução (-) numa receita reduz a receita.
#  - arquivo "todos positivos": o sinal não existe, a categoria define (valor absoluto).
VALOR_NA_LINHA_SQL = """(CASE
    WHEN COALESCE(l.sinal_confiavel, 1) = 0 THEN ABS(l.valor)
    WHEN m.linha IN ('receita_bruta', 'receitas_financeiras') THEN l.valor
    ELSE -l.valor END)"""


# ----------------------------------------------------------------- clientes
def listar_clientes(con) -> pd.DataFrame:
    return pd.read_sql("SELECT id, nome, pacote FROM clientes ORDER BY nome", con)


def criar_cliente(con, nome: str, pacote: str = None) -> int:
    cur = con.execute("INSERT INTO clientes (nome, pacote) VALUES (?, ?)", (nome.strip(), pacote))
    con.commit()
    return cur.lastrowid


# -------------------------------------------------------------- lançamentos
def salvar_lancamentos(con, cliente_id: int, df: pd.DataFrame, arquivo: str, substituir_meses=True,
                       sinal_confiavel=True):
    """Grava os lançamentos. Com substituir_meses, apaga antes o que já havia
    nos meses presentes no arquivo (reimportar o mesmo mês não duplica).
    sinal_confiavel=False quando o arquivo veio no modo 'todos os valores positivos'."""
    meses = sorted(df["competencia"].unique())
    with con:
        if substituir_meses and meses:
            marcadores = ",".join("?" * len(meses))
            con.execute(
                f"DELETE FROM lancamentos WHERE cliente_id = ? AND competencia IN ({marcadores})",
                [cliente_id, *meses],
            )
        agora = datetime.now().isoformat(timespec="seconds")
        con.executemany(
            "INSERT INTO lancamentos (cliente_id, competencia, data, categoria, descricao, valor, arquivo, "
            "importado_em, sinal_confiavel) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
            [(cliente_id, r.competencia, r.data, r.categoria, r.descricao, float(r.valor), arquivo, agora,
              1 if sinal_confiavel else 0) for r in df.itertuples()],
        )
    return len(df), meses


def excluir_mes(con, cliente_id: int, competencia: str):
    with con:
        con.execute("DELETE FROM lancamentos WHERE cliente_id = ? AND competencia = ?", (cliente_id, competencia))


def competencias(con, cliente_id: int) -> list:
    linhas = con.execute(
        "SELECT DISTINCT competencia FROM lancamentos WHERE cliente_id = ? ORDER BY competencia", (cliente_id,)
    ).fetchall()
    return [l[0] for l in linhas]


def resumo_importacoes(con, cliente_id: int) -> pd.DataFrame:
    return pd.read_sql(
        "SELECT competencia, COUNT(*) AS lancamentos, MAX(importado_em) AS ultima_importacao, "
        "GROUP_CONCAT(DISTINCT arquivo) AS arquivos FROM lancamentos WHERE cliente_id = ? "
        "GROUP BY competencia ORDER BY competencia DESC",
        con, params=(cliente_id,),
    )


def totais_por_categoria(con, cliente_id: int, competencia: str = None) -> pd.DataFrame:
    """Soma dos valores por categoria (com a linha da DRE, se já classificada)."""
    filtro = "AND l.competencia = ?" if competencia else ""
    params = (cliente_id, competencia) if competencia else (cliente_id,)
    return pd.read_sql(
        f"""SELECT l.categoria, m.linha, SUM(l.valor) AS total, SUM({VALOR_NA_LINHA_SQL}) AS valor_dre,
                   COUNT(*) AS lancamentos
            FROM lancamentos l
            LEFT JOIN mapeamento m ON m.cliente_id = l.cliente_id AND m.categoria = l.categoria
            WHERE l.cliente_id = ? {filtro}
            GROUP BY l.categoria, m.linha ORDER BY l.categoria""",
        con, params=params,
    )


# ------------------------------------------------------------ fluxo de caixa
# Agrupamento da DFC gerencial (não altera a DRE: só lê os mesmos lançamentos).
GRUPO_DFC = {
    "receita_bruta": "Operacional", "impostos_vendas": "Operacional", "custos_variaveis": "Operacional",
    "despesas_variaveis": "Operacional", "custos_fixos": "Operacional", "receitas_financeiras": "Operacional",
    "despesas_financeiras": "Operacional", "impostos_lucro": "Operacional",
    "investimentos": "Investimentos",
    "amortizacao_emprestimos": "Financiamento e Sócios", "distribuicao_lucros": "Financiamento e Sócios",
}
ORDEM_DFC = ["Operacional", "Investimentos", "Financiamento e Sócios", "Não classificado"]
SEM_CLASSIFICACAO = "Sem classificação"


def lancamentos_caixa(con, cliente_id: int, competencia: str) -> pd.DataFrame:
    """Lançamentos do mês com o efeito no caixa: + entrada, - saída.
    Usa o mesmo valor-na-natureza da DRE (estornos já compensados). Ficam de fora transferências
    entre contas (ignorar) e depreciação (não mexe no caixa). Categorias sem classificação entram
    pelo sinal do arquivo; se o arquivo não tinha sinal, ficam de fora (avisadas na tela)."""
    df = pd.read_sql(
        f"""SELECT l.data, l.categoria, l.descricao, l.valor, COALESCE(l.sinal_confiavel, 1) AS sinal_confiavel,
                   m.linha, {VALOR_NA_LINHA_SQL} AS valor_dre
            FROM lancamentos l
            LEFT JOIN mapeamento m ON m.cliente_id = l.cliente_id AND m.categoria = l.categoria
            WHERE l.cliente_id = ? AND l.competencia = ?
            ORDER BY l.data, l.id""",
        con, params=(cliente_id, competencia),
    )
    if df.empty:
        return df.assign(caixa=pd.Series(dtype=float), grupo=pd.Series(dtype=str))
    df = df[~df["linha"].isin(["ignorar", "depreciacao"])].copy()
    credora = df["linha"].isin(LINHAS_CREDORAS)
    df["caixa"] = df["valor_dre"].where(credora, -df["valor_dre"])
    sem_linha = df["linha"].isna()
    df.loc[sem_linha, "caixa"] = df.loc[sem_linha, "valor"].where(df.loc[sem_linha, "sinal_confiavel"] == 1)
    df["grupo"] = df["linha"].map(GRUPO_DFC).fillna("Não classificado")
    df["linha_nome"] = (df["linha"].map(LINHAS).fillna(SEM_CLASSIFICACAO).str.split(" – ").str[0]
                        .str.replace(r" \(.*\)$", "", regex=True))
    return df.reset_index(drop=True)


def carregar_saldo_inicial(con, cliente_id: int, competencia: str):
    linha = con.execute("SELECT valor FROM saldos_iniciais WHERE cliente_id = ? AND competencia = ?",
                        (cliente_id, competencia)).fetchone()
    return float(linha[0]) if linha else None


def salvar_saldo_inicial(con, cliente_id: int, competencia: str, valor):
    with con:
        if valor is None:
            con.execute("DELETE FROM saldos_iniciais WHERE cliente_id = ? AND competencia = ?",
                        (cliente_id, competencia))
        else:
            con.execute(
                "INSERT INTO saldos_iniciais (cliente_id, competencia, valor) VALUES (?, ?, ?) "
                "ON CONFLICT(cliente_id, competencia) DO UPDATE SET valor = excluded.valor",
                (cliente_id, competencia, float(valor)))


def saldo_inicial_do_mes(con, cliente_id: int, competencia: str):
    """Saldo no 1º dia do mês e a origem: 'informado' (digitado para este mês), 'calculado' (saldo
    informado num mês anterior + fluxo dos meses seguintes) ou 'sem_saldo' (nenhum informado: parte de 0)."""
    informado = carregar_saldo_inicial(con, cliente_id, competencia)
    if informado is not None:
        return informado, "informado", competencia
    base = con.execute("SELECT competencia, valor FROM saldos_iniciais WHERE cliente_id = ? AND competencia < ? "
                       "ORDER BY competencia DESC LIMIT 1", (cliente_id, competencia)).fetchone()
    if not base:
        return 0.0, "sem_saldo", None
    saldo, m = float(base[1]), base[0]
    while m < competencia:
        mov = lancamentos_caixa(con, cliente_id, m)
        saldo += float(mov["caixa"].sum()) if len(mov) else 0.0
        m = mes_seguinte(m)
    return saldo, "calculado", base[0]


def extrato_diario(lanc: pd.DataFrame, competencia: str, saldo_inicial: float) -> pd.DataFrame:
    """Uma linha por dia do mês: entradas, saídas e saldo acumulado ao fim do dia."""
    dias = pd.date_range(f"{competencia}-01", periods=1, freq="MS")[0]
    todos = pd.date_range(dias, dias + pd.offsets.MonthEnd(0), freq="D")
    validos = lanc.dropna(subset=["caixa"])
    datas = pd.to_datetime(validos["data"])
    entradas = validos["caixa"].clip(lower=0).groupby(datas).sum()
    saidas = (-validos["caixa"].clip(upper=0)).groupby(datas).sum()
    df = pd.DataFrame({"dia": todos})
    df["entradas"] = df["dia"].map(entradas).fillna(0.0)
    df["saidas"] = df["dia"].map(saidas).fillna(0.0)
    df["saldo"] = saldo_inicial + (df["entradas"] - df["saidas"]).cumsum()
    return df


def dfc_gerencial(lanc: pd.DataFrame) -> pd.DataFrame:
    """Totais por grupo da DFC e por linha da DRE dentro dele (com entradas e saídas)."""
    validos = lanc.dropna(subset=["caixa"])
    if validos.empty:
        return pd.DataFrame(columns=["grupo", "linha_nome", "entradas", "saidas", "liquido"])
    validos = validos.assign(linha=validos["linha"].fillna(""))
    df = validos.groupby(["grupo", "linha", "linha_nome"], as_index=False).agg(
        entradas=("caixa", lambda x: x.clip(lower=0).sum()), saidas=("caixa", lambda x: -x.clip(upper=0).sum()),
        liquido=("caixa", "sum"))
    df["ordem"] = df["grupo"].map({g: i for i, g in enumerate(ORDEM_DFC)})
    df["ordem_linha"] = df["linha"].map({k: i for i, k in enumerate(LINHAS)}).fillna(99)  # mesma ordem da DRE
    return (df.sort_values(["ordem", "ordem_linha"]).drop(columns=["ordem", "ordem_linha", "linha"])
            .reset_index(drop=True))


def valores_por_mes(con, cliente_id: int) -> pd.DataFrame:
    """Tabela mês x linha da DRE, na natureza de cada linha (ver VALOR_NA_LINHA_SQL).
    O saldo é líquido: estornos reduzem a despesa e devoluções reduzem a receita. Se o mês
    ficar com saldo contrário à natureza da linha, o valor fica NEGATIVO (não é trocado de sinal)."""
    df = pd.read_sql(
        f"""SELECT l.competencia, m.linha, SUM({VALOR_NA_LINHA_SQL}) AS total
           FROM lancamentos l
           JOIN mapeamento m ON m.cliente_id = l.cliente_id AND m.categoria = l.categoria
           WHERE l.cliente_id = ? AND m.linha <> 'ignorar'
           GROUP BY l.competencia, m.linha""",
        con, params=(cliente_id,),
    )
    if df.empty:
        return pd.DataFrame()
    return df.pivot(index="competencia", columns="linha", values="total").fillna(0).sort_index()


# -------------------------------------------------------- plano de contas
def carregar_centros_categorias(con, cliente_id: int) -> dict:
    return dict(con.execute("SELECT categoria, centro_custo FROM mapeamento WHERE cliente_id = ?",
                            (cliente_id,)).fetchall())


def carregar_mapeamento(con, cliente_id: int) -> dict:
    linhas = con.execute("SELECT categoria, linha FROM mapeamento WHERE cliente_id = ?", (cliente_id,)).fetchall()
    return dict(linhas)


def salvar_mapeamento(con, cliente_id: int, mapeamento: dict):
    with con:
        con.executemany(
            "INSERT INTO mapeamento (cliente_id, categoria, linha) VALUES (?, ?, ?) "
            "ON CONFLICT(cliente_id, categoria) DO UPDATE SET linha = excluded.linha",
            [(cliente_id, cat, linha) for cat, linha in mapeamento.items()],
        )


# ------------------------------------------------------------- parâmetros
def obter_funcionarios(con, cliente_id: int, competencia: str) -> int:
    """Quantidade do mês; se não informada, repete a do último mês anterior."""
    linha = con.execute(
        "SELECT funcionarios FROM parametros_mes WHERE cliente_id = ? AND competencia <= ? "
        "AND funcionarios IS NOT NULL ORDER BY competencia DESC LIMIT 1",
        (cliente_id, competencia),
    ).fetchone()
    return int(linha[0]) if linha else 0


def salvar_funcionarios(con, cliente_id: int, competencia: str, quantidade: int):
    with con:
        con.execute(
            "INSERT INTO parametros_mes (cliente_id, competencia, funcionarios) VALUES (?, ?, ?) "
            "ON CONFLICT(cliente_id, competencia) DO UPDATE SET funcionarios = excluded.funcionarios",
            (cliente_id, competencia, int(quantidade)),
        )

# ------------------------------------------------- natureza da despesa
CENTROS_PADRAO = ["Pessoal", "Ocupação e Utilidades", "Manutenção", "Administrativo", "Tecnologia",
                  "Comercial e Marketing", "Outras despesas fixas"]
CENTRO_RESERVA = "Outras despesas fixas"

# Sugestão pelo nome da categoria (o operador sempre pode trocar no Plano de Contas).
REGRAS_CENTRO = [
    ("Manutenção", ["manutenc", "reparo", "conserto", "pecas", "assistencia tecnica", "dedetiz"]),
    ("Pessoal", ["salari", "folha", "pro-labore", "pro labore", "prolabore", "encargo", "fgts", "inss", "ferias",
                 "13", "decimo", "rescis", "vale transporte", "vale refeic", "vale alimenta", "beneficio",
                 "plano de saude", "uniforme", "treinamento", "premiac", "funcionari", "hora extra", "estagi"]),
    ("Ocupação e Utilidades", ["aluguel", "condominio", "iptu", "agua", "energia", "luz", "gas tub", "gas enc",
                               "gas de", "gas glp", "botijao", "esgoto",
                               "limpeza", "seguranca", "vigilancia", "seguro"]),
    ("Tecnologia", ["sistema", "software", "internet", "telefon", "informatic", "licenc", "hospedagem",
                    "tv a cabo", "celular", "pms", "channel", "erp", "nuvem", "dominio"]),
    ("Comercial e Marketing", ["marketing", "publicidade", "propaganda", "anuncio", "google", "meta ads", "brinde",
                               "evento", "patrocin", "redes sociais"]),
    ("Administrativo", ["contab", "honorario", "consultor", "advoca", "juridic", "correio", "cartorio",
                        "material de escritorio", "papelaria", "despachante", "assinatura", "taxa", "alvara"]),
]


def listar_centros(con, cliente_id: int) -> list:
    linhas = con.execute("SELECT nome FROM centros_custo WHERE cliente_id = ? ORDER BY ordem, nome",
                         (cliente_id,)).fetchall()
    if not linhas:  # primeira vez: cria a lista padrão para o cliente
        with con:
            con.executemany("INSERT OR IGNORE INTO centros_custo (cliente_id, nome, ordem) VALUES (?, ?, ?)",
                            [(cliente_id, n, i) for i, n in enumerate(CENTROS_PADRAO)])
        return list(CENTROS_PADRAO)
    return [l[0] for l in linhas]


def criar_centro(con, cliente_id: int, nome: str):
    ordem = con.execute("SELECT COALESCE(MAX(ordem), 0) + 1 FROM centros_custo WHERE cliente_id = ?",
                        (cliente_id,)).fetchone()[0]
    with con:
        con.execute("INSERT OR IGNORE INTO centros_custo (cliente_id, nome, ordem) VALUES (?, ?, ?)",
                    (cliente_id, nome.strip(), ordem))


def remover_centro(con, cliente_id: int, nome: str) -> int:
    """Remove o centro; as categorias que o usavam vão para 'Outras despesas fixas'. Devolve quantas mudaram."""
    with con:
        mudadas = con.execute("UPDATE mapeamento SET centro_custo = ? WHERE cliente_id = ? AND centro_custo = ?",
                              (CENTRO_RESERVA, cliente_id, nome)).rowcount
        con.execute("DELETE FROM centros_custo WHERE cliente_id = ? AND nome = ?", (cliente_id, nome))
    if mudadas:
        criar_centro(con, cliente_id, CENTRO_RESERVA)
    return mudadas


def sugerir_centro(categoria: str, centros: list) -> str:
    nome = _sem_acento(categoria)
    for centro, termos in REGRAS_CENTRO:
        if centro in centros and _tem_termo(nome, termos):
            return centro
    return CENTRO_RESERVA if CENTRO_RESERVA in centros else (centros[-1] if centros else CENTRO_RESERVA)


def salvar_centros_categorias(con, cliente_id: int, centros_por_categoria: dict):
    with con:
        con.executemany("UPDATE mapeamento SET centro_custo = ? WHERE cliente_id = ? AND categoria = ?",
                        [(centro, cliente_id, cat) for cat, centro in centros_por_categoria.items()])


def preencher_centros_vazios(con, cliente_id: int) -> int:
    """Despesas fixas sem natureza da despesa recebem uma sugestão. Devolve quantas foram preenchidas."""
    centros = listar_centros(con, cliente_id)
    sem_centro = [l[0] for l in con.execute(
        "SELECT categoria FROM mapeamento WHERE cliente_id = ? AND linha = 'custos_fixos' "
        "AND (centro_custo IS NULL OR centro_custo = '' OR centro_custo NOT IN "
        "(SELECT nome FROM centros_custo WHERE cliente_id = ?))", (cliente_id, cliente_id)).fetchall()]
    if sem_centro:
        salvar_centros_categorias(con, cliente_id, {c: sugerir_centro(c, centros) for c in sem_centro})
        if CENTRO_RESERVA not in centros:
            criar_centro(con, cliente_id, CENTRO_RESERVA)
    return len(sem_centro)


# Linhas da DRE que são despesa (as "só caixa" ficam de fora da aba Despesas).
LINHAS_DESPESA = ["custos_fixos", "custos_variaveis", "despesas_variaveis", "impostos_vendas", "despesas_financeiras",
                  "impostos_lucro", "depreciacao"]


def despesas_detalhadas(con, cliente_id: int) -> pd.DataFrame:
    """Uma linha por mês e categoria de despesa, com linha da DRE e natureza da despesa. Valor positivo."""
    marcadores = ",".join("?" * len(LINHAS_DESPESA))
    df = pd.read_sql(
        f"""SELECT l.competencia, l.categoria, m.linha, m.centro_custo AS centro,
                   SUM({VALOR_NA_LINHA_SQL}) AS total
            FROM lancamentos l
            JOIN mapeamento m ON m.cliente_id = l.cliente_id AND m.categoria = l.categoria
            WHERE l.cliente_id = ? AND m.linha IN ({marcadores})
            GROUP BY l.competencia, l.categoria, m.linha, m.centro_custo""",
        con, params=(cliente_id, *LINHAS_DESPESA),
    )
    df["valor"] = df["total"]  # despesa positiva; estornos maiores que a despesa ficam negativos
    df["centro"] = df["centro"].fillna("Sem natureza da despesa")
    return df.drop(columns="total")


# ------------------------------------------------- layout de importação
def carregar_layout(con, cliente_id: int) -> dict:
    linha = con.execute("SELECT config FROM layout_importacao WHERE cliente_id = ?", (cliente_id,)).fetchone()
    try:
        return json.loads(linha[0]) if linha else {}
    except (TypeError, ValueError):
        return {}


def salvar_layout(con, cliente_id: int, config: dict):
    with con:
        con.execute(
            "INSERT INTO layout_importacao (cliente_id, config, atualizado_em) VALUES (?, ?, ?) "
            "ON CONFLICT(cliente_id) DO UPDATE SET config = excluded.config, atualizado_em = excluded.atualizado_em",
            (cliente_id, json.dumps(config, ensure_ascii=False), datetime.now().isoformat(timespec="seconds")),
        )


# --------------------------------------------------------------- orçamento
def carregar_orcamento(con, cliente_id: int, competencia: str) -> dict:
    linhas = con.execute(
        "SELECT linha, valor FROM orcamento WHERE cliente_id = ? AND competencia = ?", (cliente_id, competencia)
    ).fetchall()
    return dict(linhas)


def salvar_orcamento(con, cliente_id: int, competencia: str, valores: dict):
    """Substitui o orçamento do mês pelos valores informados (zeros não são gravados)."""
    with con:
        con.execute("DELETE FROM orcamento WHERE cliente_id = ? AND competencia = ?", (cliente_id, competencia))
        con.executemany(
            "INSERT INTO orcamento (cliente_id, competencia, linha, valor) VALUES (?, ?, ?, ?)",
            [(cliente_id, competencia, linha, float(v)) for linha, v in valores.items() if v],
        )


def meses_com_orcamento(con, cliente_id: int) -> list:
    linhas = con.execute(
        "SELECT DISTINCT competencia FROM orcamento WHERE cliente_id = ? ORDER BY competencia", (cliente_id,)
    ).fetchall()
    return [l[0] for l in linhas]


# ------------------------------------------------------------- comentários
def carregar_comentario(con, cliente_id: int, competencia: str) -> str:
    linha = con.execute(
        "SELECT texto FROM comentarios_mes WHERE cliente_id = ? AND competencia = ?", (cliente_id, competencia)
    ).fetchone()
    return (linha[0] or "") if linha else ""


def salvar_comentario(con, cliente_id: int, competencia: str, texto: str):
    with con:
        con.execute(
            "INSERT INTO comentarios_mes (cliente_id, competencia, texto, atualizado_em) VALUES (?, ?, ?, ?) "
            "ON CONFLICT(cliente_id, competencia) DO UPDATE SET texto = excluded.texto, "
            "atualizado_em = excluded.atualizado_em",
            (cliente_id, competencia, texto.strip(), datetime.now().isoformat(timespec="seconds")),
        )


# =============================================================================
#  4. ORÇADO x REALIZADO
# =============================================================================
# Linhas onde ficar ACIMA do orçado é bom. Nas demais (custos, impostos,
# despesas, saídas de caixa) ficar acima é estouro.
FAVORAVEL_SE_MAIOR = {
    "receita_bruta", "receitas_financeiras", "receita_liquida", "lucro_bruto", "margem_contribuicao",
    "ebitda", "lucro_liquido", "geracao_caixa",
}
TOLERANCIA_PCT = 5.0  # desvio desfavorável até 5% = atenção; acima = alerta

ESTRUTURA_ORCAMENTO = [
    ("Receita Operacional Bruta", "receita_bruta"),
    ("(-) Deduções / Impostos sobre Vendas", "impostos_vendas"),
    ("(=) Receita Líquida", "receita_liquida"),
    ("(-) CMV / CSV", "custos_variaveis"),
    ("(=) Lucro Bruto", "lucro_bruto"),
    ("(-) Despesas Comerciais Variáveis", "despesas_variaveis"),
    ("(=) Margem de Contribuição", "margem_contribuicao"),
    ("(-) Custos e Despesas Fixas", "custos_fixos"),
    ("(=) EBITDA", "ebitda"),
    ("(-) Depreciação e Amortização", "depreciacao"),
    ("(+) Receitas Financeiras", "receitas_financeiras"),
    ("(-) Despesas Financeiras", "despesas_financeiras"),
    ("(-) Impostos sobre o Lucro (IR/CSLL)", "impostos_lucro"),
    ("(=) Lucro Líquido", "lucro_liquido"),
    ("(-) Amortização de Empréstimos", "amortizacao_emprestimos"),
    ("(-) Investimentos / Compra de Ativos", "investimentos"),
    ("(-) Distribuição de Lucros / Retiradas", "distribuicao_lucros"),
    ("(=) Geração de Caixa do Mês", "geracao_caixa"),
]


def comparar_orcamento(r_orcado: dict, r_real: dict) -> list:
    """Lista de dicionários com orçado, realizado, diferença e status de cada linha.
    nivel: 0 = ok, 1 = atenção, 2 = alerta."""
    resultado = []
    for rotulo, chave in ESTRUTURA_ORCAMENTO:
        orcado, real = r_orcado[chave], r_real[chave]
        dif = real - orcado
        dif_pct = (dif / abs(orcado) * 100) if orcado else None
        maior_e_bom = chave in FAVORAVEL_SE_MAIOR
        if not orcado and not real:
            nivel, status = 0, "—"
        elif (dif >= 0) if maior_e_bom else (dif <= 0):
            nivel, status = 0, "🟢 OK"
        else:
            desvio = abs(dif_pct) if dif_pct is not None else float("inf")
            nivel = 1 if desvio <= TOLERANCIA_PCT else 2
            if nivel == 1:
                status = "🟡 Atenção"
            else:
                status = "🔴 Abaixo do orçado" if maior_e_bom else "🔴 Estouro"
        resultado.append({
            "rotulo": rotulo, "chave": chave, "orcado": orcado, "realizado": real,
            "dif": dif, "dif_pct": dif_pct, "nivel": nivel, "status": status,
            "total": rotulo.startswith("(=)") or chave == "receita_bruta",
        })
    return resultado


def _para_competencia(cabecalho):
    """Aceita '2026-10', '10/2026', 'out/2026', 'outubro 2026', 'out/26' ou data do Excel."""
    if hasattr(cabecalho, "year"):
        return f"{cabecalho.year}-{cabecalho.month:02d}"
    s = _sem_acento(cabecalho)
    m = re.fullmatch(r"(\d{4})[-/.](\d{1,2})", s)
    if m:
        ano, mes = int(m[1]), int(m[2])
    else:
        m = re.fullmatch(r"(\d{1,2})[-/.](\d{4})", s)
        if m:
            mes, ano = int(m[1]), int(m[2])
        else:
            m = re.fullmatch(r"([a-z]{3})[a-z]*\.?[-/ ]*(\d{2}|\d{4})", s)
            if not m or m[1] not in MESES:
                return None
            mes, ano = MESES.index(m[1]) + 1, int(m[2])
            ano = ano + 2000 if ano < 100 else ano
    return f"{ano}-{mes:02d}" if 1 <= mes <= 12 else None


# Nomes antigos das linhas (planilhas de orçamento feitas antes da separação continuam valendo)
NOMES_ANTIGOS_LINHAS = {"custos variaveis / cmv": "custos_variaveis", "custos variaveis": "custos_variaveis",
                        "cmv": "custos_variaveis", "csv": "custos_variaveis"}


def _linha_do_orcamento(texto):
    alvo = _sem_acento(texto)
    if alvo in NOMES_ANTIGOS_LINHAS:
        return NOMES_ANTIGOS_LINHAS[alvo]
    for chave in LINHAS_VALOR:
        if alvo in (chave, _sem_acento(LINHAS[chave])):
            return chave
    for chave in LINHAS_VALOR:  # aceita o começo do nome (ex.: "Custos Variáveis")
        if alvo and _sem_acento(LINHAS[chave]).startswith(alvo):
            return chave
    return None


def modelo_orcamento(meses_modelo: list, preenchido: dict) -> bytes:
    """Planilha para preencher: linhas da DRE x meses. preenchido = {mês: {linha: valor}}."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill

    wb = Workbook()
    ws = wb.active
    ws.title = "Orçamento"
    ws.append(["Linha da DRE"] + [nome_mes(m) for m in meses_modelo])
    for chave in LINHAS_VALOR:
        ws.append([LINHAS[chave]] + [preenchido.get(m, {}).get(chave) for m in meses_modelo])
    cabecalho = PatternFill("solid", fgColor="1F3A5F")
    for celula in ws[1]:
        celula.font = Font(bold=True, color="FFFFFF")
        celula.fill = cabecalho
        celula.alignment = Alignment(horizontal="center")
    for linha in ws.iter_rows(min_row=2, min_col=2):
        for celula in linha:
            celula.number_format = '#,##0.00'
    ws.column_dimensions["A"].width = 52
    for i in range(len(meses_modelo)):
        ws.column_dimensions[ws.cell(1, i + 2).column_letter].width = 14
    ws.freeze_panes = "B2"

    ajuda = wb.create_sheet("Como preencher")
    for texto in [
        "COMO PREENCHER O ORÇAMENTO",
        "",
        "1. Na aba 'Orçamento', digite o valor previsto de cada linha em cada mês.",
        "2. Digite todos os valores POSITIVOS (o sistema já sabe o que é receita e o que é custo).",
        "3. Deixe em branco o que não tiver previsão.",
        "4. Pode acrescentar meses: escreva o mês no cabeçalho no formato out/2026 ou 10/2026.",
        "5. Não mude os nomes da coluna 'Linha da DRE'.",
        "6. Salve e envie o arquivo na tela 'Orçado x Realizado' do painel.",
    ]:
        ajuda.append([texto])
    ajuda["A1"].font = Font(bold=True, size=13)
    ajuda.column_dimensions["A"].width = 95

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


def ler_orcamento(conteudo: bytes):
    """Lê a planilha de orçamento. Devolve ({mês: {linha: valor}}, avisos)."""
    tabela = pd.read_excel(io.BytesIO(conteudo), sheet_name=0, header=0, dtype=object)
    avisos = []
    colunas_mes = {}
    for coluna in tabela.columns[1:]:
        comp = _para_competencia(coluna)
        if comp:
            colunas_mes[coluna] = comp
        elif not str(coluna).startswith("Unnamed"):
            avisos.append(f"Coluna '{coluna}' ignorada: não reconheci o mês.")
    orcamento = {}
    for _, linha in tabela.iterrows():
        nome = linha.iloc[0]
        if pd.isna(nome) or not str(nome).strip():
            continue
        chave = _linha_do_orcamento(nome)
        if not chave:
            avisos.append(f"Linha '{nome}' ignorada: não corresponde a nenhuma linha da DRE.")
            continue
        for coluna, comp in colunas_mes.items():
            valor = para_numero(pd.Series([linha[coluna]])).iloc[0]
            if pd.notna(valor) and valor:
                orcamento.setdefault(comp, {})[chave] = abs(float(valor))
    return orcamento, avisos


def mes_seguinte(competencia: str, n: int = 1) -> str:
    ano, mes = map(int, competencia.split("-"))
    total = ano * 12 + (mes - 1) + n
    return f"{total // 12}-{total % 12 + 1:02d}"


# ---------------------------------------------- comparativo ano anterior
CORES_NIVEL = {2: "background-color: #fde2e1; color: #7a1414",
               1: "background-color: #fff4d6; color: #5c3d00", 0: ""}


def com_sinal(valor: float) -> str:
    return ("+" if valor > 0.004 else "") + brl(valor)


def pct_com_sinal(valor) -> str:
    return "–" if valor is None else ("+" if valor > 0.04 else "") + pct(valor)


def variacao_pct(atual: float, base: float):
    return (atual - base) / abs(base) * 100 if base else None


def avaliar_variacao(dif: float, dif_pct, maior_e_bom):
    """Devolve (tendência, leitura, nível). maior_e_bom=None quando não dá para julgar."""
    if abs(dif) < 0.005:
        return "➖ Igual", "—", 0
    tendencia = "🔺 Subiu" if dif > 0 else "🔻 Caiu"
    if maior_e_bom is None:
        return tendencia, "—", 0
    if (dif > 0) == maior_e_bom:
        return tendencia, "🟢 Favorável", 0
    if dif_pct is not None and abs(dif_pct) <= TOLERANCIA_PCT:
        return tendencia, "🟡 Atenção", 1
    return tendencia, "🔴 Desfavorável", 2


def comparar_ano_anterior(r_atual: dict, r_base: dict) -> list:
    """Mesmo mês do ano anterior x mês atual, linha a linha da DRE e do caixa."""
    linhas = []
    for rotulo, chave in ESTRUTURA_ORCAMENTO:
        atual, base = r_atual[chave], r_base[chave]
        dif = atual - base
        dif_pct = variacao_pct(atual, base)
        tendencia, leitura, nivel = avaliar_variacao(dif, dif_pct, chave in FAVORAVEL_SE_MAIOR)
        linhas.append({"rotulo": rotulo, "chave": chave, "atual": atual, "base": base, "dif": dif,
                       "dif_pct": dif_pct, "tendencia": tendencia, "leitura": leitura, "nivel": nivel,
                       "total": rotulo.startswith("(=)") or chave == "receita_bruta"})
    return linhas


def comparar_categorias(totais_atual: pd.DataFrame, totais_base: pd.DataFrame) -> pd.DataFrame:
    """Cada categoria do ERP no mês atual x mesmo mês do ano anterior (ex.: conta de água)."""
    a = totais_atual[["categoria", "linha", "valor_dre"]].rename(columns={"valor_dre": "atual", "linha": "linha_a"})
    b = totais_base[["categoria", "linha", "valor_dre"]].rename(columns={"valor_dre": "base", "linha": "linha_b"})
    df = a.merge(b, on="categoria", how="outer")
    df["linha"] = df["linha_a"].combine_first(df["linha_b"])
    df = df[df["linha"] != "ignorar"].copy()
    # Valores já na natureza da linha (despesa positiva, estorno reduz); sem dado = 0.
    df["atual"] = df["atual"].fillna(0)
    df["base"] = df["base"].fillna(0)
    df["dif"] = df["atual"] - df["base"]
    df["dif_pct"] = [variacao_pct(x, y) for x, y in zip(df["atual"], df["base"])]
    avaliacoes = [
        avaliar_variacao(d, p, None if pd.isna(l) else l in FAVORAVEL_SE_MAIOR)
        for d, p, l in zip(df["dif"], df["dif_pct"], df["linha"])
    ]
    df["tendencia"] = [t for t, _, _ in avaliacoes]
    df["leitura"] = [l for _, l, _ in avaliacoes]
    df["nivel"] = [n for _, _, n in avaliacoes]
    df["situacao"] = ["🆕 Nova" if not y and x else ("⛔ Não houve" if not x and y else "")
                      for x, y in zip(df["atual"], df["base"])]
    df["eh_receita"] = df["linha"].isin(FAVORAVEL_SE_MAIOR)
    # Despesa recorrente que sumiu costuma ser lançamento faltando, não economia: pede conferência.
    sumiu = (df["situacao"] == "⛔ Não houve") & ~df["eh_receita"] & df["linha"].notna()
    df.loc[sumiu, "leitura"] = "🟡 Conferir: sem lançamento"
    df.loc[sumiu, "nivel"] = 1
    return df.assign(ordem=df["dif"].abs()).sort_values("ordem", ascending=False).drop(columns="ordem") \
        .drop(columns=["linha_a", "linha_b"]).reset_index(drop=True)


# =============================================================================
#  5. QUEBRA DO FATURAMENTO E RESUMO EXECUTIVO EM PDF
# =============================================================================
# Cores da paleta categórica validada (teste de daltonismo para vizinhos).
CORES_QUEBRA = ["#2a78d6", "#eb6834", "#0e9fb5", "#eda100", "#e87ba4"]
TEXTO_SOBRE_COR = ["#ffffff", "#ffffff", "#0b0b0b", "#0b0b0b", "#0b0b0b"]


def quebra_faturamento(r: dict):
    """Para onde foi cada real faturado: impostos, CMV/CSV, despesas comerciais variáveis, custos fixos e sobra.
    Devolve (partes, aviso). partes = [(nome, valor, % do faturamento)]."""
    bruta = r["receita_bruta"]
    if bruta <= 0:
        return [], "Sem faturamento no mês."
    impostos, cmv, comerciais = r["impostos_vendas"], r["custos_variaveis"], r["despesas_variaveis"]
    margem = r["margem_contribuicao"]
    invertidas = [nome for nome, v in [("impostos sobre vendas", impostos), ("CMV/CSV", cmv),
                                        ("despesas comerciais variáveis", comerciais),
                                        ("custos fixos", r["custos_fixos"])] if v < 0]
    if invertidas:
        return [], (f"Gráfico não exibido: {', '.join(invertidas)} ficou com saldo credor neste mês "
                    "(estornos maiores que os lançamentos). Veja o aviso no topo do painel.")
    aviso = None
    if margem < 0:
        partes = [("Impostos sobre vendas", impostos), ("CMV / CSV", cmv), ("Despesas comerciais variáveis", comerciais)]
        aviso = (f"Impostos, CMV e despesas comerciais variáveis passaram do faturamento em {brl(-margem)}: "
                 "cada venda está dando prejuízo antes mesmo dos custos fixos.")
    else:
        fixos_cobertos = min(r["custos_fixos"], margem)
        partes = [("Impostos sobre vendas", impostos), ("CMV / CSV", cmv),
                  ("Despesas comerciais variáveis", comerciais),
                  ("Custos fixos", fixos_cobertos), ("Sobra operacional (EBITDA)", margem - fixos_cobertos)]
        if r["custos_fixos"] > margem:
            aviso = (f"A margem de contribuição não cobriu os custos fixos: faltaram "
                     f"{brl(r['custos_fixos'] - margem)} (prejuízo operacional).")
    return [(nome, valor, valor / bruta * 100) for nome, valor in partes], aviso


# Paleta categórica validada (8 cores em ordem fixa; a 9ª+ vira "Outros" nos gráficos).
CORES_CATEGORICAS = ["#2a78d6", "#eb6834", "#0e9fb5", "#eda100", "#e87ba4", "#008300", "#6250d6", "#e34948"]
COR_OUTROS = "#9b9a94"
GRUPOS_OUTRAS_DESPESAS = {
    "custos_variaveis": "CMV / CSV",
    "despesas_variaveis": "Despesas comerciais variáveis",
    "impostos_vendas": "Impostos sobre vendas",
    "despesas_financeiras": "Despesas financeiras",
    "impostos_lucro": "IR / CSLL",
    "depreciacao": "Depreciação",
}


def cores_por_grupo(ordem_grupos: list) -> dict:
    """A cor segue o grupo (não a posição): os 8 maiores ganham cor própria, o resto fica cinza."""
    return {g: (CORES_CATEGORICAS[i] if i < len(CORES_CATEGORICAS) else COR_OUTROS)
            for i, g in enumerate(ordem_grupos)}



# =============================================================================
#  GRÁFICOS INTERATIVOS (PLOTLY) – identidade Eficaz Performance BPO
# =============================================================================
try:
    import plotly.express as px
    import plotly.graph_objects as go
    PLOTLY_OK = True
except ImportError:  # o lançador instala; se faltar, os gráficos mostram um aviso em vez de quebrar
    PLOTLY_OK = False

COR_VERDE = "#0f7a3a"       # receitas / lucros
COR_VERMELHO = "#b42318"    # despesas / prejuízos
FONTE_GRAFICOS = "Inter, Segoe UI, sans-serif"


def layout_marca(fig, altura: int = 380, legenda: bool = True, eixo_reais: bool = True):
    """Fundo transparente, fonte e cores da marca, números no padrão brasileiro (1.234,56)."""
    fig.update_layout(
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family=FONTE_GRAFICOS, color=COR_TEXTO, size=13),
        separators=",.",  # vírgula decimal, ponto de milhar
        height=altura, margin=dict(l=8, r=8, t=24, b=8),
        showlegend=legenda,
        legend=dict(orientation="h", yanchor="top", y=-0.12, xanchor="left", x=0, title=None,
                    font=dict(size=12, color=COR_TEXTO_SECUNDARIO)),
        hoverlabel=dict(bgcolor="#FFFFFF", bordercolor=COR_BORDA, font=dict(family=FONTE_GRAFICOS, size=13,
                                                                            color=COR_TEXTO)),
    )
    fig.update_xaxes(showgrid=False, linecolor="#C5D0DD", tickfont=dict(color=COR_TEXTO_SECUNDARIO),
                     title=None)
    fig.update_yaxes(gridcolor="#E3E9F1", zerolinecolor="#9AA8BA", zerolinewidth=1,
                     tickfont=dict(color=COR_TEXTO_SECUNDARIO), title=None)
    if eixo_reais:
        fig.update_layout(yaxis=dict(tickprefix="R$ ", tickformat=",.0f"))
    return fig


CONFIG_PLOTLY = {"displaylogo": False, "scrollZoom": False,
                 "modeBarButtonsToRemove": ["lasso2d", "select2d", "autoScale2d", "toggleSpikelines"],
                 "toImageButtonOptions": {"format": "png", "scale": 2, "filename": "grafico_eficaz"}}


def _versao_streamlit():
    try:
        from importlib.metadata import version
        return tuple(int(p) for p in version("streamlit").split(".")[:2])
    except Exception:  # noqa: BLE001
        return (0, 0)


def mostrar_grafico(construir, *args, **kwargs):
    """Monta e exibe um gráfico Plotly ocupando toda a largura disponível."""
    if not PLOTLY_OK:
        st.info("📊 Para ver os gráficos interativos, feche a janela preta e abra o painel de novo: "
                "ele instala o Plotly sozinho.")
        return
    fig = construir(*args, **kwargs)
    if _versao_streamlit() >= (1, 50):  # versões novas trocaram use_container_width por width="stretch"
        st.plotly_chart(fig, width="stretch", config=CONFIG_PLOTLY)
    else:
        st.plotly_chart(fig, use_container_width=True, config=CONFIG_PLOTLY)


# --- 📊 Painel do Mês ----------------------------------------------------------
def grafico_cascata_dre(r: dict):
    """Cascata da formação do resultado: Receita Bruta até Lucro/Prejuízo Líquido."""
    abaixo_ebitda = r["lucro_liquido"] - r["ebitda"]  # depreciação + resultado financeiro + IR/CSLL
    passos = [
        ("Receita Bruta", "relative", r["receita_bruta"]),  # parte do zero; "relative" para sair verde
        ("(-) Impostos e Deduções", "relative", -r["impostos_vendas"]),
        ("Receita Líquida", "total", r["receita_liquida"]),
        ("(-) CMV / CSV", "relative", -r["custos_variaveis"]),
        ("Lucro Bruto", "total", r["lucro_bruto"]),
        ("(-) Desp. Comerciais Variáveis", "relative", -r["despesas_variaveis"]),
        ("Margem de Contribuição", "total", r["margem_contribuicao"]),
        ("(-) Custos Fixos", "relative", -r["custos_fixos"]),
        ("EBITDA", "total", r["ebitda"]),
        ("(+/-) Depreciação, Financ. e Impostos", "relative", abaixo_ebitda),
        ("Lucro Líquido" if r["lucro_liquido"] >= 0 else "Prejuízo Líquido", "total", r["lucro_liquido"]),
    ]
    rotulos = [p[0] for p in passos]
    textos = [brl(p[2]) for p in passos]
    percentuais = [pct(sobre_bruta(p[2], r)) for p in passos]
    fig = go.Figure(go.Waterfall(
        x=rotulos, measure=[p[1] for p in passos], y=[p[2] for p in passos],
        text=textos, textposition="outside", cliponaxis=False,
        textfont=dict(size=11, color=COR_TEXTO),
        customdata=list(zip(textos, percentuais)),
        hovertemplate="<b>%{x}</b><br>%{customdata[0]}<br>%{customdata[1]} da receita bruta<extra></extra>",
        increasing=dict(marker=dict(color=COR_VERDE)),
        decreasing=dict(marker=dict(color=COR_VERMELHO)),
        totals=dict(marker=dict(color=COR_MARINHO)),
        connector=dict(line=dict(color="#9AA8BA", width=1, dash="dot")),
    ))
    layout_marca(fig, altura=440, legenda=False)
    fig.update_xaxes(tickangle=-30, tickfont=dict(size=11))
    topo = max(r["receita_bruta"], 1.0)
    fundo = min(0.0, r["lucro_liquido"], r["ebitda"], r["margem_contribuicao"])
    fig.update_yaxes(range=[fundo * 1.25, topo * 1.15])
    return fig


def grafico_donut_faturamento(partes):
    """Para onde foi cada real faturado, em rosca (donut)."""
    df = pd.DataFrame({"Destino": [n for n, _, _ in partes], "valor": [v for _, v, _ in partes],
                       "Valor": [brl(v) for _, v, _ in partes], "Percentual": [pct(p) for _, _, p in partes]})
    fig = px.pie(df, names="Destino", values="valor", hole=0.4, color="Destino",
                 color_discrete_map={n: CORES_QUEBRA[i] for i, n in enumerate(df["Destino"])},
                 custom_data=["Valor", "Percentual"])
    fig.update_traces(sort=False, direction="clockwise", texttemplate=[pct(p) for _, _, p in partes],
                      textposition="inside",
                      insidetextorientation="horizontal", textfont=dict(size=13),
                      marker=dict(line=dict(color="#FFFFFF", width=2)),
                      hovertemplate="<b>%{label}</b><br>%{customdata[0][0]}<br>"
                                    "%{customdata[0][1]} do faturamento<extra></extra>")
    layout_marca(fig, altura=360, eixo_reais=False)
    fig.update_layout(legend=dict(orientation="v", y=0.5, yanchor="middle", x=1.02, xanchor="left"),
                      margin=dict(l=8, r=8, t=8, b=8))
    return fig


# --- 📈 Histórico ----------------------------------------------------------------
def grafico_evolucao(periodo: list, resultados: dict):
    """Área suave para Receita Líquida e linhas para Margem de Contribuição e Lucro Líquido."""
    x = [nome_mes(m) for m in periodo]

    def dados(chave):
        return [resultados[m][chave] for m in periodo]

    def extra(chave):
        return [[brl(resultados[m][chave]), pct(sobre_bruta(resultados[m][chave], resultados[m]))] for m in periodo]

    modelo = "%{customdata[0]} · %{customdata[1]} da receita bruta"
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=x, y=dados("receita_liquida"), name="Receita Líquida", mode="lines", line_shape="spline",
        line=dict(color=COR_CIANO_ESCURO, width=2), fill="tozeroy", fillcolor="rgba(24, 208, 232, 0.18)",
        customdata=extra("receita_liquida"), hovertemplate=modelo))
    fig.add_trace(go.Scatter(
        x=x, y=dados("margem_contribuicao"), name="Margem de Contribuição", mode="lines+markers",
        line_shape="spline", line=dict(color=COR_MARINHO, width=3),
        marker=dict(size=8, color=COR_MARINHO, line=dict(color="#FFFFFF", width=1.5)),
        customdata=extra("margem_contribuicao"), hovertemplate=modelo))
    lucro = dados("lucro_liquido")
    fig.add_trace(go.Scatter(
        x=x, y=lucro, name="Lucro Líquido", mode="lines+markers", line_shape="spline",
        line=dict(color=COR_VERDE, width=3, dash="dot"),
        marker=dict(size=9, color=[COR_VERDE if v >= 0 else COR_VERMELHO for v in lucro],
                    line=dict(color="#FFFFFF", width=1.5)),
        customdata=extra("lucro_liquido"), hovertemplate=modelo))
    layout_marca(fig, altura=420)
    fig.update_layout(hovermode="x unified")
    return fig


# --- 💸 Despesas -----------------------------------------------------------------
def grafico_pareto(por_categoria: pd.Series):
    """Curva ABC: barras decrescentes por categoria e linha do % acumulado (eixo da direita, 0–100%)."""
    serie = por_categoria[por_categoria > 0].sort_values(ascending=False)
    total = float(serie.sum()) or 1.0
    acumulado = serie.cumsum() / total * 100
    classe = ["A" if a <= 80 or i == 0 else ("B" if a <= 95 else "C") for i, a in enumerate(acumulado)]
    cores_classe = {"A": COR_MARINHO, "B": COR_MARINHO_MEDIO, "C": "#8FA3C0"}
    nomes = list(serie.index)
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=nomes, y=serie.values, name="Despesa no mês",
        marker=dict(color=[cores_classe[c] for c in classe], cornerradius=4),
        customdata=[[brl(v), pct(v / total * 100), c] for v, c in zip(serie.values, classe)],
        hovertemplate="<b>%{x}</b><br>%{customdata[0]} · %{customdata[1]} do total<br>"
                      "Classe %{customdata[2]}<extra></extra>"))
    fig.add_trace(go.Scatter(
        x=nomes, y=acumulado.values, name="% acumulado", yaxis="y2", mode="lines+markers",
        line=dict(color=COR_CIANO_ESCURO, width=2.5),
        marker=dict(size=7, color=COR_CIANO, line=dict(color="#FFFFFF", width=1)),
        customdata=[pct(a) for a in acumulado.values],
        hovertemplate="<b>%{x}</b><br>Acumulado: %{customdata}<extra></extra>"))
    layout_marca(fig, altura=430)
    fig.update_layout(
        yaxis2=dict(overlaying="y", side="right", range=[0, 105], ticksuffix="%", showgrid=False,
                    tickfont=dict(color=COR_CIANO_ESCURO), title=None),
        hovermode="x unified", bargap=0.25)
    fig.add_hline(y=80, yref="y2", line=dict(color=COR_CIANO_ESCURO, width=1, dash="dash"),
                  annotation_text="80%", annotation_position="top left",
                  annotation_font=dict(color=COR_CIANO_ESCURO, size=11))
    fig.update_xaxes(tickangle=-35, tickfont=dict(size=11))
    return fig


def grafico_despesas_mes_a_mes(desp: pd.DataFrame, meses_grafico: list, cores: dict):
    """Barras empilhadas por grupo/natureza (os 8 maiores com cor própria, o resto em 'Outros')."""
    df = (desp[desp["competencia"].isin(meses_grafico)]
          .groupby(["competencia", "grupo"], as_index=False)["valor"].sum())
    df["grupo"] = df["grupo"].where(df["grupo"].map(cores) != COR_OUTROS, "Outros")
    df = df.groupby(["competencia", "grupo"], as_index=False)["valor"].sum()
    ordem = [g for g, c in cores.items() if c != COR_OUTROS] + (["Outros"] if (df["grupo"] == "Outros").any() else [])
    x = [nome_mes(m) for m in meses_grafico]
    fig = go.Figure()
    for grupo in ordem:
        parte = df[df["grupo"] == grupo].set_index("competencia")["valor"].reindex(meses_grafico, fill_value=0)
        fig.add_trace(go.Bar(x=x, y=parte.values, name=grupo, marker=dict(color=cores.get(grupo, COR_OUTROS)),
                             customdata=[brl(v) for v in parte.values],
                             hovertemplate=f"{grupo}: %{{customdata}}<extra></extra>"))
    totais = df.groupby("competencia")["valor"].sum().reindex(meses_grafico, fill_value=0)
    fig.add_trace(go.Scatter(x=x, y=totais.values, mode="text", text=[brl(v) for v in totais.values],
                             textposition="top center", textfont=dict(size=11, color=COR_TEXTO_SECUNDARIO),
                             showlegend=False, hoverinfo="skip"))
    layout_marca(fig, altura=400)
    fig.update_layout(barmode="stack", hovermode="x unified", bargap=0.3, legend=dict(traceorder="normal"))
    fig.update_traces(marker_line=dict(color="#FFFFFF", width=1), selector=dict(type="bar"))
    fig.update_yaxes(range=[0, float(totais.max() or 1) * 1.15])
    return fig


# --- 💧 Fluxo de Caixa -------------------------------------------------------------
def grafico_caixa_diario_plotly(extrato: pd.DataFrame):
    """Barras de entradas (ciano) e saídas (marinho) por dia + linha do saldo acumulado."""
    dias = extrato["dia"]
    fig = go.Figure()
    fig.add_trace(go.Bar(x=dias, y=extrato["entradas"], name="Entradas", marker=dict(color=COR_CIANO),
                         customdata=[brl(v) for v in extrato["entradas"]],
                         hovertemplate="Entradas: %{customdata}<extra></extra>"))
    fig.add_trace(go.Bar(x=dias, y=-extrato["saidas"], name="Saídas", marker=dict(color=COR_MARINHO_MEDIO),
                         customdata=[brl(v) for v in extrato["saidas"]],
                         hovertemplate="Saídas: %{customdata}<extra></extra>"))
    saldo = extrato["saldo"]
    fig.add_trace(go.Scatter(
        x=dias, y=saldo, name="Saldo acumulado", mode="lines+markers", line=dict(color=COR_MARINHO, width=3),
        marker=dict(size=6, color=[COR_MARINHO if v >= 0 else COR_VERMELHO for v in saldo],
                    line=dict(color="#FFFFFF", width=1)),
        customdata=[brl(v) for v in saldo], hovertemplate="<b>Saldo: %{customdata}</b><extra></extra>"))
    layout_marca(fig, altura=400)
    fig.update_layout(barmode="relative", hovermode="x unified", bargap=0.25)
    fig.update_xaxes(tickformat="%d/%m", dtick=86400000 * 2, hoverformat="%d/%m/%Y")
    return fig


# --- 🎯 Orçado x Realizado --------------------------------------------------------
def grafico_orcado_realizado(comp: list):
    """Barras horizontais lado a lado (Orçado x Realizado) por linha da DRE, com o desvio no rótulo."""
    itens = [c for c in comp if c["orcado"] or c["realizado"]][::-1]  # de cima para baixo na ordem da DRE
    rotulos = [c["rotulo"] for c in itens]

    def desvio_txt(c):
        p = "" if c["dif_pct"] is None else f" ({'+' if c['dif_pct'] > 0 else ''}{pct(c['dif_pct'])})"
        return ("+" if c["dif"] > 0 else "") + brl(c["dif"]) + p

    favoravel = [(c["dif"] >= 0) == (c["chave"] in FAVORAVEL_SE_MAIOR) or abs(c["dif"]) < 0.005 for c in itens]
    dados_hover = [[brl(c["orcado"]), brl(c["realizado"]), desvio_txt(c), c["status"]] for c in itens]
    modelo = ("<b>%{y}</b><br>Orçado: %{customdata[0]}<br>Realizado: %{customdata[1]}<br>"
              "Desvio: %{customdata[2]}<br>%{customdata[3]}<extra></extra>")
    fig = go.Figure()
    fig.add_trace(go.Bar(y=rotulos, x=[c["orcado"] for c in itens], name="Orçado", orientation="h",
                         marker=dict(color=COR_CIANO), customdata=dados_hover, hovertemplate=modelo))
    fig.add_trace(go.Bar(y=rotulos, x=[c["realizado"] for c in itens], name="Realizado", orientation="h",
                         marker=dict(color=COR_MARINHO), customdata=dados_hover, hovertemplate=modelo,
                         text=[("▲ " if f else "▼ ") + (pct_com_sinal(c["dif_pct"]) if c["dif_pct"] is not None
                                                        else com_sinal(c["dif"])) for c, f in zip(itens, favoravel)],
                         textposition="outside", cliponaxis=False,
                         textfont=dict(size=11, color=[COR_VERDE if f else COR_VERMELHO for f in favoravel])))
    layout_marca(fig, altura=max(360, 46 * len(itens)), eixo_reais=False)
    fig.update_layout(barmode="group", bargap=0.28, bargroupgap=0.08,
                      legend=dict(orientation="h", y=1.02, yanchor="bottom", x=0))
    fig.update_xaxes(tickprefix="R$ ", tickformat=",.0f", showgrid=True, gridcolor="#E3E9F1")
    fig.update_yaxes(gridcolor="rgba(0,0,0,0)", tickfont=dict(size=12, color=COR_TEXTO))
    maximo = max([max(c["orcado"], c["realizado"]) for c in itens] + [1.0])
    minimo = min([min(c["orcado"], c["realizado"]) for c in itens] + [0.0])
    fig.update_xaxes(range=[minimo * 1.2, maximo * 1.22])
    return fig


def gerar_pdf(cliente_nome, competencia, r, ra, anterior, comentario, comparacao=None) -> bytes:
    """Resumo executivo do mês: KPIs, comentário do analista, quebra do faturamento,
    DRE resumida e (se houver) orçado x realizado."""
    from xml.sax.saxutils import escape

    from reportlab.lib import colors
    from reportlab.lib.enums import TA_LEFT
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.graphics.shapes import Drawing, Rect, String
    from reportlab.platypus import (HRFlowable, Image, KeepTogether, Paragraph, SimpleDocTemplate, Spacer,
                                    Table, TableStyle)

    AZUL = colors.HexColor(COR_MARINHO)
    CIANO = colors.HexColor(COR_CIANO)
    CIANO_ESCURO = colors.HexColor(COR_CIANO_ESCURO)
    CINZA = colors.HexColor(COR_TEXTO_SECUNDARIO)
    LINHA = colors.HexColor(COR_BORDA)
    FUNDO = colors.HexColor("#F1F5FA")
    VERDE = colors.HexColor("#0f7a3a")
    VERMELHO = colors.HexColor("#b42318")

    def estilo(nome, **kw):
        base = dict(fontName="Helvetica", fontSize=9, leading=12, textColor=colors.black, alignment=TA_LEFT)
        base.update(kw)
        return ParagraphStyle(nome, **base)

    titulo = estilo("t", fontName="Helvetica-Bold", fontSize=18, leading=22, textColor=AZUL)
    subtitulo = estilo("st", fontSize=10, textColor=CINZA)
    secao = estilo("s", fontName="Helvetica-Bold", fontSize=12, leading=15, textColor=AZUL, spaceBefore=10, spaceAfter=6)
    rotulo = estilo("r", fontSize=8, leading=10, textColor=CINZA)
    valor = estilo("v", fontName="Helvetica-Bold", fontSize=13, leading=16)
    pequeno = estilo("p", fontSize=8, leading=10, textColor=CINZA)
    normal = estilo("n", fontSize=9.5, leading=13)
    celula = estilo("c", fontSize=8.5, leading=11)
    celula_negrito = estilo("cn", fontName="Helvetica-Bold", fontSize=8.5, leading=11)

    largura = A4[0] - 3 * cm

    # Cabeçalho: logo à esquerda, título à direita, fio ciano embaixo
    try:
        from PIL import Image as PILImage
        logo = PILImage.open(io.BytesIO(logo_bytes("completo"))).convert("RGBA")
        fundo_branco = PILImage.new("RGBA", logo.size, (255, 255, 255, 255))
        fundo_branco.alpha_composite(logo)
        arquivo_logo = io.BytesIO()
        fundo_branco.convert("RGB").save(arquivo_logo, format="PNG")
        arquivo_logo.seek(0)
        largura_logo = 5.0 * cm
        imagem_logo = Image(arquivo_logo, width=largura_logo, height=largura_logo * logo.height / logo.width)
    except Exception:  # noqa: BLE001 - sem logo o PDF sai só com o nome
        imagem_logo = Paragraph(f"<b>{MARCA_NOME}</b>", estilo("m", fontSize=12, textColor=AZUL))
    bloco_titulo = [
        Paragraph(f"Resumo Executivo – {escape(cliente_nome)}", estilo("t2", fontName="Helvetica-Bold",
                                                                        fontSize=16, leading=20, textColor=AZUL,
                                                                        alignment=2)),
        Paragraph(f"Competência: {nome_mes(competencia)} &nbsp;·&nbsp; Emitido em {datetime.now():%d/%m/%Y}",
                  estilo("st2", fontSize=9.5, textColor=CINZA, alignment=2)),
    ]
    cabecalho = Table([[imagem_logo, bloco_titulo]], colWidths=[5.6 * cm, largura - 5.6 * cm])
    cabecalho.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                                   ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0)]))
    historia = [
        cabecalho,
        Spacer(1, 6),
        HRFlowable(width="100%", thickness=2.2, color=CIANO, spaceBefore=2, spaceAfter=12),
    ]

    # --- Cards de KPIs -------------------------------------------------------
    def variacao(chave, inverso=False):
        if not ra:
            return ""
        dif = r[chave] - ra[chave]
        bom = dif <= 0 if inverso else dif >= 0
        cor = "#0f7a3a" if bom else "#b42318"
        sinal = "+" if dif > 0 else ""
        return f'<font color="{cor}">{sinal}{brl(dif)}</font> vs {nome_mes(anterior)}'

    kpis = [
        ("Receita Bruta", brl(r["receita_bruta"]), variacao("receita_bruta")),
        ("Receita Líquida", brl(r["receita_liquida"]), variacao("receita_liquida")),
        ("Margem de Contribuição", brl(r["margem_contribuicao"]), f"{pct(r['perc_mc'])} da receita bruta"),
        ("EBITDA", brl(r["ebitda"]), variacao("ebitda")),
        ("Lucro Líquido", brl(r["lucro_liquido"]), variacao("lucro_liquido")),
        ("Margem Líquida", pct(r["margem_liquida"]), "da receita bruta"),
        ("Geração de Caixa", brl(r["geracao_caixa"]), variacao("geracao_caixa")),
        ("Ponto de Equilíbrio", brl(r["peq"]), variacao("peq", inverso=True)),
    ]
    cards = [[Paragraph(escape(n), rotulo), Paragraph(v, valor), Paragraph(d or "&nbsp;", pequeno)]
             for n, v, d in kpis]
    grade = [cards[0:4], cards[4:8]]
    tabela_kpi = Table(grade, colWidths=[largura / 4] * 4)
    tabela_kpi.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), FUNDO),
        ("BOX", (0, 0), (-1, -1), 0.5, LINHA),
        ("INNERGRID", (0, 0), (-1, -1), 3, colors.white),
        ("LINEBEFORE", (0, 0), (-1, -1), 2.5, CIANO),
        ("LEFTPADDING", (0, 0), (-1, -1), 9),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 7), ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
    ]))
    historia.append(tabela_kpi)

    # --- Comentário do analista -----------------------------------------------
    if comentario.strip():
        texto = escape(comentario.strip()).replace("\n", "<br/>")
        caixa = Table([[Paragraph("<b>Comentário do Analista Financeiro</b>", estilo("ca", textColor=AZUL))],
                       [Paragraph(texto, normal)]], colWidths=[largura])
        caixa.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor(COR_CIANO_SUAVE)),
            ("LINEBEFORE", (0, 0), (0, -1), 3, CIANO_ESCURO),
            ("LEFTPADDING", (0, 0), (-1, -1), 10), ("TOPPADDING", (0, 0), (-1, -1), 6),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ]))
        historia += [Spacer(1, 12), caixa]

    # --- Para onde foi o faturamento ------------------------------------------
    partes, aviso_quebra = quebra_faturamento(r)
    if partes:
        total = sum(v for _, v, _ in partes) or 1
        desenho = Drawing(largura, 34)
        x = 0.0
        for i, (nome, v, perc) in enumerate(partes):
            w = largura * v / total
            if w <= 0:
                continue
            desenho.add(Rect(x, 4, max(w - 2, 0.5), 26, fillColor=colors.HexColor(CORES_QUEBRA[i]),
                             strokeColor=None))
            if perc >= 6:
                desenho.add(String(x + w / 2, 13, pct(perc), textAnchor="middle", fontName="Helvetica-Bold",
                                   fontSize=9, fillColor=colors.HexColor(TEXTO_SOBRE_COR[i])))
            x += w
        legenda = [[" ", Paragraph(escape(nome), celula), Paragraph(brl(v), celula), Paragraph(pct(p), celula)]
                   for nome, v, p in partes]
        tabela_legenda = Table(legenda, colWidths=[0.45 * cm, 7 * cm, 3.5 * cm, 2.2 * cm], hAlign="LEFT")
        estilo_legenda = [("TOPPADDING", (0, 0), (-1, -1), 2), ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
                          ("VALIGN", (0, 0), (-1, -1), "MIDDLE")]
        for i in range(len(partes)):
            estilo_legenda.append(("BACKGROUND", (0, i), (0, i), colors.HexColor(CORES_QUEBRA[i])))
        tabela_legenda.setStyle(TableStyle(estilo_legenda))
        bloco = [Paragraph("Para onde foi o faturamento", secao),
                 Paragraph(f"De cada R$ 100,00 faturados em {nome_mes(competencia)}:", pequeno),
                 Spacer(1, 4), desenho, Spacer(1, 4), tabela_legenda]
        if aviso_quebra:
            bloco += [Spacer(1, 4), Paragraph(f'<font color="#b42318">{escape(aviso_quebra)}</font>', pequeno)]
        historia.append(KeepTogether(bloco))

    # --- DRE resumida ----------------------------------------------------------
    cab = ["Conta", "Valor", "% Rec. Bruta"] + ([nome_mes(anterior)] if ra else [])
    dados = [[Paragraph(f"<b>{c}</b>", celula) for c in cab]]
    destaques = []
    for i, (rot, k) in enumerate(ESTRUTURA_DRE, start=1):
        forte = rot.startswith("(=)") or rot.startswith("1.")
        st_cel = celula_negrito if forte else celula
        linha = [Paragraph(escape(rot), st_cel), Paragraph(brl(r[k]), st_cel),
                 Paragraph(pct(sobre_bruta(r[k], r)), st_cel)]
        if ra:
            linha.append(Paragraph(brl(ra[k]), st_cel))
        dados.append(linha)
        if forte:
            destaques.append(i)
    larguras = [8.6 * cm, 3.4 * cm, 2.4 * cm] + ([3.4 * cm] if ra else [])
    tabela_dre = Table(dados, colWidths=larguras, repeatRows=1)
    estilo_dre = [
        ("LINEBELOW", (0, 0), (-1, 0), 1, AZUL),
        ("LINEBELOW", (0, 1), (-1, -1), 0.25, LINHA),
        ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
        ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]
    for i in destaques:
        estilo_dre.append(("BACKGROUND", (0, i), (-1, i), FUNDO))
    tabela_dre.setStyle(TableStyle(estilo_dre))
    historia.append(KeepTogether([Paragraph("DRE Gerencial – resumo", secao), tabela_dre]))

    # --- Orçado x Realizado ---------------------------------------------------
    if comparacao:
        cab = ["Linha", "Orçado", "Realizado", "Diferença", "Dif. %", "Situação"]
        dados = [[Paragraph(f"<b>{c}</b>", celula) for c in cab]]
        estilo_orc = [
            ("LINEBELOW", (0, 0), (-1, 0), 1, AZUL),
            ("LINEBELOW", (0, 1), (-1, -1), 0.25, LINHA),
            ("ALIGN", (1, 0), (-2, -1), "RIGHT"),
            ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ]
        textos_status = {0: ("OK", VERDE), 1: ("Atenção", colors.HexColor("#8a5a00")), 2: ("Alerta", VERMELHO)}
        n = 0
        for item in comparacao:
            if not item["orcado"] and not item["realizado"]:
                continue
            n += 1
            st_cel = celula_negrito if item["total"] else celula
            texto, cor = textos_status[item["nivel"]]
            if item["nivel"] == 2:
                texto = "Abaixo do orçado" if item["chave"] in FAVORAVEL_SE_MAIOR else "Estouro"
            dados.append([
                Paragraph(escape(item["rotulo"]), st_cel), Paragraph(brl(item["orcado"]), st_cel),
                Paragraph(brl(item["realizado"]), st_cel),
                Paragraph(("+" if item["dif"] > 0 else "") + brl(item["dif"]), st_cel),
                Paragraph(("+" if (item["dif_pct"] or 0) > 0 else "") + pct(item["dif_pct"])
                          if item["dif_pct"] is not None else "–", st_cel),
                Paragraph(f'<font color="{cor.hexval().replace("0x", "#")}"><b>{texto}</b></font>', celula),
            ])
            if item["nivel"] == 2:
                estilo_orc.append(("BACKGROUND", (0, n), (-1, n), colors.HexColor("#fde2e1")))
            elif item["nivel"] == 1:
                estilo_orc.append(("BACKGROUND", (0, n), (-1, n), colors.HexColor("#fff4d6")))
        tabela_orc = Table(dados, colWidths=[5.6 * cm, 2.6 * cm, 2.6 * cm, 2.6 * cm, 1.5 * cm, 3.1 * cm],
                           repeatRows=1)
        tabela_orc.setStyle(TableStyle(estilo_orc))
        historia.append(KeepTogether([Paragraph("Orçado x Realizado", secao), tabela_orc]))

    def rodape(canvas, doc):
        canvas.saveState()
        canvas.setStrokeColor(CIANO)
        canvas.setLineWidth(1)
        canvas.line(1.5 * cm, 1.35 * cm, A4[0] - 1.5 * cm, 1.35 * cm)
        canvas.setFont("Helvetica-Bold", 7.5)
        canvas.setFillColor(AZUL)
        canvas.drawString(1.5 * cm, 1 * cm, MARCA_NOME)
        canvas.setFont("Helvetica", 7.5)
        canvas.setFillColor(CINZA)
        canvas.drawString(1.5 * cm + canvas.stringWidth(MARCA_NOME, "Helvetica-Bold", 7.5) + 6, 1 * cm,
                          f"·  {cliente_nome}  ·  Resumo Executivo {nome_mes(competencia)}")
        canvas.drawRightString(A4[0] - 1.5 * cm, 1 * cm, f"Página {doc.page}")
        canvas.restoreState()

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, leftMargin=1.5 * cm, rightMargin=1.5 * cm,
                            topMargin=1.4 * cm, bottomMargin=1.6 * cm,
                            title=f"Resumo Executivo {cliente_nome} {nome_mes(competencia)}",
                            author=MARCA_NOME, creator=MARCA_NOME)
    doc.build(historia, onFirstPage=rodape, onLaterPages=rodape)
    return buffer.getvalue()

# =============================================================================
#  6. TELAS DO PAINEL
# =============================================================================

try:
    from PIL import Image as _PILImage
    _icone_aba = _PILImage.open(io.BytesIO(logo_bytes("icone")))
except Exception:  # noqa: BLE001
    _icone_aba = "🔷"
st.set_page_config(page_title=MARCA_NOME, page_icon=_icone_aba, layout="wide")

# Acabamento visual da marca (complementa o tema marinho/ciano que o lançador aplica)
st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Montserrat:wght@600;700&display=swap');
html, body, .stApp, .stMarkdown, [data-testid="stMetricValue"], [data-testid="stMetricLabel"] {{
    font-family: 'Inter', 'Segoe UI', sans-serif;
}}
.stApp h1, .stApp h2, .stApp h3 {{
    font-family: 'Montserrat', 'Segoe UI', sans-serif; color: {COR_MARINHO}; letter-spacing: -0.01em;
}}
.stApp h1 {{ font-weight: 700; border-bottom: 3px solid {COR_CIANO}; padding-bottom: .4rem; margin-bottom: 1rem; }}
.stApp h2, .stApp h3 {{ font-weight: 600; }}
[data-testid="stMetric"] {{
    background: #FFFFFF; border: 1px solid {COR_BORDA}; border-left: 4px solid {COR_CIANO};
    border-radius: 10px; padding: 12px 16px; box-shadow: 0 1px 3px rgba(16, 40, 80, .06);
}}
[data-testid="stMetricLabel"] p {{ color: {COR_TEXTO_SECUNDARIO}; font-weight: 600; }}
[data-testid="stMetricValue"] {{ color: {COR_MARINHO}; font-weight: 700; }}
[data-testid="stSidebar"] {{ background: {COR_FUNDO_LATERAL}; border-right: 1px solid {COR_BORDA}; }}
button[kind="primary"], [data-testid="stBaseButton-primary"] {{
    background: {COR_MARINHO}; border-color: {COR_MARINHO}; color: #FFFFFF;
}}
button[kind="primary"]:hover, [data-testid="stBaseButton-primary"]:hover {{
    background: {COR_CIANO_ESCURO}; border-color: {COR_CIANO_ESCURO}; color: #FFFFFF;
}}
button[kind="secondary"]:hover, [data-testid="stBaseButton-secondary"]:hover {{
    border-color: {COR_CIANO_ESCURO}; color: {COR_MARINHO};
}}
[data-testid="stDataFrame"] {{ border: 1px solid {COR_BORDA}; border-radius: 8px; }}
[data-testid="stExpander"] details {{ border-color: {COR_BORDA}; border-radius: 10px; background: #FFFFFF; }}
.stTabs [data-baseweb="tab-highlight"] {{ background-color: {COR_CIANO_ESCURO}; }}
hr {{ border-color: {COR_BORDA}; }}
.marca-lateral {{
    background: #FFFFFF; border: 1px solid {COR_BORDA}; border-bottom: 3px solid {COR_CIANO};
    border-radius: 12px; padding: 14px 14px 10px; margin: -0.5rem 0 1rem 0; text-align: center;
}}
.marca-lateral img {{ width: 100%; max-width: 230px; height: auto; display: block; margin: 0 auto; }}
.marca-lateral .marca-legenda {{
    color: {COR_TEXTO_SECUNDARIO}; font-size: .72rem; letter-spacing: .12em; text-transform: uppercase;
    margin-top: 6px; font-weight: 600;
}}
.marca-rodape {{ color: {COR_TEXTO_SECUNDARIO}; font-size: .72rem; margin-top: 1.5rem; text-align: center; }}
</style>
""", unsafe_allow_html=True)

PACOTES = {
    "Nível 1: Essencial": 80,
    "Nível 2: Gerencial": 180,
    "Nível 3: Estratégico (FP&A)": 300,
}


@st.cache_resource
def conexao():
    return conectar()


con = conexao()

# ---------------------------------------------------------------------------
# Barra lateral: cliente e navegação
# ---------------------------------------------------------------------------
st.sidebar.markdown(
    f'<div class="marca-lateral"><img src="{logo_data_uri("completo")}" alt="{MARCA_NOME}">'
    '<div class="marca-legenda">Painel Financeiro</div></div>',
    unsafe_allow_html=True,
)
clientes = listar_clientes(con)

with st.sidebar.expander("➕ Cadastrar cliente", expanded=clientes.empty):
    novo_nome = st.text_input("Nome do cliente")
    novo_pacote = st.selectbox("Pacote contratado", list(PACOTES))
    if st.button("Cadastrar"):
        if not novo_nome.strip():
            st.error("Informe o nome.")
        elif novo_nome.strip().lower() in clientes["nome"].str.lower().tolist():
            st.error("Já existe um cliente com esse nome.")
        else:
            criar_cliente(con, novo_nome, novo_pacote)
            st.rerun()

if clientes.empty:
    st.title(f"{MARCA_NOME} – Painel Financeiro")
    st.info("Cadastre o primeiro cliente na barra lateral para começar.")
    st.stop()

nome_cliente = st.sidebar.selectbox("Cliente", clientes["nome"].tolist())
cliente = clientes[clientes["nome"] == nome_cliente].iloc[0]
cliente_id = int(cliente["id"])
if cliente["pacote"]:
    st.sidebar.caption(f"Pacote: {cliente['pacote']}")

pagina = st.sidebar.radio(
    "Navegação",
    ["📊 Painel do Mês", "📈 Histórico", "💸 Despesas", "💧 Fluxo de Caixa", "🎯 Orçado x Realizado",
     "📥 Importar Lançamentos", "🗂️ Plano de Contas"],
)
st.sidebar.markdown(f'<div class="marca-rodape">© {datetime.now().year} {MARCA_NOME}</div>', unsafe_allow_html=True)

meses = competencias(con, cliente_id)
tabela_meses = valores_por_mes(con, cliente_id)


def resultado_do_mes(competencia):
    valores = tabela_meses.loc[competencia].to_dict() if competencia in tabela_meses.index else {}
    return calcular(valores, obter_funcionarios(con, cliente_id, competencia))


def aviso_sem_classificacao(competencia=None):
    totais = totais_por_categoria(con, cliente_id, competencia)
    pendentes = totais[totais["linha"].isna()]
    if not pendentes.empty:
        st.warning(
            f"⚠️ {len(pendentes)} categoria(s) ainda sem classificação na DRE "
            f"({int(pendentes['lancamentos'].sum())} lançamentos) e fora dos números abaixo. "
            "Classifique em **🗂️ Plano de Contas**."
        )


def _controle_do_streamlit(erro) -> bool:
    """st.stop() e st.rerun() funcionam por exceções internas: elas não podem ser tratadas como erro."""
    return any(c.__name__ in ("StopException", "RerunException", "ScriptControlException")
               for c in type(erro).__mro__)


def aviso_nova_estrutura():
    movidas = categorias_movidas_na_migracao(con, cliente_id)
    if not movidas:
        return
    with st.container(border=True):
        st.markdown(
            "**ℹ️ Nova estrutura da DRE.** As despesas que nascem de cada venda agora têm linha própria, "
            "**Despesas Comerciais Variáveis**, entre o novo **Lucro Bruto** e a **Margem de Contribuição**. "
            f"Estas categorias deste cliente foram movidas para ela: **{', '.join(movidas)}**. "
            "A Margem de Contribuição e o Ponto de Equilíbrio não mudam; se alguma não for despesa de venda, "
            "ajuste no **🗂️ Plano de Contas**. Se houver orçamento cadastrado, revise a linha de CMV, que antes "
            "incluía essas despesas.".replace("$", "\\$"))
        if st.button("Entendi", key=f"ok_migracao_{cliente_id}"):
            dispensar_aviso(con, cliente_id, MIGRACAO_DESPESAS_VARIAVEIS)
            st.rerun()


def precisa_de_dados():
    st.info("Nenhum lançamento importado para este cliente ainda. Use **📥 Importar Lançamentos**.")
    st.stop()


# ---------------------------------------------------------------------------
# 📊 Painel do Mês
# ---------------------------------------------------------------------------
if pagina == "📊 Painel do Mês":
    st.title(f"📊 {nome_cliente}")
    if not meses:
        precisa_de_dados()

    topo_mes, topo_pdf = st.columns([3, 1])
    competencia = topo_mes.selectbox("Mês", meses[::-1], format_func=nome_mes)
    anterior = meses[meses.index(competencia) - 1] if meses.index(competencia) > 0 else None

    r = resultado_do_mes(competencia)
    ra = resultado_do_mes(anterior) if anterior else None
    ano_passado = mes_seguinte(competencia, -12)  # 2026-09 -> 2025-09
    ry = resultado_do_mes(ano_passado) if ano_passado in meses else None
    comentario = carregar_comentario(con, cliente_id, competencia)
    orcado_mes = carregar_orcamento(con, cliente_id, competencia)
    comparacao = comparar_orcamento(calcular(orcado_mes), r) if orcado_mes else None

    try:
        pdf = gerar_pdf(nome_cliente, competencia, r, ra, anterior, comentario, comparacao)
        topo_pdf.write("")  # alinha o botão com a caixa de seleção
        topo_pdf.download_button(
            "📄 Baixar Resumo Executivo em PDF", pdf, type="primary",
            file_name=f"Resumo Executivo - {nome_cliente} - {nome_mes(competencia).replace('/', '-')}.pdf",
            mime="application/pdf",
        )
    except ImportError:
        topo_pdf.caption("Para gerar o PDF, feche e abra o painel de novo (ele instala o que falta).")

    st.markdown(f"### DRE Gerencial – {nome_mes(competencia)}")
    aviso_nova_estrutura()
    aviso_sem_classificacao(competencia)
    problemas_cascata = verificar_cascata(r)
    if problemas_cascata:  # não deve acontecer nunca: é a trava de segurança dos cálculos
        st.error("❌ Inconsistência no cálculo da DRE. Não apresente este mês ao cliente e envie esta tela ao "
                 "suporte: " + "; ".join(problemas_cascata))
    invertidas = saldos_invertidos(r)
    if invertidas:
        texto_aviso = (
            "⚠️ **Saldo invertido em " + nome_mes(competencia) + ":** " + "; ".join(
                f"{LINHAS[k].split(' – ')[0]} ficou {'negativa' if k in LINHAS_CREDORAS else 'credora'} "
                f"em {brl(-r[k])}" for k in invertidas)
            + ". Isso acontece quando estornos/devoluções superam os lançamentos do mês. A DRE considera o valor "
              "líquido (ex.: estorno reduz a despesa). Se não houve estorno de verdade, confira o sinal do arquivo "
              "importado ou a classificação da categoria no Plano de Contas.")
        st.warning(texto_aviso.replace("$", "\\$"))
    if comentario:
        # "$" vira fórmula no Markdown do Streamlit (ex.: "R$ 1.000 e R$ 2.000"); quebras de linha precisam de "  \n"
        texto_md = comentario.replace("\\", "\\\\").replace("$", "\\$").replace("<", "&lt;").replace("\n", "  \n")
        st.info(f"💬 **Comentário do Analista:**  \n{texto_md}")
    if comparacao:
        criticos = [c for c in comparacao
                    if c["nivel"] == 2 and (not c["total"] or c["chave"] == "receita_bruta")]
        estouros = [c for c in criticos if c["chave"] not in FAVORAVEL_SE_MAIOR]
        receitas_baixas = [c for c in criticos if c["chave"] in FAVORAVEL_SE_MAIOR]
        if estouros:
            st.error(
                "🔴 **Estouro do orçamento:** " + "; ".join(
                    f"{c['rotulo'].split(') ')[-1]} (+{brl(c['dif'])})" for c in estouros)
                + ". Detalhes em **🎯 Orçado x Realizado**."
            )
        if receitas_baixas:
            st.warning(
                "🔻 **Receita abaixo do orçado:** " + "; ".join(
                    f"{c['rotulo'].split(') ')[-1]} ({brl(c['dif'])})" for c in receitas_baixas)
            )

    def variacao(chave):
        return brl(r[chave] - ra[chave]) if ra else None

    ajuda_var = f"Variação em relação a {nome_mes(anterior)}." if anterior else None

    col1, col2, col3, col4 = st.columns(4)
    for col, titulo, chave in [
        (col1, "Receita Líquida", "receita_liquida"),
        (col2, "Margem de Contribuição", "margem_contribuicao"),
        (col3, "EBITDA", "ebitda"),
        (col4, "Lucro Líquido", "lucro_liquido"),
    ]:
        col.metric(titulo, brl(r[chave]), variacao(chave), help=ajuda_var)
        col.caption(f"{pct(sobre_bruta(r[chave], r))} da receita bruta")

    st.subheader("🌊 Formação do resultado do mês")
    st.caption("Da receita bruta ao lucro: verde soma, vermelho subtrai, azul-marinho são os subtotais da DRE. "
               "Passe o mouse para ver valor e % da receita bruta.")
    mostrar_grafico(grafico_cascata_dre, r)

    st.subheader("🧩 Para onde foi o faturamento")
    partes, aviso_quebra = quebra_faturamento(r)
    if partes:
        st.caption(f"De cada R$ 100,00 faturados em {nome_mes(competencia)}:")
        rosca, tabela_rosca = st.columns([3, 2])
        with rosca:
            mostrar_grafico(grafico_donut_faturamento, partes)
        tabela_rosca.write("")
        tabela_rosca.dataframe(pd.DataFrame({
            "Destino": [n for n, _, _ in partes],
            "Valor": [brl(v) for _, v, _ in partes],
            "% do faturamento": [pct(p) for _, _, p in partes],
        }), hide_index=True)
        st.caption("Margem de contribuição = faturamento − impostos − CMV/CSV − despesas comerciais variáveis "
                   "= custos fixos + sobra operacional.")
    if aviso_quebra:
        st.warning(aviso_quebra)

    st.markdown("---")
    st.subheader("📋 DRE Gerencial Estruturada")
    dre = {
        "Conta / Indicador": [rot for rot, _ in ESTRUTURA_DRE],
        "Valor": [brl(r[k]) for _, k in ESTRUTURA_DRE],
        "% da Receita Bruta": [pct(sobre_bruta(r[k], r)) for _, k in ESTRUTURA_DRE],
    }
    if ra:
        dre[f"Valor {nome_mes(anterior)}"] = [brl(ra[k]) for _, k in ESTRUTURA_DRE]
        dre[f"% {nome_mes(anterior)}"] = [pct(sobre_bruta(ra[k], ra)) for _, k in ESTRUTURA_DRE]
    if ry:
        dre[f"Valor {nome_mes(ano_passado)}"] = [brl(ry[k]) for _, k in ESTRUTURA_DRE]
        dre["Var. % ano anterior"] = [pct_com_sinal(variacao_pct(r[k], ry[k])) for _, k in ESTRUTURA_DRE]
    st.dataframe(pd.DataFrame(dre), hide_index=True)

    st.subheader("💰 Do Lucro ao Caixa")
    st.caption(
        "Amortização de empréstimos, investimentos e distribuição de lucros não são despesas, "
        "por isso ficam fora da DRE. Mas saem do caixa, e é aqui que aparecem."
    )
    caixa = {"Item": [rot for rot, _ in ESTRUTURA_CAIXA], "Valor": [brl(r[k]) for _, k in ESTRUTURA_CAIXA]}
    if ra:
        caixa[f"Valor {nome_mes(anterior)}"] = [brl(ra[k]) for _, k in ESTRUTURA_CAIXA]
    if ry:
        caixa[f"Valor {nome_mes(ano_passado)}"] = [brl(ry[k]) for _, k in ESTRUTURA_CAIXA]
    st.dataframe(pd.DataFrame(caixa), hide_index=True)

    # -----------------------------------------------------------------------
    # 📅 Comparativo com o mesmo mês do ano anterior (YoY)
    # -----------------------------------------------------------------------
    st.markdown("---")
    st.subheader(f"📅 Mesmo mês do ano anterior: {nome_mes(competencia)} x {nome_mes(ano_passado)}")
    if not ry:
        st.info(
            f"Ainda não há lançamentos de **{nome_mes(ano_passado)}** importados para este cliente. Importe o "
            "relatório do ERP desse mês em **📥 Importar Lançamentos** e o comparativo aparece aqui sozinho.")
    else:
        st.caption("Comparar com o mesmo mês do ano passado tira o efeito da sazonalidade "
                   "(datas comemorativas, férias, 13º salário, safra...).")
        aa = comparar_ano_anterior(r, ry)
        por_chave_aa = {c["chave"]: c for c in aa}
        cards_aa = st.columns(4)
        for col, chave, titulo in zip(cards_aa, ["receita_bruta", "custos_fixos", "ebitda", "lucro_liquido"],
                                      ["Receita Bruta", "Custos Fixos", "EBITDA", "Lucro Líquido"]):
            c = por_chave_aa[chave]
            col.metric(titulo, brl(c["atual"]),
                       f"{com_sinal(c['dif'])} ({pct_com_sinal(c['dif_pct'])}) vs {nome_mes(ano_passado)}",
                       delta_color="normal" if chave in FAVORAVEL_SE_MAIOR else "inverse")
            col.caption(f"{nome_mes(ano_passado)}: {brl(c['base'])}")

        aba_dre_aa, aba_cat_aa = st.tabs(["📋 DRE linha a linha", "🔎 Por categoria do ERP"])

        with aba_dre_aa:
            tabela_aa = pd.DataFrame({
                "Linha": [c["rotulo"] for c in aa],
                nome_mes(competencia): [brl(c["atual"]) for c in aa],
                nome_mes(ano_passado): [brl(c["base"]) for c in aa],
                "Diferença (R$)": [com_sinal(c["dif"]) for c in aa],
                "Variação (%)": [pct_com_sinal(c["dif_pct"]) for c in aa],
                "Tendência": [c["tendencia"] for c in aa],
                "Leitura": [c["leitura"] for c in aa],
            })

            def colorir_aa(linha):
                c = aa[linha.name]
                return [CORES_NIVEL[c["nivel"]] + ("; font-weight: 700" if c["total"] else "")] * len(linha)

            st.dataframe(tabela_aa.style.apply(colorir_aa, axis=1), hide_index=True, height=600)

            margens_aa = [("Margem Bruta", r["margem_bruta"], ry["margem_bruta"]),
                          ("Margem de Contribuição", r["perc_mc"], ry["perc_mc"]),
                          ("Margem EBITDA", sobre_bruta(r["ebitda"], r), sobre_bruta(ry["ebitda"], ry)),
                          ("Margem Líquida", r["margem_liquida"], ry["margem_liquida"])]
            st.dataframe(pd.DataFrame({
                "Margem (% da receita bruta)": [n for n, _, _ in margens_aa],
                nome_mes(competencia): [pct(a) for _, a, _ in margens_aa],
                nome_mes(ano_passado): [pct(b) for _, _, b in margens_aa],
                "Variação": [f"{'+' if a - b > 0.04 else ''}{(a - b):.1f} p.p.".replace(".", ",", 1)
                             for _, a, b in margens_aa],
            }), hide_index=True)
            st.caption("🟢 favorável · 🟡 desfavorável até 5% · 🔴 desfavorável acima de 5%. Em receitas e "
                       "resultados, subir é bom; em custos, despesas e saídas de caixa, subir é ruim. "
                       "p.p. = pontos percentuais.")

        with aba_cat_aa:
            cat_aa = comparar_categorias(totais_por_categoria(con, cliente_id, competencia),
                                         totais_por_categoria(con, cliente_id, ano_passado))
            filtro_aa = st.radio("Mostrar", ["Custos e despesas", "Receitas", "Todas"], horizontal=True,
                                 key=f"filtro_aa_{cliente_id}")
            if filtro_aa == "Custos e despesas":
                cat_aa = cat_aa[~cat_aa["eh_receita"]]
            elif filtro_aa == "Receitas":
                cat_aa = cat_aa[cat_aa["eh_receita"]]
            cat_aa = cat_aa.reset_index(drop=True)

            piores = cat_aa[(cat_aa["nivel"] == 2)].head(3)
            if len(piores):
                st.markdown("**Maiores pioras:** " + " · ".join(
                    f"{row.categoria} ({com_sinal(row.dif)}"
                    + (f", {pct_com_sinal(row.dif_pct)}" if row.dif_pct is not None and not pd.isna(row.dif_pct)
                       else "") + ")"
                    for row in piores.itertuples()).replace("$", "\\$"))
            if cat_aa.empty:
                st.info("Nenhuma categoria nesse filtro.")
            else:
                tabela_cat = pd.DataFrame({
                    "Categoria do ERP": cat_aa["categoria"],
                    "Linha da DRE": cat_aa["linha"].map(LINHAS).fillna("⚠️ Sem classificação"),
                    nome_mes(competencia): cat_aa["atual"].map(brl),
                    nome_mes(ano_passado): cat_aa["base"].map(brl),
                    "Diferença (R$)": cat_aa["dif"].map(com_sinal),
                    "Variação (%)": [s_ or pct_com_sinal(None if pd.isna(p_) else p_)
                                     for s_, p_ in zip(cat_aa["situacao"], cat_aa["dif_pct"])],
                    "Tendência": cat_aa["tendencia"],
                    "Leitura": cat_aa["leitura"],
                })
                niveis_cat = cat_aa["nivel"].tolist()
                st.dataframe(
                    tabela_cat.style.apply(lambda linha: [CORES_NIVEL[niveis_cat[linha.name]]] * len(linha), axis=1),
                    hide_index=True)
                st.caption("Ordenado pela maior diferença em reais. 🆕 Nova = não existia no ano anterior; "
                           "⛔ Não houve = existia no ano anterior e não teve lançamento agora.")

    st.markdown("---")
    st.subheader("🎯 Indicadores de Desempenho (KPIs para Decisão)")
    kpi1, kpi2, kpi3 = st.columns(3)

    kpi1.metric(
        "Ponto de Equilíbrio (Faturamento Bruto)", brl(r["peq"]),
        variacao("peq"), delta_color="inverse",
        help="Faturamento bruto mínimo para cobrir os custos e despesas fixas. "
             "Subir é ruim, por isso a cor da variação é invertida.",
    )
    if r["peq_situacao"] == "ok":
        folga = r["receita_bruta"] - r["peq"]
        kpi1.caption(f"Faturamento está {brl(abs(folga))} {'acima' if folga >= 0 else 'abaixo'} do equilíbrio.")
    else:
        kpi1.caption({
            "sem_faturamento": "Sem faturamento no mês: o ponto de equilíbrio não se aplica.",
            "margem_negativa": "Margem de contribuição zero ou negativa: nenhum faturamento cobre os custos.",
            "sem_custo_fixo": "Sem custos fixos no mês: qualquer venda com margem positiva já dá lucro.",
        }[r["peq_situacao"]])

    salvo = obter_funcionarios(con, cliente_id, competencia)
    funcionarios = kpi2.number_input("Funcionários no mês", min_value=0, step=1, value=salvo,
                                     key=f"func_{cliente_id}_{competencia}")
    if funcionarios != salvo:
        salvar_funcionarios(con, cliente_id, competencia, funcionarios)
        st.rerun()
    kpi2.metric("Receita por Funcionário", brl(r["receita_por_funcionario"]))
    if not funcionarios:
        kpi2.caption("Informe a quantidade de funcionários para calcular.")

    kpi3.metric("Geração de Caixa do Mês", brl(r["geracao_caixa"]), variacao("geracao_caixa"),
                help="Lucro líquido + depreciação - amortização de empréstimos - investimentos "
                     "- distribuição de lucros.")
    kpi3.caption(f"{pct(sobre_bruta(r['geracao_caixa'], r))} da receita bruta")

    with st.expander("🔎 Composição por categoria do ERP"):
        comp = totais_por_categoria(con, cliente_id, competencia)
        comp["Linha da DRE"] = comp["linha"].map(LINHAS).fillna("⚠️ Sem classificação")
        comp["Valor no arquivo"] = comp["total"].map(brl)
        comp["Valor na DRE"] = [brl(v) if pd.notna(l) and l != "ignorar" else "–"
                                for v, l in zip(comp["valor_dre"], comp["linha"])]
        st.dataframe(
            comp[["categoria", "Linha da DRE", "Valor no arquivo", "Valor na DRE", "lancamentos"]].rename(
                columns={"categoria": "Categoria", "lancamentos": "Lançamentos"}),
            hide_index=True,
        )

    st.markdown("---")
    st.subheader("✍️ Comentário do Analista")
    st.caption("Aparece no topo deste painel e no PDF do Resumo Executivo.")
    novo_comentario = st.text_area(
        "Comentário", value=comentario, height=140, label_visibility="collapsed",
        placeholder="Ex.: Aumento de 15% nos custos fixos devido à compra de novos equipamentos de informática.",
        key=f"coment_{cliente_id}_{competencia}",
    )
    if st.button("💾 Salvar comentário"):
        salvar_comentario(con, cliente_id, competencia, novo_comentario)
        st.success("Comentário salvo.")
        st.rerun()

# ---------------------------------------------------------------------------
# 📈 Histórico
# ---------------------------------------------------------------------------
elif pagina == "📈 Histórico":
    st.title(f"📈 Histórico – {nome_cliente}")
    if not meses:
        precisa_de_dados()
    aviso_sem_classificacao()

    if len(meses) > 1:
        inicio, fim = st.select_slider("Período", options=meses, value=(meses[max(0, len(meses) - 12)], meses[-1]),
                                       format_func=nome_mes)
        periodo = [m for m in meses if inicio <= m <= fim]
    else:
        periodo = meses

    resultados = {m: resultado_do_mes(m) for m in periodo}

    linhas_tabela = ESTRUTURA_DRE + [rl for rl in ESTRUTURA_CAIXA[1:]] + [
        ("Ponto de Equilíbrio (Faturamento Bruto)", "peq"),
    ]
    historico = pd.DataFrame(
        {nome_mes(m): [brl(res[k]) for _, k in linhas_tabela] for m, res in resultados.items()},
        index=[rot for rot, _ in linhas_tabela],
    )
    margens = pd.DataFrame(
        {nome_mes(m): [pct(res["margem_bruta"]), pct(res["perc_mc"]), pct(sobre_bruta(res["ebitda"], res)),
                       pct(res["margem_liquida"])]
         for m, res in resultados.items()},
        index=["Margem Bruta %", "Margem de Contribuição %", "Margem EBITDA %", "Margem Líquida %"],
    )
    st.subheader("📋 DRE mês a mês")
    st.dataframe(pd.concat([historico, margens]))

    st.subheader("📊 Evolução")
    st.caption("Área: Receita Líquida · linhas: Margem de Contribuição e Lucro Líquido. Passe o mouse sobre um mês "
               "para ver os três valores e as margens sobre a receita bruta. Ponto vermelho = prejuízo.")
    mostrar_grafico(grafico_evolucao, periodo, resultados)

    # Exportação em Excel com números (não texto) para o BPO trabalhar
    numerico = pd.DataFrame(
        {nome_mes(m): [round(res[k], 2) for _, k in linhas_tabela] for m, res in resultados.items()},
        index=[rot for rot, _ in linhas_tabela],
    )
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as escritor:
        numerico.to_excel(escritor, sheet_name="DRE mês a mês")
    st.download_button("⬇️ Baixar histórico em Excel", buffer.getvalue(),
                       file_name=f"historico_{nome_cliente}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ---------------------------------------------------------------------------
# 💸 Despesas (por natureza da despesa, mês a mês e contra o ano anterior)
# ---------------------------------------------------------------------------
elif pagina == "💸 Despesas":
    st.title(f"💸 Despesas – {nome_cliente}")
    if not meses:
        precisa_de_dados()
    preenchidos = preencher_centros_vazios(con, cliente_id)
    if preenchidos:
        st.info(f"🏷️ {preenchidos} despesa(s) fixa(s) receberam uma natureza da despesa sugerida pelo nome. "
                "Revise em **🗂️ Plano de Contas** (coluna 'Natureza da despesa').")
    aviso_sem_classificacao()

    todas = despesas_detalhadas(con, cliente_id)
    if todas.empty:
        st.info("Ainda não há despesas classificadas para este cliente. Classifique as categorias em "
                "**🗂️ Plano de Contas**.")
        st.stop()

    f1, f2 = st.columns([1, 2])
    competencia = f1.selectbox("Mês", meses[::-1], format_func=nome_mes, key=f"mes_desp_{cliente_id}")
    escopo = f2.radio("O que mostrar", ["Despesas fixas por natureza da despesa", "Todas as despesas"],
                      horizontal=True, key=f"escopo_desp_{cliente_id}")
    mes_ant = mes_seguinte(competencia, -1)
    ano_ant = mes_seguinte(competencia, -12)

    if escopo.startswith("Despesas fixas"):
        desp = todas[todas["linha"] == "custos_fixos"].copy()
        desp["grupo"] = desp["centro"]
        nome_grupo = "Natureza da despesa"
    else:
        desp = todas.copy()
        desp["grupo"] = ("Fixas · " + desp["centro"]).where(
            desp["linha"] == "custos_fixos", desp["linha"].map(GRUPOS_OUTRAS_DESPESAS))
        nome_grupo = "Grupo de despesa"
    if desp.empty:
        st.info("Nenhuma despesa fixa classificada ainda. Em **🗂️ Plano de Contas**, classifique as categorias "
                "em 'Custos e Despesas Fixas'.")
        st.stop()

    def total_mes(m, grupo=None, categoria=None):
        filtro = desp["competencia"] == m
        if grupo is not None:
            filtro &= desp["grupo"] == grupo
        if categoria is not None:
            filtro &= desp["categoria"] == categoria
        return float(desp.loc[filtro, "valor"].sum())

    tem_ant, tem_ano = mes_ant in meses, ano_ant in meses
    total = total_mes(competencia)
    r_mes = resultado_do_mes(competencia)

    # --- Cards --------------------------------------------------------------
    k1, k2, k3, k4 = st.columns(4)
    k1.metric(f"Total em {nome_mes(competencia)}", brl(total))
    k1.caption(f"{pct(sobre_bruta(total, r_mes))} da receita bruta")
    for col, titulo, base_mes, tem in [(k2, "Variação mensal", mes_ant, tem_ant),
                                       (k3, "Variação anual", ano_ant, tem_ano)]:
        if tem:
            base = total_mes(base_mes)
            col.metric(titulo, pct_com_sinal(variacao_pct(total, base)),
                       f"{com_sinal(total - base)} vs {nome_mes(base_mes)}", delta_color="inverse")
            col.caption(f"{nome_mes(base_mes)}: {brl(base)}")
        else:
            col.metric(titulo, "–")
            col.caption(f"Sem lançamentos de {nome_mes(base_mes)} importados.")
    por_grupo_mes = desp[desp["competencia"] == competencia].groupby("grupo")["valor"].sum().sort_values(ascending=False)
    if len(por_grupo_mes):
        k4.metric(f"Maior {nome_grupo.lower()}", brl(por_grupo_mes.iloc[0]))
        k4.caption(f"{por_grupo_mes.index[0]} · {pct(por_grupo_mes.iloc[0] / total * 100 if total else 0)} do total")

    # Cores fixas por grupo (a cor acompanha o grupo em todos os gráficos)
    ordem_grupos = desp.groupby("grupo")["valor"].sum().sort_values(ascending=False).index.tolist()
    cores_grupo = cores_por_grupo(ordem_grupos)

    # --- Gráficos -------------------------------------------------------------
    st.subheader(f"🅰️ Curva ABC das despesas – {nome_mes(competencia)}")
    por_categoria_mes = desp[desp["competencia"] == competencia].groupby("categoria")["valor"].sum()
    if (por_categoria_mes > 0).any():
        st.caption("Barras: cada categoria, da maior para a menor · linha ciano: % acumulado do total (eixo da "
                   "direita). Classe A (marinho) = as que somam até 80% dos gastos: é onde está o foco.")
        mostrar_grafico(grafico_pareto, por_categoria_mes)
        positivos = por_categoria_mes[por_categoria_mes > 0].sort_values(ascending=False)
        n_a = int(((positivos.cumsum() / positivos.sum() * 100) <= 80).sum()) or 1
        st.caption(f"**{n_a} de {len(positivos)} categorias** concentram ~80% das despesas do mês.")
    else:
        st.caption("Sem despesas neste mês.")
    # Janela de calendário: do mesmo mês do ano anterior até o mês escolhido (13 meses), só os que têm dados
    janela = [mes_seguinte(competencia, -i) for i in range(12, -1, -1)]
    meses_grafico = [m for m in janela if m in meses]
    st.subheader(f"🗓️ Despesas mês a mês por {nome_grupo.lower()}")
    mostrar_grafico(grafico_despesas_mes_a_mes, desp, meses_grafico, cores_grupo)

    # --- Tabela detalhada -------------------------------------------------------
    st.markdown("---")
    st.subheader(f"📋 Despesas por {nome_grupo.lower()}")
    t1, t2 = st.columns([2, 1])
    filtro_grupo = t1.selectbox(nome_grupo, ["Todos"] + ordem_grupos, key=f"filtro_grupo_{cliente_id}_{escopo}")
    detalhar = t2.toggle("Mostrar categorias", value=True, key=f"detalhar_{cliente_id}")

    col_atual, col_ant, col_ano = nome_mes(competencia), nome_mes(mes_ant), nome_mes(ano_ant)

    def celula_valor(m, tem, **kw):
        return brl(total_mes(m, **kw)) if tem else "–"

    def celula_pct(atual, m, tem, **kw):
        if not tem:
            return "–"
        base = total_mes(m, **kw)
        if not base:
            return "🆕 nova" if atual else "–"
        if not atual:  # despesa que existia e sumiu: costuma ser lançamento faltando
            return "⛔ sem lançamento"
        return pct_com_sinal(variacao_pct(atual, base))

    linhas_tab, tipos = [], []
    grupos_tab = ordem_grupos if filtro_grupo == "Todos" else [filtro_grupo]
    # Ordena pelo valor do mês escolhido (maior primeiro), como no sistema de referência
    grupos_tab = sorted(grupos_tab, key=lambda g: (-total_mes(competencia, grupo=g), -total_mes(ano_ant, grupo=g)))
    for grupo in grupos_tab:
        atual = total_mes(competencia, grupo=grupo)
        if not atual and not total_mes(mes_ant, grupo=grupo) and not total_mes(ano_ant, grupo=grupo):
            continue
        linhas_tab.append([grupo.upper(), brl(atual), celula_valor(mes_ant, tem_ant, grupo=grupo),
                           celula_pct(atual, mes_ant, tem_ant, grupo=grupo),
                           celula_valor(ano_ant, tem_ano, grupo=grupo),
                           celula_pct(atual, ano_ant, tem_ano, grupo=grupo)])
        tipos.append("grupo")
        if detalhar:
            cats = desp.loc[desp["grupo"] == grupo, "categoria"].unique()
            cats = sorted(cats, key=lambda c: -total_mes(competencia, grupo=grupo, categoria=c))
            for cat in cats:
                a = total_mes(competencia, grupo=grupo, categoria=cat)
                if not a and not total_mes(mes_ant, grupo=grupo, categoria=cat) \
                        and not total_mes(ano_ant, grupo=grupo, categoria=cat):
                    continue
                linhas_tab.append(["      " + cat, brl(a), celula_valor(mes_ant, tem_ant, grupo=grupo, categoria=cat),
                                   celula_pct(a, mes_ant, tem_ant, grupo=grupo, categoria=cat),
                                   celula_valor(ano_ant, tem_ano, grupo=grupo, categoria=cat),
                                   celula_pct(a, ano_ant, tem_ano, grupo=grupo, categoria=cat)])
                tipos.append("categoria")
    if filtro_grupo == "Todos":
        linhas_tab.append(["TOTAL", brl(total), brl(total_mes(mes_ant)) if tem_ant else "–",
                           celula_pct(total, mes_ant, tem_ant), brl(total_mes(ano_ant)) if tem_ano else "–",
                           celula_pct(total, ano_ant, tem_ano)])
        tipos.append("total")

    tabela_desp = pd.DataFrame(linhas_tab, columns=["Despesa", col_atual, col_ant, "% mensal", col_ano, "% anual"])

    def estilo_linha(linha):
        tipo = tipos[linha.name]
        base = {"grupo": f"background-color: {COR_CIANO_SUAVE}; color: {COR_MARINHO}; font-weight: 700",
                "total": f"background-color: {COR_MARINHO}; color: #ffffff; font-weight: 700",
                "categoria": ""}[tipo]
        estilos = []
        for coluna, valor in linha.items():
            estilo = base
            if coluna in ("% mensal", "% anual") and tipo == "categoria" and isinstance(valor, str):
                if valor.startswith("⛔"):
                    estilo = "color: #8a5a00; font-weight: 600"  # conferir
                elif valor.startswith("+") or valor.startswith("🆕"):
                    estilo = "color: #b42318"  # despesa subiu
                elif valor.startswith("-"):
                    estilo = "color: #0f7a3a"  # despesa caiu
            estilos.append(estilo)
        return estilos

    st.dataframe(tabela_desp.style.apply(estilo_linha, axis=1), hide_index=True,
                 height=min(38 * (len(tabela_desp) + 1), 640))
    st.caption(f"% mensal = {col_atual} contra {col_ant}. % anual = {col_atual} contra {col_ano} "
               "(mesmo mês do ano anterior, sem efeito da sazonalidade). Nas categorias, vermelho = a despesa "
               "subiu, verde = caiu. 🆕 nova = não havia essa despesa no mês de comparação. ⛔ sem lançamento = havia "
               "no mês de comparação e não há agora: confira se faltou importar.")

    # --- Mês a mês -------------------------------------------------------------
    st.subheader(f"🗓️ Mês a mês por {nome_grupo.lower()}")
    meses_tabela = meses_grafico
    pivo = (desp[desp["competencia"].isin(meses_tabela)]
            .pivot_table(index="grupo", columns="competencia", values="valor", aggfunc="sum", fill_value=0)
            .reindex(columns=meses_tabela, fill_value=0))
    pivo = pivo.loc[[g for g in ordem_grupos if g in pivo.index]]
    pivo.loc["TOTAL"] = pivo.sum()
    exibir = pivo.apply(lambda col: col.map(brl))
    exibir.columns = [nome_mes(m) for m in exibir.columns]
    exibir.index.name = nome_grupo
    st.dataframe(exibir)
    if len(meses_tabela) > 1:
        media = pivo.loc["TOTAL"].mean()
        st.caption(f"Últimos 12 meses + o mesmo mês do ano anterior (só meses importados). "
                   f"Média mensal no período: {brl(media)}. "
                   f"{nome_mes(competencia)} ficou {pct_com_sinal(variacao_pct(total, media))} em relação à média.")

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as escritor:
        numerico = pivo.copy()
        numerico.columns = [nome_mes(m) for m in numerico.columns]
        numerico.round(2).to_excel(escritor, sheet_name="Por natureza da despesa")
        detalhe = (desp[desp["competencia"].isin(meses_tabela)]
                   .pivot_table(index=["grupo", "categoria"], columns="competencia", values="valor",
                                aggfunc="sum", fill_value=0)
                   .reindex(columns=meses_tabela, fill_value=0))
        detalhe.columns = [nome_mes(m) for m in detalhe.columns]
        detalhe.index.names = [nome_grupo, "Categoria"]
        detalhe.round(2).to_excel(escritor, sheet_name="Por categoria")
    st.download_button("⬇️ Baixar despesas em Excel", buffer.getvalue(),
                       file_name=f"Despesas - {nome_cliente}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ---------------------------------------------------------------------------
# 💧 Fluxo de Caixa
# ---------------------------------------------------------------------------
elif pagina == "💧 Fluxo de Caixa":
    st.title(f"💧 Fluxo de Caixa – {nome_cliente}")
    if not meses:
        precisa_de_dados()
    mensagem_caixa = st.session_state.pop("msg_caixa", None)
    if mensagem_caixa:
        st.success(mensagem_caixa)

    # --- 1. Período e saldo inicial ------------------------------------------
    f1, f2 = st.columns([1, 2])
    competencia = f1.selectbox("Mês", meses[::-1], format_func=nome_mes, key=f"mes_caixa_{cliente_id}")
    saldo_ini, origem, mes_base = saldo_inicial_do_mes(con, cliente_id, competencia)
    with f2.expander(f"🏦 Saldo inicial em banco em 01/{competencia[5:]}/{competencia[:4]}: {brl(saldo_ini)}",
                     expanded=origem == "sem_saldo"):
        if origem == "informado":
            st.caption("Saldo informado manualmente para este mês.")
        elif origem == "calculado":
            st.caption(f"Calculado: saldo informado em {nome_mes(mes_base)} + entradas e saídas dos meses seguintes.")
        else:
            st.caption("Nenhum saldo informado ainda: o fluxo parte de R$ 0,00. Informe o saldo do extrato "
                       "bancário no 1º dia do mês (some todas as contas, aplicações de liquidez imediata inclusive).")
        novo_saldo = st.number_input("Saldo inicial (R$)", value=float(round(saldo_ini, 2)), step=100.0,
                                     format="%.2f", key=f"saldo_ini_{cliente_id}_{competencia}")
        b1, b2 = st.columns(2)
        if b1.button("💾 Salvar saldo inicial", type="primary", key=f"salvar_saldo_{cliente_id}_{competencia}"):
            salvar_saldo_inicial(con, cliente_id, competencia, novo_saldo)
            st.session_state["msg_caixa"] = (f"Saldo inicial de {nome_mes(competencia)} salvo. Os meses seguintes "
                                             "sem saldo informado passam a partir dele.")
            st.rerun()
        if origem == "informado" and b2.button("↩️ Voltar ao saldo calculado",
                                               key=f"limpar_saldo_{cliente_id}_{competencia}"):
            salvar_saldo_inicial(con, cliente_id, competencia, None)
            st.rerun()
        st.caption("Dica: informe o saldo só no primeiro mês. Os seguintes são encadeados sozinhos; informe de novo "
                   "apenas para acertar com o extrato do banco.")

    lanc_caixa = lancamentos_caixa(con, cliente_id, competencia)
    sem_sinal = lanc_caixa[lanc_caixa["caixa"].isna()] if len(lanc_caixa) else lanc_caixa
    if len(sem_sinal):
        st.warning(f"⚠️ {len(sem_sinal)} lançamento(s) de categorias sem classificação vieram de arquivo sem sinal e "
                   "ficaram fora do caixa (não dá para saber se são entrada ou saída). Classifique em "
                   "**🗂️ Plano de Contas**: " + ", ".join(sorted(sem_sinal["categoria"].unique())[:8]))
    validos = lanc_caixa.dropna(subset=["caixa"]) if len(lanc_caixa) else lanc_caixa
    entradas = float(validos["caixa"].clip(lower=0).sum()) if len(validos) else 0.0
    saidas = float(-validos["caixa"].clip(upper=0).sum()) if len(validos) else 0.0
    saldo_fim = saldo_ini + entradas - saidas

    # --- 2. Cards --------------------------------------------------------------
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Saldo inicial", brl(saldo_ini))
    k1.caption({"informado": "informado", "calculado": f"calculado desde {nome_mes(mes_base) if mes_base else ''}",
                "sem_saldo": "não informado"}[origem])
    k2.metric("Entradas do mês", brl(entradas))
    k2.caption(f"{int((validos['caixa'] > 0).sum()) if len(validos) else 0} lançamentos")
    k3.metric("Saídas do mês", brl(saidas))
    k3.caption(f"{int((validos['caixa'] < 0).sum()) if len(validos) else 0} lançamentos")
    k4.metric("Saldo final do mês", brl(saldo_fim), f"{com_sinal(saldo_fim - saldo_ini)} no mês")
    if saldo_fim < 0:
        st.error(f"🔴 O saldo final ficou negativo ({brl(saldo_fim)}): o caixa não cobriu as saídas do mês.".replace("$", "\\$"))

    # --- 3. Gráfico diário -------------------------------------------------------
    extrato = extrato_diario(lanc_caixa if len(lanc_caixa) else pd.DataFrame({"data": [], "caixa": []}),
                             competencia, saldo_ini)
    st.subheader("📈 Evolução diária do caixa")
    mostrar_grafico(grafico_caixa_diario_plotly, extrato)
    menor = extrato.loc[extrato["saldo"].idxmin()]
    st.caption(f"Menor saldo do mês: {brl(menor['saldo'])} em {menor['dia']:%d/%m}. "
               "As datas são as da coluna escolhida na importação (competência ou pagamento); para um fluxo de "
               "caixa fiel ao banco, importe usando a data de pagamento/baixa.")
    negativos = extrato[extrato["saldo"] < 0]
    if len(negativos) and saldo_fim >= 0:
        st.warning(f"⚠️ O caixa ficou negativo em {len(negativos)} dia(s) do mês (primeiro: "
                   f"{negativos.iloc[0]['dia']:%d/%m}), mesmo terminando positivo.")

    # --- 4. DFC gerencial --------------------------------------------------------
    st.markdown("---")
    st.subheader("📋 DFC Gerencial – Demonstração dos Fluxos de Caixa")
    dfc = dfc_gerencial(lanc_caixa) if len(lanc_caixa) else dfc_gerencial(pd.DataFrame({"caixa": []}))
    linhas_dfc, tipos_dfc = [["Saldo inicial", "", "", brl(saldo_ini)]], ["saldo"]
    numeros_dfc = [["Saldo inicial", None, None, saldo_ini]]  # mesma tabela em números, para o Excel
    rotulos = {"Operacional": "(+/-) Fluxo de Caixa Operacional",
               "Investimentos": "(+/-) Fluxo de Caixa de Investimentos",
               "Financiamento e Sócios": "(+/-) Fluxo de Financiamento e Sócios",
               "Não classificado": "(+/-) Não classificado (revise o Plano de Contas)"}
    totais_grupo = {}
    for grupo in ORDEM_DFC:
        parte = dfc[dfc["grupo"] == grupo]
        total = float(parte["liquido"].sum()) if len(parte) else 0.0
        totais_grupo[grupo] = total
        if grupo == "Não classificado" and parte.empty:
            continue
        linhas_dfc.append([rotulos[grupo], brl(parte["entradas"].sum()) if len(parte) else brl(0),
                           brl(parte["saidas"].sum()) if len(parte) else brl(0), com_sinal(total)])
        numeros_dfc.append([rotulos[grupo], float(parte["entradas"].sum()) if len(parte) else 0.0,
                            -float(parte["saidas"].sum()) if len(parte) else 0.0, total])
        tipos_dfc.append("grupo")
        for item in parte.itertuples():
            linhas_dfc.append(["      " + item.linha_nome, brl(item.entradas) if item.entradas else "–",
                               brl(item.saidas) if item.saidas else "–", com_sinal(item.liquido)])
            numeros_dfc.append(["      " + item.linha_nome, float(item.entradas), -float(item.saidas),
                                float(item.liquido)])
            tipos_dfc.append("linha")
    variacao_caixa = sum(totais_grupo.values())
    linhas_dfc.append(["(=) Variação do caixa no mês", brl(entradas), brl(saidas), com_sinal(variacao_caixa)])
    tipos_dfc.append("subtotal")
    linhas_dfc.append(["(=) Saldo final", "", "", brl(saldo_ini + variacao_caixa)])
    numeros_dfc += [["(=) Variação do caixa no mês", entradas, -saidas, variacao_caixa],
                    ["(=) Saldo final", None, None, saldo_ini + variacao_caixa]]
    tipos_dfc.append("saldo")
    tabela_dfc = pd.DataFrame(linhas_dfc, columns=["Fluxo", "Entradas", "Saídas", "Líquido"])

    def estilo_dfc(linha):
        tipo = tipos_dfc[linha.name]
        return [{"grupo": f"background-color: {COR_CIANO_SUAVE}; color: {COR_MARINHO}; font-weight: 700",
                 "subtotal": f"background-color: #EAF0F7; color: {COR_MARINHO}; font-weight: 700",
                 "saldo": f"background-color: {COR_MARINHO}; color: #ffffff; font-weight: 700",
                 "linha": ""}[tipo]] * len(linha)

    st.dataframe(tabela_dfc.style.apply(estilo_dfc, axis=1), hide_index=True,
                 height=min(36 * (len(tabela_dfc) + 1), 720))
    diferenca = abs((saldo_ini + variacao_caixa) - saldo_fim)
    if diferenca > 0.01:  # trava de segurança: nunca deve acontecer
        st.error(f"❌ Reconciliação não fechou (diferença {brl(diferenca)}). Envie esta tela ao suporte.")
    else:
        st.caption("✅ Reconciliação: saldo inicial + operacional + investimentos + financiamento/sócios = saldo final. "
                   "Operacional inclui receitas, impostos, custos, despesas e juros; Investimentos, compra de ativos; "
                   "Financiamento e Sócios, amortização de empréstimos e distribuição de lucros/retiradas. "
                   "Transferências entre contas e depreciação não entram no caixa.")
    r_mes = resultado_do_mes(competencia)
    if totais_grupo.get("Não classificado", 0) == 0 and abs(r_mes["geracao_caixa"] - variacao_caixa) <= 0.01:
        st.caption(f"✅ Bate com a DRE: Geração de Caixa do Mês = {brl(r_mes['geracao_caixa'])}.")
    else:
        st.caption(f"ℹ️ Geração de Caixa na DRE: {brl(r_mes['geracao_caixa'])}. A diferença vem de lançamentos "
                   "sem classificação no Plano de Contas.")

    # --- 5. Extrato e exportação --------------------------------------------------
    with st.expander("🧾 Extrato de lançamentos do mês"):
        if len(validos):
            extrato_lanc = validos.copy()
            extrato_lanc["saldo"] = saldo_ini + extrato_lanc["caixa"].cumsum()
            st.dataframe(pd.DataFrame({
                "Data": pd.to_datetime(extrato_lanc["data"]).dt.strftime("%d/%m/%Y"),
                "Categoria": extrato_lanc["categoria"], "Descrição": extrato_lanc["descricao"],
                "Fluxo": extrato_lanc["grupo"], "Entrada": extrato_lanc["caixa"].clip(lower=0).map(
                    lambda v: brl(v) if v else ""),
                "Saída": (-extrato_lanc["caixa"].clip(upper=0)).map(lambda v: brl(v) if v else ""),
                "Saldo": extrato_lanc["saldo"].map(brl),
            }), hide_index=True)
        else:
            st.caption("Sem lançamentos de caixa neste mês.")

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as escritor:
        resumo = pd.DataFrame({
            "Item": ["Saldo inicial", "Entradas", "Saídas", "Saldo final"] + [f"Fluxo {g}" for g in ORDEM_DFC],
            "Valor (R$)": [round(v, 2) for v in [saldo_ini, entradas, -saidas, saldo_fim]
                           + [totais_grupo.get(g, 0.0) for g in ORDEM_DFC]],
        })
        resumo.to_excel(escritor, sheet_name="Resumo", index=False)
        pd.DataFrame(numeros_dfc, columns=["Fluxo", "Entradas", "Saídas", "Líquido"]).round(2).to_excel(
            escritor, sheet_name="DFC", index=False)
        pd.DataFrame({"Dia": extrato["dia"].dt.date, "Entradas": extrato["entradas"].round(2),
                      "Saídas": (-extrato["saidas"]).round(2), "Saldo ao fim do dia": extrato["saldo"].round(2)}
                     ).to_excel(escritor, sheet_name="Saldo diário", index=False)
        if len(validos):
            pd.DataFrame({"Data": pd.to_datetime(validos["data"]).dt.date, "Categoria": validos["categoria"],
                          "Descrição": validos["descricao"], "Linha da DRE": validos["linha_nome"],
                          "Fluxo": validos["grupo"], "Valor no caixa": validos["caixa"].round(2),
                          "Saldo": (saldo_ini + validos["caixa"].cumsum()).round(2)}
                         ).to_excel(escritor, sheet_name="Extrato", index=False)
        for aba in escritor.book.worksheets:
            for coluna in aba.columns:
                aba.column_dimensions[coluna[0].column_letter].width = max(
                    12, min(48, max(len(str(c.value or "")) for c in coluna) + 2))
                for celula in coluna[1:]:
                    if isinstance(celula.value, float):
                        celula.number_format = '#,##0.00;-#,##0.00'
    st.download_button("⬇️ Exportar fluxo de caixa em Excel", buffer.getvalue(),
                       file_name=f"Fluxo de Caixa - {nome_cliente} - {nome_mes(competencia).replace('/', '-')}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ---------------------------------------------------------------------------
# 🎯 Orçado x Realizado
# ---------------------------------------------------------------------------
elif pagina == "🎯 Orçado x Realizado":
    st.title(f"🎯 Orçado x Realizado – {nome_cliente}")
    mensagem = st.session_state.pop("msg_orc", None)
    if mensagem:
        st.success(mensagem)

    meses_orc = meses_com_orcamento(con, cliente_id)
    hoje = datetime.now().strftime("%Y-%m")
    opcoes = sorted(set(meses) | set(meses_orc) | {mes_seguinte(hoje, i) for i in range(-1, 12)})[::-1]
    padrao = meses[-1] if meses else hoje
    competencia = st.selectbox("Mês", opcoes, index=opcoes.index(padrao), format_func=nome_mes)
    orcado_mes = carregar_orcamento(con, cliente_id, competencia)

    aba_comp, aba_cad, aba_imp = st.tabs(["📊 Comparativo", "✏️ Cadastrar orçamento", "📥 Importar planilha"])

    with aba_comp:
        if not orcado_mes:
            st.info("Ainda não há orçamento para este mês. Cadastre na aba **✏️ Cadastrar orçamento** "
                    "ou envie a planilha na aba **📥 Importar planilha**.")
        else:
            if competencia not in meses:
                st.warning("Ainda não há lançamentos importados para este mês: o realizado aparece zerado.")
            aviso_sem_classificacao(competencia)
            comp = comparar_orcamento(calcular(orcado_mes), resultado_do_mes(competencia))
            por_chave = {c["chave"]: c for c in comp}

            def sinal(valor):
                return ("+" if valor > 0 else "") + brl(valor)

            def desvio(c):
                return f"{sinal(c['dif'])} ({'+' if c['dif_pct'] and c['dif_pct'] > 0 else ''}" \
                       f"{pct(c['dif_pct'])})" if c["dif_pct"] is not None else sinal(c["dif"])

            cards = st.columns(4)
            for col, chave, titulo in zip(cards, ["receita_bruta", "custos_fixos", "ebitda", "lucro_liquido"],
                                          ["Receita Bruta", "Custos Fixos", "EBITDA", "Lucro Líquido"]):
                c = por_chave[chave]
                col.metric(titulo, brl(c["realizado"]), f"{desvio(c)} vs orçado",
                           delta_color="normal" if chave in FAVORAVEL_SE_MAIOR else "inverse",
                           help=f"Orçado: {brl(c['orcado'])}")
                col.caption(f"Orçado: {brl(c['orcado'])}")

            estouros = [c for c in comp if c["nivel"] == 2 and not c["total"] and c["chave"] not in FAVORAVEL_SE_MAIOR]
            abaixo = [c for c in comp if c["nivel"] == 2 and c["chave"] in FAVORAVEL_SE_MAIOR]
            if estouros:
                st.error("🔴 **Estouro de despesas:** " + "; ".join(
                    f"{c['rotulo'].split(') ')[-1]} ficou {desvio(c)} acima do orçado" for c in estouros))
            if abaixo:
                st.warning("🔻 **Abaixo do orçado:** " + "; ".join(
                    f"{c['rotulo'].split(') ')[-1]} {desvio(c)}" for c in abaixo))
            if not estouros and not abaixo:
                st.success(f"🟢 Nenhuma linha com desvio desfavorável acima de {pct(TOLERANCIA_PCT)}.")

            tabela = pd.DataFrame({
                "Linha": [c["rotulo"] for c in comp],
                "Orçado": [brl(c["orcado"]) for c in comp],
                "Realizado": [brl(c["realizado"]) for c in comp],
                "Diferença (R$)": [sinal(c["dif"]) for c in comp],
                "Diferença (%)": [("+" if c["dif_pct"] > 0 else "") + pct(c["dif_pct"]) if c["dif_pct"] is not None
                                  else "–" for c in comp],
                "Situação": [c["status"] for c in comp],
            })
            cores = {2: "background-color: #fde2e1; color: #7a1414",
                     1: "background-color: #fff4d6; color: #5c3d00", 0: ""}

            def colorir(linha):
                c = comp[linha.name]
                estilo = cores[c["nivel"]] + ("; font-weight: 700" if c["total"] else "")
                return [estilo] * len(linha)

            st.markdown("**Orçado x Realizado por linha da DRE**")
            st.caption("Ciano = orçado · marinho = realizado. O rótulo mostra o desvio: ▲ verde favorável, "
                       "▼ vermelho desfavorável. Passe o mouse para ver o desvio exato em R$.")
            mostrar_grafico(grafico_orcado_realizado, comp)
            st.dataframe(tabela.style.apply(colorir, axis=1), hide_index=True, height=600)
            st.caption(
                f"🟢 dentro do orçado ou melhor · 🟡 desvio desfavorável de até {pct(TOLERANCIA_PCT)} · "
                f"🔴 desvio desfavorável acima de {pct(TOLERANCIA_PCT)}. Em receitas e resultados, ficar abaixo "
                "é desfavorável; em custos, despesas e saídas de caixa, ficar acima é estouro."
            )

    with aba_cad:
        st.caption("Digite os valores previstos, todos positivos. Receita líquida, margem, EBITDA, lucro e "
                   "geração de caixa orçados são calculados sozinhos.")
        editor = pd.DataFrame({
            "Linha da DRE": [LINHAS[k] for k in LINHAS_VALOR],
            "Orçado (R$)": [float(orcado_mes.get(k, 0.0)) for k in LINHAS_VALOR],
        })
        editado = st.data_editor(
            editor, hide_index=True, disabled=["Linha da DRE"], height=460,
            column_config={"Orçado (R$)": st.column_config.NumberColumn(min_value=0.0, step=100.0, format="R$ %.2f")},
            key=f"orc_{cliente_id}_{competencia}",
        )
        b1, b2 = st.columns(2)
        if b1.button("💾 Salvar orçamento", type="primary"):
            salvar_orcamento(con, cliente_id, competencia, {
                LINHAS_POR_ROTULO[linha]: float(v or 0) for linha, v in zip(editado["Linha da DRE"], editado["Orçado (R$)"])
            })
            st.session_state["msg_orc"] = f"Orçamento de {nome_mes(competencia)} salvo."
            st.rerun()
        anteriores = [m for m in meses_orc if m < competencia]
        if anteriores and b2.button(f"📋 Copiar o orçamento de {nome_mes(anteriores[-1])}"):
            salvar_orcamento(con, cliente_id, competencia, carregar_orcamento(con, cliente_id, anteriores[-1]))
            st.session_state.pop(f"orc_{cliente_id}_{competencia}", None)
            st.session_state["msg_orc"] = f"Orçamento de {nome_mes(anteriores[-1])} copiado para {nome_mes(competencia)}."
            st.rerun()

    with aba_imp:
        st.markdown(
            "1. Baixe o modelo (12 meses a partir do mês escolhido, já com o que estiver cadastrado).\n"
            "2. Preencha no Excel os valores previstos, todos positivos.\n"
            "3. Envie a planilha abaixo e confira antes de importar."
        )
        meses_modelo = [mes_seguinte(competencia, i) for i in range(12)]
        preenchido = {m: carregar_orcamento(con, cliente_id, m) for m in meses_modelo}
        st.download_button(
            "⬇️ Baixar modelo de orçamento", modelo_orcamento(meses_modelo, preenchido),
            file_name=f"Orcamento - {nome_cliente}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        arquivo_orc = st.file_uploader("Planilha de orçamento preenchida (.xlsx)", type=["xlsx"], key="upload_orcamento")
        if arquivo_orc is not None:
            try:
                lido, avisos = ler_orcamento(arquivo_orc.getvalue())
            except Exception as erro:  # noqa: BLE001
                lido, avisos = {}, [f"Não consegui ler a planilha: {erro}"]
            for aviso in avisos:
                st.warning(aviso)
            if lido:
                previa = pd.DataFrame(
                    {nome_mes(m): {LINHAS[k]: brl(v) for k, v in valores.items()} for m, valores in sorted(lido.items())}
                ).reindex([LINHAS[k] for k in LINHAS_VALOR]).fillna("")
                st.dataframe(previa)
                ja_tem = [m for m in lido if m in meses_orc]
                if ja_tem:
                    st.caption(f"Vai substituir o orçamento já cadastrado de: {', '.join(nome_mes(m) for m in sorted(ja_tem))}.")
                if st.button(f"✅ Importar orçamento de {len(lido)} mês(es)", type="primary"):
                    for m, valores in lido.items():
                        salvar_orcamento(con, cliente_id, m, valores)
                        st.session_state.pop(f"orc_{cliente_id}_{m}", None)
                    st.session_state["msg_orc"] = (
                        f"Orçamento importado: {', '.join(nome_mes(m) for m in sorted(lido))}.")
                    st.rerun()
            elif not avisos:
                st.warning("Não encontrei nenhum valor preenchido na planilha.")

# ---------------------------------------------------------------------------
# 📥 Importar Lançamentos
# ---------------------------------------------------------------------------
elif pagina == "📥 Importar Lançamentos":

    def pagina_importar():
        st.title(f"📥 Importar Lançamentos – {nome_cliente}")
        st.markdown(
            "Exporte do ERP (Conta Azul, Omie, Nibo, Bling, Sankhya, TOTVS…) o relatório de **lançamentos / "
            "movimentações** do período, em **.xlsx**, **.xls** ou **.csv**. O painel acha a tabela sozinho, "
            "mesmo com logotipo, título e filtros em cima."
        )
        arquivo = st.file_uploader("Arquivo do ERP", type=["xlsx", "xlsm", "xls", "csv", "txt"])
        if arquivo is None:
            return
        conteudo = arquivo.getvalue()
        chave = f"{arquivo.name}_{len(conteudo)}"

        # --- 1. Leitura: acha aba e cabeçalho -------------------------------------
        try:
            tabela, info = ler_arquivo(arquivo.name, conteudo)
            with st.expander("⚙️ Ajustar leitura (aba e linha do cabeçalho)", expanded=not info["confiavel"]):
                aba = info["aba"]
                if len(info["abas"]) > 1:
                    aba = st.selectbox("Aba da planilha", info["abas"], index=info["abas"].index(info["aba"]),
                                       key=f"aba_{chave}")
                    if aba != info["aba"]:
                        tabela, info = ler_arquivo(arquivo.name, conteudo, aba=aba)
                linha = st.number_input(
                    "Linha do cabeçalho (o número da linha no Excel onde estão os nomes das colunas)",
                    min_value=1, step=1, value=info["linha_cabecalho"], key=f"linha_{chave}_{aba}")
                if linha != info["linha_cabecalho"]:
                    tabela, info = ler_arquivo(arquivo.name, conteudo, aba=aba, linha_cabecalho=int(linha))
        except ErroImportacao as erro:
            st.error(f"❌ {erro}")
            return

        if info["confiavel"]:
            st.success(f"✅ Tabela encontrada: aba '{info['aba']}', cabeçalho na linha {info['linha_cabecalho']}, "
                       f"{len(tabela)} linhas de dados.")
        else:
            st.warning(
                f"⚠️ Não tenho certeza de onde começa a tabela (usei a linha {info['linha_cabecalho']}). "
                "Confira a prévia: se os nomes das colunas estiverem errados, informe a linha certa em "
                "**⚙️ Ajustar leitura**.")
        st.dataframe(tabela.head(10), hide_index=True)

        # --- 2. Colunas: lembra o layout do cliente, senão adivinha por apelidos ---
        colunas = list(tabela.columns)
        layout = carregar_layout(con, cliente_id)
        lembrado = bool(layout) and all(layout.get(p) in colunas for p in ("data", "categoria"))

        def parece_numerica(coluna):
            if not coluna:
                return False
            texto = tabela[coluna].notna() & (tabela[coluna].astype(str).str.strip() != "")
            return bool(texto.sum()) and para_numero(tabela[coluna])[texto].notna().mean() > 0.5

        def sugestao(papel, usadas=()):
            if lembrado and layout.get(papel) in colunas:
                return layout[papel]
            return adivinhar_coluna(colunas, papel, usadas)

        s_data = sugestao("data")
        s_cat = sugestao("categoria", [s_data])
        s_valor = sugestao("valor", [s_data, s_cat])
        s_desc = sugestao("descricao", [s_data, s_cat, s_valor])
        s_tipo = sugestao("tipo", [s_data, s_cat, s_valor, s_desc])
        s_ent = sugestao("entrada", [s_data, s_cat, s_desc, s_tipo])
        s_sai = sugestao("saida", [s_data, s_cat, s_desc, s_tipo, s_ent])
        if not parece_numerica(s_ent) or not parece_numerica(s_sai):
            s_ent = s_sai = None
        if lembrado and layout.get("modo") in MODOS:
            s_modo = layout["modo"]
        else:
            s_modo = sugerir_modo(tabela, s_valor, s_tipo, s_ent, s_sai)

        escolha_vazia, nenhuma = "— escolha a coluna —", "— nenhuma —"

        def seletor(onde, rotulo, sugerida, papel, opcional=False, ajuda=None):
            opcoes = [nenhuma if opcional else escolha_vazia] + colunas
            escolhida = onde.selectbox(rotulo, opcoes, index=opcoes.index(sugerida) if sugerida in opcoes else 0,
                                       key=f"{papel}_{chave}", help=ajuda)
            return None if escolhida in (escolha_vazia, nenhuma) else escolhida

        st.subheader("1. Quais colunas usar?")
        if lembrado:
            st.info("🧠 Usei as mesmas colunas da última importação deste cliente. Se o layout mudou, é só trocar.")
        else:
            st.caption("O painel sugeriu as colunas pelos nomes. Se o ERP usa outro nome (ex.: 'Plano de Contas' em vez "
                       "de 'Categoria', 'Histórico' em vez de 'Descrição'), troque na lista.")
        c1, c2 = st.columns(2)
        col_data = seletor(c1, "Data que define o mês", s_data, "data",
                           ajuda="Data de competência → DRE por competência (recomendado). "
                                 "Data de pagamento/baixa → visão de caixa.")
        col_categoria = seletor(c2, "Categoria / Plano de Contas", s_cat, "categoria",
                                ajuda="A classificação do lançamento no ERP. Se o ERP não tiver, use a descrição.")
        col_descricao = seletor(c1, "Descrição / Histórico (opcional)", s_desc, "descricao", opcional=True)

        st.subheader("2. Como o arquivo mostra entradas e saídas?")
        modo = st.radio("Formato dos valores", MODOS, index=MODOS.index(s_modo), key=f"modo_{chave}",
                        label_visibility="collapsed")
        c3, c4 = st.columns(2)
        col_valor = col_tipo = col_entrada = col_saida = None
        if modo == MODO_DUAS:
            col_entrada = seletor(c3, "Coluna de Entradas / Crédito", s_ent, "entrada")
            col_saida = seletor(c4, "Coluna de Saídas / Débito", s_sai, "saida")
        else:
            col_valor = seletor(c3, "Valor", s_valor, "valor")
            if modo == MODO_TIPO:
                col_tipo = seletor(c4, "Coluna de tipo", s_tipo, "tipo",
                                   ajuda="Coluna que diz se é entrada ou saída: Receita/Despesa, C/D, E/S, "
                                         "Recebimento/Pagamento, A Receber/A Pagar...")
                if col_tipo:
                    valores_tipo = tabela[col_tipo].dropna().astype(str).str.strip()
                    resumo_tipo = valores_tipo.value_counts().head(15)
                    leitura = {1: "➕ Entrada", -1: "➖ Saída", None: "❓ Não reconhecido (fica de fora)"}
                    c4.dataframe(pd.DataFrame({
                        "Valor no arquivo": resumo_tipo.index,
                        "Linhas": resumo_tipo.values,
                        "Entendido como": [leitura[sentido_do_tipo(v)] for v in resumo_tipo.index],
                    }), hide_index=True)

        # --- 3. Conversão e conferência ------------------------------------------
        try:
            lanc, descartadas, avisos = preparar(
                tabela, col_data, col_categoria, col_valor, col_descricao, modo, col_tipo, col_entrada, col_saida)
        except ErroImportacao as erro:
            st.error(f"❌ {erro}")
            return
        for aviso in avisos:
            st.warning(f"⚠️ {aviso}")

        meses_arquivo = sorted(lanc["competencia"].unique())
        st.subheader("3. Conferência")
        m1, m2, m3 = st.columns(3)
        m1.metric("Lançamentos válidos", len(lanc))
        m2.metric("Meses", ", ".join(nome_mes(m) for m in meses_arquivo))
        m3.metric("Categorias", lanc["categoria"].nunique())
        m4, m5, m6 = st.columns(3)
        m4.metric("Total de entradas", brl(lanc.loc[lanc["valor"] > 0, "valor"].sum()))
        m5.metric("Total de saídas", brl(-lanc.loc[lanc["valor"] < 0, "valor"].sum()))
        m6.metric("Saldo do arquivo", brl(lanc["valor"].sum()))
        st.caption("Compare esses totais com o relatório do ERP antes de importar.")
        if len(descartadas):
            resumo = descartadas["Motivo"].value_counts()
            with st.expander(f"⚠️ {len(descartadas)} linha(s) ignoradas: "
                             + " · ".join(f"{m}: {n}" for m, n in resumo.items())):
                st.caption("Normalmente são linhas de total, subtotal ou rodapé do relatório.")
                st.dataframe(descartadas.head(200), hide_index=True)

        sinal_confiavel = modo != MODO_POSITIVO
        pode_importar = True
        if modo == MODO_SINAL and not (lanc["valor"] < 0).any():
            pode_importar = st.checkbox(
                "Confirmo que este arquivo realmente não tem saídas (só entradas). Sem essa confirmação, "
                "a importação fica bloqueada, porque no modo 'com sinal' despesas positivas viram receita.",
                value=False, key=f"confirma_sem_saidas_{chave}")

        ja_existem = [m for m in meses_arquivo if m in meses]
        substituir = True
        if ja_existem:
            substituir = st.checkbox(
                f"Substituir os lançamentos já importados de {', '.join(nome_mes(m) for m in ja_existem)}",
                value=True, help="Desmarque só se este arquivo complementa o anterior (ex.: outra conta bancária).",
            )

        if st.button("✅ Importar lançamentos", type="primary", disabled=not pode_importar):
            quantidade, meses_gravados = salvar_lancamentos(con, cliente_id, lanc, arquivo.name, substituir,
                                                            sinal_confiavel=sinal_confiavel)
            mapeamento = carregar_mapeamento(con, cliente_id)
            somas = lanc.groupby("categoria")["valor"].sum()
            novas = {cat: sugerir_linha(cat, soma, sinal_confiavel)
                     for cat, soma in somas.items() if cat not in mapeamento}
            salvar_mapeamento(con, cliente_id, novas)
            preencher_centros_vazios(con, cliente_id)
            salvar_layout(con, cliente_id, {
                "data": col_data, "categoria": col_categoria, "valor": col_valor, "descricao": col_descricao,
                "tipo": col_tipo, "entrada": col_entrada, "saida": col_saida, "modo": modo,
            })
            st.success(f"{quantidade} lançamentos importados em {', '.join(nome_mes(m) for m in meses_gravados)}.")
            if novas:
                st.info(
                    f"{len(novas)} categoria(s) nova(s) receberam uma classificação sugerida. "
                    "**Revise em 🗂️ Plano de Contas** antes de apresentar ao cliente."
                )

    try:
        pagina_importar()
    except BaseException as erro:  # noqa: BLE001
        if _controle_do_streamlit(erro) or isinstance(erro, (KeyboardInterrupt, SystemExit)):
            raise
        st.error("❌ Não consegui concluir a leitura deste arquivo. Confira se é o relatório de lançamentos "
                 "exportado do ERP. Se o problema continuar, envie o arquivo e os detalhes abaixo para o suporte.")
        with st.expander("Detalhes técnicos"):
            st.code("".join(traceback.format_exception(erro)))

# ---------------------------------------------------------------------------
# 🗂️ Plano de Contas
# ---------------------------------------------------------------------------
elif pagina == "🗂️ Plano de Contas":
    st.title(f"🗂️ Plano de Contas – {nome_cliente}")
    st.markdown(
        "Diga em qual linha da DRE cada categoria do ERP entra. "
        "A classificação vale para todos os meses deste cliente."
    )
    totais = totais_por_categoria(con, cliente_id)
    if totais.empty:
        precisa_de_dados()

    aviso_nova_estrutura()
    preenchidos = preencher_centros_vazios(con, cliente_id)
    if preenchidos:
        st.info(f"🏷️ {preenchidos} despesa(s) fixa(s) receberam uma natureza da despesa sugerida pelo nome. "
                "Confira na coluna 'Natureza da despesa' e troque se precisar.")
    centros = listar_centros(con, cliente_id)
    centro_da_categoria = carregar_centros_categorias(con, cliente_id)

    por_categoria = totais.groupby("categoria", as_index=False).agg(
        total=("total", "sum"), lancamentos=("lancamentos", "sum"), linha=("linha", "first"))
    editor = pd.DataFrame({
        "Categoria do ERP": por_categoria["categoria"],
        "Linha da DRE": por_categoria["linha"].map(LINHAS),
        "Natureza da despesa": [centro_da_categoria.get(c) if l == "custos_fixos" else None
                            for c, l in zip(por_categoria["categoria"], por_categoria["linha"])],
        "Total importado": por_categoria["total"].map(brl),
        "Lançamentos": por_categoria["lancamentos"],
    }).sort_values(["Linha da DRE", "Natureza da despesa", "Categoria do ERP"], na_position="first")

    pendentes = editor["Linha da DRE"].isna().sum()
    if pendentes:
        st.warning(f"{pendentes} categoria(s) sem classificação (no topo da tabela).")

    mostrar = st.radio("Mostrar", ["Todas as categorias", "Só despesas fixas (natureza da despesa)"], horizontal=True,
                       key=f"mostrar_plano_{cliente_id}")
    if mostrar.startswith("Só"):
        editor = editor[editor["Linha da DRE"] == LINHAS["custos_fixos"]]
    st.caption("A **natureza da despesa** é usada só nas categorias em 'Custos e Despesas Fixas' "
               "(ex.: Salários → Pessoal; Manutenção freezer → Manutenção). Ela organiza a aba **💸 Despesas**.")

    editado = st.data_editor(
        editor, hide_index=True,
        disabled=["Categoria do ERP", "Total importado", "Lançamentos"],
        column_config={
            "Linha da DRE": st.column_config.SelectboxColumn(options=list(LINHAS.values()), required=True),
            "Natureza da despesa": st.column_config.SelectboxColumn(
                options=centros, required=False, help="Só para categorias em 'Custos e Despesas Fixas'."),
        },
        key=f"plano_{cliente_id}_{mostrar}",
    )
    if st.button("💾 Salvar classificação", type="primary"):
        novo, centros_novos = {}, {}
        for _, linha in editado.iterrows():
            if pd.isna(linha["Linha da DRE"]):
                continue
            chave_linha = LINHAS_POR_ROTULO[linha["Linha da DRE"]]
            novo[linha["Categoria do ERP"]] = chave_linha
            if chave_linha == "custos_fixos":
                centro = linha["Natureza da despesa"]
                centros_novos[linha["Categoria do ERP"]] = centro if centro in centros else \
                    sugerir_centro(linha["Categoria do ERP"], centros)
            else:
                centros_novos[linha["Categoria do ERP"]] = None
        salvar_mapeamento(con, cliente_id, novo)
        salvar_centros_categorias(con, cliente_id, centros_novos)
        st.success("Classificação salva.")
        st.rerun()

    with st.expander("🏷️ Naturezas da despesa (subgrupos das despesas fixas)"):
        uso = pd.Series(carregar_centros_categorias(con, cliente_id)).value_counts()
        st.dataframe(pd.DataFrame({"Natureza da despesa": centros,
                                   "Categorias": [int(uso.get(c, 0)) for c in centros]}), hide_index=True)
        n1, n2 = st.columns([3, 1])
        novo_centro = n1.text_input("Nova natureza da despesa", placeholder="Ex.: Frota, Cozinha, Recepção...",
                                    key=f"novo_centro_{cliente_id}")
        n2.write("")
        if n2.button("➕ Criar"):
            if not novo_centro.strip():
                st.error("Digite o nome da natureza da despesa.")
            elif novo_centro.strip().lower() in [c.lower() for c in centros]:
                st.error("Essa natureza da despesa já existe.")
            else:
                criar_centro(con, cliente_id, novo_centro)
                st.rerun()
        r1, r2 = st.columns([3, 1])
        excluir = r1.selectbox("Remover natureza da despesa", centros, key=f"excluir_centro_{cliente_id}")
        r2.write("")
        if r2.button("🗑️ Remover", disabled=excluir == CENTRO_RESERVA):
            movidas = remover_centro(con, cliente_id, excluir)
            if movidas:
                st.session_state["msg_plano"] = f"{movidas} categoria(s) foram para '{CENTRO_RESERVA}'."
            st.rerun()
        st.caption(f"Ao remover uma natureza, as categorias dela vão para '{CENTRO_RESERVA}'.")
    mensagem_plano = st.session_state.pop("msg_plano", None)
    if mensagem_plano:
        st.info(mensagem_plano)

    st.markdown("---")
    st.subheader("📦 Dados importados")
    resumo = resumo_importacoes(con, cliente_id)
    limite = PACOTES.get(cliente["pacote"])
    resumo.insert(0, "Mês", resumo["competencia"].map(nome_mes))
    resumo["ultima_importacao"] = pd.to_datetime(resumo["ultima_importacao"]).dt.strftime("%d/%m/%Y %H:%M")
    if limite:
        resumo["Uso do pacote"] = resumo["lancamentos"].map(
            lambda n: f"{n}/{limite}" + (" ⚠️ acima do limite" if n > limite else ""))
    st.dataframe(
        resumo.drop(columns=["competencia"]).rename(columns={
            "lancamentos": "Lançamentos", "ultima_importacao": "Última importação", "arquivos": "Arquivos"}),
        hide_index=True,
    )
    if limite:
        st.caption("Uso do pacote conta todas as linhas importadas no mês, inclusive transferências entre contas.")

    with st.expander("🗑️ Excluir um mês importado"):
        mes_excluir = st.selectbox("Mês", meses[::-1], format_func=nome_mes, key="mes_excluir")
        confirma = st.checkbox(f"Confirmo a exclusão de todos os lançamentos de {nome_mes(mes_excluir)}")
        if st.button("Excluir", disabled=not confirma):
            excluir_mes(con, cliente_id, mes_excluir)
            st.rerun()
