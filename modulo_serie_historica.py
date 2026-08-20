"""
Modulo: Serie Historica de Investimentos -- Fundos Exclusivos
================================================================
Le todos os CSVs da pasta "Fundos Exclusivos" (um por competencia,
nome no padrao YYYYMM ou YYYY-MM), empilha em serie temporal e
gera os datasets necessarios para o dashboard.

Layout de cada CSV (sem cabecalho):
  A = DATA (YYYYMM)
  B = COD_EFPC (NU_MATRICULA_EFPC)
  C = CNPB (codigo do plano)
  D = TIPO_ATIVO
  E = VALOR
  F = CNPJ (preenchido so se for fundo exclusivo)
"""

import glob
import os
import re

import numpy as np
import pandas as pd

# Caminho padrão usado só se a função for chamada sem argumento — na prática
# o gerador principal (efpc_dashboard_gerador_v4.py) sempre passa o caminho
# explicitamente (relativo a BASE), então este valor nunca entra em uso real.
PASTA_FUNDOS_EXCLUSIVOS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Fundos Exclusivos")

CLASSE_MAP = {
    "TIT_PUB":  "Títulos Públicos",
    "TIT_PRIV": "Crédito Privado",
    "ACOES":    "Renda Variável",
    "COTAS":    "Fundos",
    "IMOVEIS":  "Imóveis",
    "ALUGUEL":  "Imóveis",
    "OPERACAO_PARTICIPANTE": "Op. Participantes",
    "CAIXA":    "Caixa/Disponível",
    "OP_COMP":  "Op. Compromissadas",
    "SWAP":     "Derivativos",
    "TERMO":    "Derivativos",
    "OPCOES":   "Derivativos",
}
TIPOS_FLUXO = {"VL_PG_RC"}


def _extrai_competencia(nome_arquivo):
    """Extrai YYYYMM do nome do arquivo, aceitando YYYYMM ou YYYY-MM / YYYY_MM."""
    base = os.path.splitext(os.path.basename(nome_arquivo))[0]
    m = re.search(r"(20\d{2})[-_]?(\d{2})", base)
    if not m:
        return None
    ano, mes = m.groups()
    if not (1 <= int(mes) <= 12):
        return None
    return ano + mes


def _tem_pontos_multiplos(serie_valor_str, limiar=0.05):
    """
    Detecta se uma coluna de valores tem uma fracao relevante de linhas com
    MULTIPLOS pontos (ex: '3.016.450.997'). Isso indica que o arquivo passou
    pelo Excel, que reformatou numeros decimais como se fossem inteiros com
    separador de milhar, descartando a posicao decimal original.
    """
    amostra = serie_valor_str.dropna().astype(str)
    if len(amostra) == 0:
        return False
    n_multi = (amostra.str.count(r"\.") >= 2).sum()
    return (n_multi / len(amostra)) > limiar


def _corrige_pontos_multiplos(valor_str):
    """
    Corrige um valor com multiplos pontos: o ULTIMO ponto e tratado como o
    separador decimal real; todos os pontos anteriores sao removidos (eram
    separador de milhar espurio introduzido pelo Excel). Valores com 0 ou 1
    ponto passam direto, sem alteracao.

    Valores em notacao cientifica sao um caso especial: quando o Excel
    reformata um numero longo (ex: um CNPJ ou um valor decimal mal lido)
    como notacao cientifica, a informacao original e perdida de forma
    irrecuperavel. Nesse caso, em vez de propagar um numero absurdo
    (ex: 1e16), o valor e descartado (tratado como 0) e reportado.
    """
    if pd.isna(valor_str):
        return np.nan
    s = str(valor_str).strip()
    if s == "":
        return np.nan
    if "E" in s.upper():  # notacao cientifica, ex "8,36E+16" ou "8.36E+16"
        s = s.replace(",", ".")
        try:
            v = float(s)
        except ValueError:
            return np.nan
        # Nenhum ativo individual de EFPC passa de ~R$ 1 trilhao numa unica
        # linha. Notacao cientifica acima disso é artefato do Excel, nao dado
        # real -- descarta para nao contaminar a serie.
        if abs(v) > 1e12:
            return 0.0
        return v
    partes = s.split(".")
    if len(partes) <= 2:
        try:
            return float(s)
        except ValueError:
            return np.nan
    inteiro = "".join(partes[:-1])
    decimal = partes[-1]
    try:
        return float(inteiro + "." + decimal)
    except ValueError:
        return np.nan


def carregar_serie_historica(pasta=PASTA_FUNDOS_EXCLUSIVOS):
    """Le todos os CSVs da pasta e retorna um DataFrame unico empilhado."""
    arquivos = sorted(glob.glob(os.path.join(pasta, "*.csv")))
    if not arquivos:
        raise FileNotFoundError("Nenhum .csv encontrado em: " + pasta)

    print("Lendo serie historica: " + str(len(arquivos)) + " arquivos encontrados em " + pasta)
    frames = []
    competencias_lidas = []
    arquivos_corrigidos = []

    for fp in arquivos:
        comp = _extrai_competencia(fp)
        if comp is None:
            print("  (ignorado, nome fora do padrao YYYYMM): " + os.path.basename(fp))
            continue

        # Le tudo como STRING primeiro (sem conversao automatica), para poder
        # detectar e corrigir o formato numerico antes de virar float/int.
        try:
            df = pd.read_csv(
                fp, sep=";", header=None, encoding="utf-8-sig",
                names=["DATA", "COD_EFPC", "CNPB", "TIPO_ATIVO", "VALOR", "CNPJ"],
                dtype=str, low_memory=False,
            )
        except UnicodeDecodeError:
            df = pd.read_csv(
                fp, sep=";", header=None, encoding="latin1",
                names=["DATA", "COD_EFPC", "CNPB", "TIPO_ATIVO", "VALOR", "CNPJ"],
                dtype=str, low_memory=False,
            )

        # Detecta se ESTE arquivo especifico tem o problema de pontos multiplos
        # (sintoma de arquivo que passou pelo Excel) e corrige so quando necessario.
        precisa_correcao = _tem_pontos_multiplos(df["VALOR"]) or _tem_pontos_multiplos(df["CNPB"])
        if precisa_correcao:
            arquivos_corrigidos.append(os.path.basename(fp))
            valor_str_original = df["VALOR"].copy()
            df["VALOR"] = df["VALOR"].apply(_corrige_pontos_multiplos)
            df["CNPB"] = df["CNPB"].apply(_corrige_pontos_multiplos)
            df["COD_EFPC"] = df["COD_EFPC"].apply(_corrige_pontos_multiplos)
            # Quantos valores originalmente nao-vazios/nao-zero viraram 0.0 por
            # terem sido descartados como notacao cientifica espuria?
            era_preenchido = valor_str_original.notna() & (valor_str_original.astype(str).str.strip() != "")
            n_descartados = ((df["VALOR"] == 0.0) & era_preenchido &
                              valor_str_original.astype(str).str.upper().str.contains("E", na=False)).sum()
            if n_descartados > 0:
                print("  !! " + os.path.basename(fp) + ": " + str(int(n_descartados)) +
                      " valor(es) em notacao cientifica espuria (>R$1tri numa linha) "
                      "foram descartados (zerados) por serem irrecuperaveis.")
        else:
            df["VALOR"] = pd.to_numeric(df["VALOR"], errors="coerce")
            df["CNPB"] = pd.to_numeric(df["CNPB"], errors="coerce")
            df["COD_EFPC"] = pd.to_numeric(df["COD_EFPC"], errors="coerce")

        df["COD_EFPC"] = df["COD_EFPC"].astype("Int64")
        df["CNPB"] = df["CNPB"].astype("Int64")
        df["COMPETENCIA"] = comp
        frames.append(df)
        competencias_lidas.append(comp)

    if arquivos_corrigidos:
        print("  !! AVISO: " + str(len(arquivos_corrigidos)) + " arquivo(s) com formato corrompido "
              "(multiplos pontos, sintoma de passagem pelo Excel) foram corrigidos automaticamente:")
        for nome in arquivos_corrigidos:
            print("       - " + nome)
        print("  !! Recomenda-se rebaixar esses arquivos direto da Previc quando possivel,")
        print("  !! ja que a correcao automatica e uma heuristica e pode nao ser 100% exata")
        print("  !! em casos extremos (ex: valor legitimamente >= R$ 1 trilhao em uma so linha).")
        print("")

    serie = pd.concat(frames, ignore_index=True)
    serie["VALOR"] = pd.to_numeric(serie["VALOR"], errors="coerce").fillna(0)
    serie["CLASSE"] = serie["TIPO_ATIVO"].map(CLASSE_MAP).fillna("Outros")
    serie["EH_FUNDO_EXCLUSIVO"] = serie["CNPJ"].notna() & (serie["CNPJ"].astype(str).str.strip() != "")
    serie["EH_FLUXO"] = serie["TIPO_ATIVO"].isin(TIPOS_FLUXO)

    # ── AUDITORIA: detecta arquivos/linhas com valores fora do padrão esperado ──
    # Heurística: nenhum ativo individual de EFPC costuma passar de ~R$ 50 bi (a maior
    # EFPC do pais, PREVI/BB, tem PL total na faixa de R$ 250-300 bi). Qualquer valor
    # bem acima disso numa unica linha indica problema de parsing daquele arquivo.
    LIMITE_AUDITORIA = 2e11  # R$ 200 bilhões por linha (acima do maior fundo individual conhecido, PREVI/BB)
    suspeitos = serie[serie["VALOR"].abs() > LIMITE_AUDITORIA]
    if len(suspeitos) > 0:
        print("  !! ALERTA: " + str(len(suspeitos)) + " linha(s) com valor individual acima de R$ 50 bi.")
        print("  !! Isso geralmente indica erro de parsing decimal/milhar num arquivo especifico.")
        por_comp = suspeitos.groupby("COMPETENCIA")["VALOR"].agg(["count", "max"])
        print(por_comp.to_string())
        print("  !! Amostra das piores linhas:")
        print(suspeitos.nlargest(5, "VALOR")[["COMPETENCIA", "COD_EFPC", "CNPB", "TIPO_ATIVO", "VALOR"]].to_string())
        print("")

    # ── AUDITORIA 2: total mensal muito fora da média historica (ex: >3x a mediana) ──
    tot_mensal = serie[~serie["EH_FLUXO"]].groupby("COMPETENCIA")["VALOR"].sum()
    if len(tot_mensal) > 2:
        mediana = tot_mensal.median()
        out_mes = tot_mensal[(tot_mensal > mediana * 3) | (tot_mensal < mediana / 3)]
        if len(out_mes) > 0:
            print("  !! ALERTA: competencia(s) com total mensal muito fora do padrao (mediana = "
                  + f"{mediana:,.0f})")
            print(out_mes.to_string())
            print("")

    print("  OK: " + str(len(competencias_lidas)) + " competencias lidas: " +
          str(min(competencias_lidas)) + " a " + str(max(competencias_lidas)))
    print("  OK: " + format(len(serie), ",") + " linhas totais")
    print("")
    return serie


def processar_serie_historica(serie, sg_map):
    """
    Gera os datasets agregados para o dashboard a partir da serie bruta.
    sg_map: dict {NU_MATRICULA_EFPC: SG_EFPC} vindo do Tiers.xlsx / master.
    """
    pos = serie[~serie["EH_FLUXO"]].copy()

    pl_serie = (
        pos.groupby(["COD_EFPC", "COMPETENCIA"])["VALOR"]
        .sum()
        .reset_index()
        .rename(columns={"VALOR": "VL_TOTAL"})
    )
    pl_serie["SG_EFPC"] = pl_serie["COD_EFPC"].map(sg_map).fillna(pl_serie["COD_EFPC"].astype(str))

    classe_serie = (
        pos.groupby(["COD_EFPC", "COMPETENCIA", "CLASSE"])["VALOR"]
        .sum()
        .reset_index()
    )
    classe_serie["SG_EFPC"] = classe_serie["COD_EFPC"].map(sg_map).fillna(classe_serie["COD_EFPC"].astype(str))

    pl_serie = pl_serie.sort_values(["COD_EFPC", "COMPETENCIA"])

    ultima_comp = pos["COMPETENCIA"].max()
    fe_ultima = pos[pos["COMPETENCIA"] == ultima_comp].copy()
    fe_grp = (
        fe_ultima.groupby("COD_EFPC")
        .agg(VL_TOTAL=("VALOR", "sum"),
             VL_EXCLUSIVO=("VALOR", lambda s: s[fe_ultima.loc[s.index, "EH_FUNDO_EXCLUSIVO"]].sum()),
             QT_CNPJ_EXCLUSIVO=("CNPJ", lambda s: s[fe_ultima.loc[s.index, "EH_FUNDO_EXCLUSIVO"]].nunique()))
        .reset_index()
    )
    fe_grp["PCT_EXCLUSIVO"] = np.where(fe_grp["VL_TOTAL"] > 0, fe_grp["VL_EXCLUSIVO"] / fe_grp["VL_TOTAL"] * 100, 0)
    fe_grp["SG_EFPC"] = fe_grp["COD_EFPC"].map(sg_map).fillna(fe_grp["COD_EFPC"].astype(str))

    mercado_serie = pos.groupby(["COMPETENCIA", "CLASSE"])["VALOR"].sum().reset_index()

    def tj(df):
        df2 = df.replace([np.inf, -np.inf], np.nan).astype(object)
        df2 = df2.where(pd.notna(df2), None)
        return df2.to_dict(orient="records")

    print("  OK serie processada: " + str(pl_serie["COD_EFPC"].nunique()) + " EFPCs, " +
          str(pl_serie["COMPETENCIA"].nunique()) + " competencias")
    print("  OK ultima competencia (fundos exclusivos): " + str(ultima_comp))
    print("")

    return {
        "hist_pl":       tj(pl_serie),
        "hist_classe":   tj(classe_serie),
        "hist_exclusivo": tj(fe_grp),
        "hist_mercado":  tj(mercado_serie),
        "hist_competencias": sorted(pl_serie["COMPETENCIA"].unique().tolist()),
    }
