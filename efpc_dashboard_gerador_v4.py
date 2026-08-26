"""
EFPC Intelligence Dashboard v2 — Design System Quando
=======================================================
Lê planilhas Previc + cadastros e gera HTML interativo.

Dependências:  pip install pandas openpyxl
Uso:           python efpc_dashboard_gerador_v2.py
"""

import glob
import json
import os
import re
import shutil
import sys
import unicodedata
from datetime import datetime, timedelta

import numpy as np
import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from modulo_serie_historica import carregar_serie_historica, processar_serie_historica
from modulo_tratamento_balancetes import gerar_bases_tratadas
from fichas_efpc import FICHAS_EFPC, FICHAS_ATUALIZADO_EM

# ─────────────────────────────────────────────────────────────────
# CAMINHOS — relativo à pasta onde este script está, para funcionar
# tanto localmente (Windows) quanto no GitHub Actions (Linux). As pastas
# "Dados Extração", "Dados Tratados", "Dados Cadastrais", "Dados Manuais",
# "Dados Participantes" e "arquivo_morto" devem ficar todas ao lado deste
# arquivo — mesma estrutura de antes, só que agora dentro do repositório.
# ─────────────────────────────────────────────────────────────────
BASE = os.path.dirname(os.path.abspath(__file__))


def _arquivo_mais_recente_por_prefixo(pasta, prefixo):
    """Localiza, em `pasta`, o .xlsx '{prefixo}*.xlsx' mais recente.

    Ordena primeiro pela competência codificada no nome (padrão '...-MM-AA.xlsx',
    ex.: 'BALANCETES CONSOLIDADOS-03-26.xlsx' → 03/2026); quando o nome não traz
    uma competência reconhecível, cai para a data de modificação do arquivo como
    critério (tanto de fallback quanto de desempate entre nomes com a mesma
    competência).
    """
    candidatos = glob.glob(os.path.join(pasta, f"{prefixo}*.xlsx"))
    if not candidatos:
        raise FileNotFoundError(f"Nenhum arquivo '{prefixo}*.xlsx' encontrado em: {pasta}")

    def chave(caminho):
        m = re.search(r"(\d{1,2})-(\d{2,4})(?=\.xlsx$)", os.path.basename(caminho))
        if m:
            mes, ano = int(m.group(1)), int(m.group(2))
            if ano < 100:
                ano += 2000
            return (1, ano, mes, os.path.getmtime(caminho))
        return (0, 0, 0, os.path.getmtime(caminho))

    candidatos.sort(key=chave)
    return candidatos[-1]


def _competencia_do_arquivo(caminho):
    """Extrai (mes, ano) do nome do arquivo, padrão '...-MM-AA.xlsx' (mesmo regex de
    `_arquivo_mais_recente_por_prefixo`). Retorna None se o nome não trouxer competência
    reconhecível."""
    m = re.search(r"(\d{1,2})-(\d{2,4})(?=\.xlsx$)", os.path.basename(caminho))
    if not m:
        return None
    mes, ano = int(m.group(1)), int(m.group(2))
    if ano < 100:
        ano += 2000
    return (mes, ano)


def _fator_anualizacao(mes):
    """VL_SALDO_FINAL de contas de despesa no balancete é o saldo acumulado desde
    janeiro (YTD), não uma despesa mensal isolada — projeta linearmente para uma base
    de 12 meses a partir do mês de competência do balancete (ex.: competência de março,
    1º tri → x4; junho, 1º sem → x2; setembro, 3 tri → x4/3; dezembro, ano fechado → x1)."""
    if not mes or mes <= 0 or mes > 12:
        return 1.0
    return 12.0 / mes


def _arquivo_mais_recente_dsi(pasta):
    """Localiza o 'DSI_AAAA.xlsx/csv' mais recente (maior ano no nome) na pasta —
    mesma ideia de _arquivo_mais_recente_por_prefixo, mas o padrão de nome do DSI é só
    o ano (sem mês)."""
    candidatos = glob.glob(os.path.join(pasta, "DSI_*.xlsx")) + glob.glob(os.path.join(pasta, "DSI_*.csv"))
    if not candidatos:
        raise FileNotFoundError(f"Nenhum arquivo 'DSI_*.xlsx/csv' encontrado em: {pasta}")

    def chave(caminho):
        m = re.search(r"(\d{4})", os.path.basename(caminho))
        if m:
            return (1, int(m.group(1)), os.path.getmtime(caminho))
        return (0, 0, os.path.getmtime(caminho))

    candidatos.sort(key=chave)
    return candidatos[-1]


def _resolver_ou_none(func, *args):
    """Roda `func(*args)` e devolve None (com aviso) em vez de derrubar a importação
    do módulo se não achar nada ainda — cenário normal na primeira execução, antes de
    'Dados Tratados' ter qualquer .xlsx gerado. O __main__ sempre reatribui essas
    chaves de ARQUIVOS logo depois de rodar gerar_bases_tratadas(); esse fallback só
    importa pra quem importa o módulo sem passar por ali (ex.: scripts de validação)."""
    try:
        return func(*args)
    except FileNotFoundError as e:
        print(f"  ⚠ {e} (ok se for a primeira execução — gerar_bases_tratadas() resolve isso no __main__)")
        return None


ARQUIVOS = {
    "consolidado":   _resolver_ou_none(_arquivo_mais_recente_por_prefixo,
                         os.path.join(BASE, "Dados Tratados", "Balancetes Consolidado Tratados"),
                         "BALANCETES CONSOLIDADOS"),
    "pga":           _resolver_ou_none(_arquivo_mais_recente_por_prefixo,
                         os.path.join(BASE, "Dados Tratados", "Balancetes PGAs Tratados"),
                         "BALANCETES PGA"),
    "tiers":         os.path.join(BASE, "Dados Manuais", "Tiers.xlsx"),
    "classificacao": os.path.join(BASE, "Dados Manuais", "Classificação de Dados no Balancete.xlsx"),
    "dsi":           _arquivo_mais_recente_dsi(os.path.join(BASE, "Dados Participantes", "Sexo e Idade dos participantes")),
    "dirigentes":    os.path.join(BASE, "Dados Cadastrais", "Cadastro Dirigentes.csv"),
    "cad_efpc":      os.path.join(BASE, "Dados Cadastrais", "Cadastro EFPC.csv"),
    "planos":        _resolver_ou_none(_arquivo_mais_recente_por_prefixo,
                         os.path.join(BASE, "Dados Tratados", "Balancetes Planos Tratados"),
                         "BALANCETES PLANOS"),
    "cad_planos":    os.path.join(BASE, "Dados Cadastrais", "Cadastro Planos.csv"),
}

# Pasta com a série histórica de investimentos (Demonstrativo de Investimentos
# mensal). Cada arquivo .csv tem o nome no padrão YYYYMM, sem cabeçalho.
PASTA_FUNDOS_EXCLUSIVOS = os.path.join(BASE, "Fundos Exclusivos")

# Pasta com a série histórica de situação de participantes (EPB). Cada semestre é um
# arquivo separado — EPB_1SEMESTRE_AAAA / EPB_2SEMESTRE_AAAA, em .xlsx ou .csv (sem
# cabeçalho, ";", UTF-8 com BOM) — que carregar_epb_serie() concatena e recorta para
# uma janela móvel dos últimos N anos.
PASTA_EPB = os.path.join(BASE, "Dados Participantes", "Situação do Participante")

# ─────────────────────────────────────────────────────────────────
# ROTEADOR DE "Dados Extração" — qualquer arquivo solto na raiz (fora das 3 subpastas
# de balancetes, que continuam com o tratamento dedicado de sempre) é reconhecido pelo
# nome e movido pro caminho onde o resto do gerador espera encontrá-lo. Cadastros/Tiers/
# Classificação/DSI são um arquivo único sempre substituído: se já existe um no destino,
# a versão antiga vai pra "arquivo_morto" antes de entrar a nova. Fundos Exclusivos/EPB
# são séries datadas: uma competência nova só entra (nada pra arquivar); uma competência
# repetida (reenvio/correção) substitui e arquiva a antiga do mesmo jeito.
# ─────────────────────────────────────────────────────────────────
PASTA_ARQUIVO_MORTO = os.path.join(BASE, "Arquivo_Morto")
SUBPASTAS_BALANCETES = {"Balancetes Consolidados", "Balancetes PGA", "Balancetes de Planos"}


def _normalizar_nome(nome):
    """minúsculas, sem acento, sem extensão — pra casar palavra-chave independente de
    maiúscula/acento/formato do arquivo que chegou."""
    base = os.path.splitext(nome)[0]
    base = unicodedata.normalize("NFKD", base).encode("ascii", "ignore").decode("ascii")
    return base.lower()


def _regras_roteamento_extracao():
    """Cada regra: (reconhece(nome_normalizado) -> bool, pasta_destino, nome_fixo).
    nome_fixo=None mantém o nome do arquivo que chegou (dados datados: DSI/Fundos
    Exclusivos/EPB, onde o nome É a competência); nos demais (arquivo único, sempre
    substituído) o nome fixo é o que o resto do gerador espera encontrar. Checadas em
    ordem — a primeira que bater vence — então as mais específicas vêm primeiro e o
    padrão genérico de data (Fundos Exclusivos) fica por último.
    """
    return [
        (lambda n: "dirigente" in n,
         os.path.join(BASE, "Dados Cadastrais"), "Cadastro Dirigentes.csv"),
        (lambda n: "efpc" in n and "cadastro" in n,
         os.path.join(BASE, "Dados Cadastrais"), "Cadastro EFPC.csv"),
        (lambda n: "plano" in n and "cadastro" in n,
         os.path.join(BASE, "Dados Cadastrais"), "Cadastro Planos.csv"),
        (lambda n: "tier" in n,
         os.path.join(BASE, "Dados Manuais"), "Tiers.xlsx"),
        (lambda n: "classifica" in n,
         os.path.join(BASE, "Dados Manuais"), "Classificação de Dados no Balancete.xlsx"),
        (lambda n: "dsi" in n,
         os.path.join(BASE, "Dados Participantes", "Sexo e Idade dos participantes"), None),
        (lambda n: "epb" in n and "semestre" in n,
         PASTA_EPB, None),
        (lambda n: re.search(r"20\d{2}[-_]?(0[1-9]|1[0-2])(?!\d)", n) is not None,
         PASTA_FUNDOS_EXCLUSIVOS, None),
    ]


def _mover_com_arquivamento(origem, pasta_destino, nome_destino=None):
    """Move `origem` para pasta_destino (criando-a se preciso). Se já existir um
    arquivo com esse nome no destino, arquiva o antigo em arquivo_morto (espelhando a
    subpasta de destino e carimbando com data/hora, pra nunca perder uma versão
    anterior) antes de colocar o novo no lugar; se não existir, só entra — é o caso
    normal de uma competência nova numa série histórica."""
    os.makedirs(pasta_destino, exist_ok=True)
    nome_final = nome_destino or os.path.basename(origem)
    caminho_destino = os.path.join(pasta_destino, nome_final)
    if os.path.exists(caminho_destino):
        pasta_morto = os.path.join(PASTA_ARQUIVO_MORTO, os.path.relpath(pasta_destino, BASE))
        os.makedirs(pasta_morto, exist_ok=True)
        stem, ext = os.path.splitext(nome_final)
        carimbo = datetime.now().strftime("%Y%m%d-%H%M%S")
        caminho_morto = os.path.join(pasta_morto, f"{stem}__{carimbo}{ext}")
        shutil.move(caminho_destino, caminho_morto)
        print(f"    ↪ versão anterior arquivada em: {caminho_morto}")
    shutil.move(origem, caminho_destino)
    print(f"    ✓ {os.path.basename(origem)} → {caminho_destino}")


def rotear_dados_extracao(base_dir):
    """Varre a raiz de 'Dados Extração' (fora das subpastas de balancetes) e
    redireciona qualquer arquivo reconhecido pelo nome pro caminho certo. Arquivos não
    reconhecidos ficam parados ali, com um aviso — nada é apagado sem reconhecimento."""
    extracao = os.path.join(base_dir, "Dados Extração")
    if not os.path.isdir(extracao):
        return
    regras = _regras_roteamento_extracao()
    print("Roteando 'Dados Extração' (fora dos balancetes)...")
    algo_roteado = False
    for nome in sorted(os.listdir(extracao)):
        caminho = os.path.join(extracao, nome)
        if os.path.isdir(caminho):
            if nome not in SUBPASTAS_BALANCETES:
                print(f"  ⚠ pasta não reconhecida em 'Dados Extração', ignorada: {nome}")
            continue
        nome_norm = _normalizar_nome(nome)
        alvo = next(((pasta, nome_fixo) for reconhece, pasta, nome_fixo in regras if reconhece(nome_norm)), None)
        if alvo is None:
            print(f"  ⚠ arquivo não reconhecido em 'Dados Extração', ignorado: {nome}")
            continue
        pasta_destino, nome_fixo = alvo
        print(f"  → {nome}")
        _mover_com_arquivamento(caminho, pasta_destino, nome_fixo)
        algo_roteado = True
    if not algo_roteado:
        print("  (nada novo pra rotear)")
    print()

OUTPUT_DIR  = BASE
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "dashboard_data.json")

HOJE = pd.Timestamp(datetime.now().date())

# Ordem de colunas dos arquivos EPB (igual nos .xlsx, com cabeçalho; os .csv não têm
# cabeçalho e seguem a mesma ordem). Arquivos de anos diferentes vêm com quantidades
# de colunas ligeiramente diferentes (ex.: 2023 não tem a coluna QT_SAIDAS duplicada
# que aparece em 2024/2025) — _epb_colnames() se adapta à largura real de cada arquivo,
# preservando sempre as duas últimas (QT_ATUAL, DT_EXTRACAO).
EPB_COLS_CORE = ["DT_COMPETENCIA", "NU_MATRICULA_EFPC", "SG_EFPC", "IN_ESTAT_PLAN_EFPC",
                  "ID_PLANO", "CNPB", "ID_BENEFICIO", "NU_BENEFICIO", "NM_BENEFICIO",
                  "TE_NIVEL", "QT_ANTERIOR", "QT_ENTRADAS", "QT_SAIDAS"]
EPB_COLS_TAIL = ["QT_ATUAL", "DT_EXTRACAO"]

def _epb_colnames(n):
    extra = n - len(EPB_COLS_CORE) - len(EPB_COLS_TAIL)
    if extra < 0:
        raise ValueError(f"Arquivo EPB com apenas {n} colunas — esperado no mínimo {len(EPB_COLS_CORE) + len(EPB_COLS_TAIL)}")
    return EPB_COLS_CORE + [f"QT_SAIDAS_EXTRA_{i}" for i in range(extra)] + EPB_COLS_TAIL

def carregar_epb_serie(pasta, anos=3):
    """Lê todos os EPB_*SEMESTRE_*.xlsx/.csv da pasta e concatena numa série única,
    mantendo apenas uma janela móvel dos últimos `anos` anos de competência
    (a partir da competência mais recente encontrada nos arquivos, não da data do
    sistema — assim a janela acompanha os dados reais disponíveis)."""
    arquivos = sorted(glob.glob(os.path.join(pasta, "EPB_*SEMESTRE_*.xlsx")) +
                       glob.glob(os.path.join(pasta, "EPB_*SEMESTRE_*.csv")))
    if not arquivos:
        raise FileNotFoundError(f"Nenhum arquivo EPB_*SEMESTRE_*.xlsx/csv encontrado em {pasta}")
    partes = []
    for p in arquivos:
        if p.lower().endswith(".csv"):
            d = pd.read_csv(p, sep=";", encoding="utf-8-sig", header=None, low_memory=False)
        else:
            d = pd.read_excel(p, header=0)
        d.columns = _epb_colnames(d.shape[1])
        partes.append(d)
    df = pd.concat(partes, ignore_index=True)

    df["DT_COMPETENCIA"] = pd.to_numeric(df["DT_COMPETENCIA"], errors="coerce")
    df = df.dropna(subset=["DT_COMPETENCIA"])
    df["DT_COMPETENCIA"] = df["DT_COMPETENCIA"].astype(int)

    max_comp = df["DT_COMPETENCIA"].max()
    per_max  = pd.Period(f"{max_comp // 100}-{max_comp % 100:02d}", freq="M")
    per_min  = per_max - (anos * 12 - 1)
    limite   = int(per_min.strftime("%Y%m"))
    df = df[df["DT_COMPETENCIA"] >= limite]

    df = df.drop_duplicates(
        subset=["DT_COMPETENCIA", "NU_MATRICULA_EFPC", "IN_ESTAT_PLAN_EFPC", "CNPB", "ID_BENEFICIO", "TE_NIVEL"],
        keep="last")
    return df.reset_index(drop=True)

# ─────────────────────────────────────────────────────────────────
def carregar():
    print("Carregando arquivos...")
    print(f"  (consolidado mais recente: {os.path.basename(ARQUIVOS['consolidado'])})")
    d = {}
    for k, p in ARQUIVOS.items():
        print(f"  → {k}")
        if p.lower().endswith(".csv"):
            d[k] = pd.read_csv(p, sep=";", encoding="latin1", low_memory=False)
        else:
            d[k] = pd.read_excel(p)
    print(f"  → epb (série últimos 3 anos)")
    d["epb"] = carregar_epb_serie(PASTA_EPB, anos=3)
    print(f"  ✓ {len(d)} arquivos\n")
    return d

# ─────────────────────────────────────────────────────────────────
# Setor de Despesas — grupos e subcontas considerados (todo o restante do
# balancete é ignorado nesta análise). NUM_CONTA no formato inteiro simples,
# como armazenado na base tratada (ex.: "4.020.101.020.000,00" → 4020101020000).
# ─────────────────────────────────────────────────────────────────
DESPESAS_GRUPOS = [
    {
        "grupo": "Pessoal e Encargos",
        "num_conta": 4020101000000,
        "itens": [
            {"nome": "Dirigentes", "num_conta": 4020101020000},
            {"nome": "Pessoal Próprio", "num_conta": 4020101030000},
            {"nome": "Pessoal Cedido", "num_conta": 4020101040000},
            {"nome": "Estagiários/Jovens Aprendizes", "num_conta": 4020101050000},
            {"nome": "Mão-de-Obra Temporária", "num_conta": 4020101060000},
        ],
        "residual": "Conselheiros e comitê auditoria",
    },
    {
        "grupo": "Serviços de Terceiros",
        "num_conta": 4020104000000,
        "itens": [
            {"nome": "Serviços Contábeis", "num_conta": 4020104020000},
            {"nome": "Recursos Humanos", "num_conta": 4020104040000},
            {"nome": "Tecnologia da Informação", "num_conta": 4020104050000},
            {"nome": "Gestão/Planejamento Estratégico", "num_conta": 4020104060000},
            {"nome": "Serviços e Consultorias de Investimento", "num_conta": 4020104090000},
        ],
        "residual": "Outros serviços de terceiros (Auditorias, jurídicos e manutenções)",
    },
]


def _extrair_despesas_grupos(df_cons_x):
    """Extrai VALOR por NU_MATRICULA_EFPC/GRUPO/ITEM (+ residual) de um balancete
    consolidado qualquer, na mesma estrutura de DESPESAS_GRUPOS — sem zero-fill de
    universo (usada para montar o histórico multi-competência; a extração da
    competência corrente, em processar(), mantém seu próprio reindex para EFPCs sem
    lançamento aparecerem como 0)."""
    registros = []
    idx_vazio = pd.MultiIndex.from_arrays([[], []], names=["NU_MATRICULA_EFPC", "SG_EFPC"])
    for grp in DESPESAS_GRUPOS:
        total_serie = (df_cons_x[df_cons_x["NUM_CONTA"] == grp["num_conta"]]
                       .groupby(["NU_MATRICULA_EFPC", "SG_EFPC"])["VL_SALDO_FINAL"].sum())
        item_series = [pd.Series(dtype=float, index=idx_vazio)]
        for item in grp["itens"]:
            serie = (df_cons_x[df_cons_x["NUM_CONTA"] == item["num_conta"]]
                     .groupby(["NU_MATRICULA_EFPC", "SG_EFPC"])["VL_SALDO_FINAL"].sum())
            item_series.append(serie)
            for (mat, sg), val in serie.items():
                registros.append({"NU_MATRICULA_EFPC": mat, "SG_EFPC": sg,
                                   "GRUPO": grp["grupo"], "ITEM": item["nome"], "VALOR": val})
        # pd.concat + groupby (em vez de somar Series uma a uma com .add) evita o erro
        # "cannot join with no overlapping index names" que aparece ao combinar um
        # acumulador vazio "cru" (sem MultiIndex nomeado) com séries reais — comum em
        # competências antigas onde alguma conta do grupo não teve nenhum lançamento.
        soma_subcontas = pd.concat(item_series).groupby(level=[0, 1]).sum()
        residual_serie = total_serie.subtract(soma_subcontas, fill_value=0.0)
        for (mat, sg), val in residual_serie.items():
            registros.append({"NU_MATRICULA_EFPC": mat, "SG_EFPC": sg,
                               "GRUPO": grp["grupo"], "ITEM": grp["residual"], "VALOR": val})
    return pd.DataFrame(registros, columns=["NU_MATRICULA_EFPC", "SG_EFPC", "GRUPO", "ITEM", "VALOR"])


def carregar_historico_despesas(pasta, prefixo):
    """Varre TODOS os balancetes consolidados tratados já gerados nessa pasta (não só
    o mais recente, usado em `processar()`) para montar uma série histórica de
    despesas por competência. A cada execução do pipeline de tratamento
    (modulo_tratamento_balancetes.py), um novo arquivo '{prefixo}-MM-AA.xlsx' é escrito
    ali sem apagar os anteriores — então essa pasta já acumula um histórico mensal, do
    mesmo jeito que a pasta "Fundos Exclusivos" acumula os CSVs mensais da série de
    investimentos. Cada competência é anualizada com seu próprio fator (12/mês) antes de
    entrar na série, para que comparar entre competências compare taxas equivalentes."""
    registros = []
    for caminho in glob.glob(os.path.join(pasta, f"{prefixo}*.xlsx")):
        comp = _competencia_do_arquivo(caminho)
        if comp is None:
            continue
        mes, ano = comp
        try:
            df_x = pd.read_excel(caminho)
        except Exception as e:
            print(f"  ⚠ Falha ao ler '{os.path.basename(caminho)}' para histórico de despesas: {e}")
            continue
        desp_x = _extrair_despesas_grupos(df_x)
        if desp_x.empty:
            continue
        desp_x["VALOR"] = desp_x["VALOR"] * _fator_anualizacao(mes)
        desp_x["COMPETENCIA"] = ano * 100 + mes
        registros.append(desp_x)
    cols = ["NU_MATRICULA_EFPC", "SG_EFPC", "COMPETENCIA", "GRUPO", "ITEM", "VALOR"]
    if not registros:
        return pd.DataFrame(columns=cols)
    todos = pd.concat(registros, ignore_index=True)
    return (todos.groupby(["NU_MATRICULA_EFPC", "SG_EFPC", "COMPETENCIA", "GRUPO", "ITEM"], as_index=False)["VALOR"]
            .sum()[cols])


def processar(dfs):
    print("Processando...")
    df_cons, df_tiers = dfs["consolidado"], dfs["tiers"]
    df_dsi, df_epb = dfs["dsi"], dfs["epb"]

    # PL
    df_pl = (df_cons[df_cons["NUM_CONTA"] == 2030000000000]
             [["NU_MATRICULA_EFPC", "SG_EFPC", "VL_SALDO_FINAL"]]
             .rename(columns={"VL_SALDO_FINAL": "PL_valor"}))

    # Despesas — restrito aos 2 grupos definidos em DESPESAS_GRUPOS (Pessoal e Encargos /
    # Serviços de Terceiros). Qualquer outra conta de despesa do balancete é ignorada: o
    # dashboard passou a considerar SOMENTE estes grupos, suas subcontas e a diferença
    # residual (total do grupo − soma das subcontas), reportada como linha própria.
    contas_no_balancete = set(df_cons["NUM_CONTA"].dropna().unique())
    efpc_universo = df_pl[["NU_MATRICULA_EFPC", "SG_EFPC"]].drop_duplicates()
    idx_universo = pd.MultiIndex.from_frame(efpc_universo[["NU_MATRICULA_EFPC", "SG_EFPC"]])

    def _serie_por_conta(num_conta):
        s = (df_cons[df_cons["NUM_CONTA"] == num_conta]
             .groupby(["NU_MATRICULA_EFPC", "SG_EFPC"])["VL_SALDO_FINAL"].sum())
        return s.reindex(idx_universo, fill_value=0.0)

    registros_despesas = []
    for grp in DESPESAS_GRUPOS:
        if grp["num_conta"] not in contas_no_balancete:
            print(f"  ⚠ NUM_CONTA {grp['num_conta']} (total de '{grp['grupo']}') não encontrado no balancete consolidado mais recente")
        total_serie = _serie_por_conta(grp["num_conta"])

        soma_subcontas = pd.Series(0.0, index=idx_universo)
        for item in grp["itens"]:
            if item["num_conta"] not in contas_no_balancete:
                print(f"  ⚠ NUM_CONTA {item['num_conta']} ('{grp['grupo']}' · {item['nome']}) não encontrado no balancete consolidado mais recente")
            serie = _serie_por_conta(item["num_conta"])
            soma_subcontas = soma_subcontas.add(serie, fill_value=0.0)
            for (mat, sg), val in serie.items():
                registros_despesas.append({"NU_MATRICULA_EFPC": mat, "SG_EFPC": sg,
                                            "GRUPO": grp["grupo"], "ITEM": item["nome"],
                                            "VALOR": val, "RESIDUAL": False})

        residual_serie = total_serie.subtract(soma_subcontas, fill_value=0.0)
        for (mat, sg), val in residual_serie.items():
            registros_despesas.append({"NU_MATRICULA_EFPC": mat, "SG_EFPC": sg,
                                        "GRUPO": grp["grupo"], "ITEM": grp["residual"],
                                        "VALOR": val, "RESIDUAL": True})

    df_desp_estrutura = pd.DataFrame(registros_despesas)

    # Anualização — VL_SALDO_FINAL de conta de despesa no balancete é o saldo acumulado
    # desde janeiro (YTD), não uma despesa mensal isolada. Para comparar EFPCs (e o
    # Desp%PL) numa base equivalente a 12 meses, projeta-se linearmente pelo mês de
    # competência do balancete consolidado mais recente (março/1ºtri → x4, junho/1ºsem →
    # x2, setembro/3tri → x4/3, dezembro/ano fechado → x1). Aplicado nas linhas de
    # despesas_estrutura antes de somar o total, para que item, grupo e total fiquem
    # consistentes entre si (e com o histórico multi-competência, ver
    # carregar_historico_despesas, que anualiza cada competência com seu próprio fator).
    comp_desp = _competencia_do_arquivo(ARQUIVOS["consolidado"])
    desp_mes, desp_ano = comp_desp if comp_desp else (12, None)
    fator_anualizacao = _fator_anualizacao(desp_mes)
    if comp_desp is None:
        print("  ⚠ Não foi possível identificar o mês de competência do balancete consolidado pelo nome do arquivo — despesas não anualizadas (fator 1,0x)")
    else:
        print(f"  ✓ Despesas anualizadas: competência {desp_mes:02d}/{desp_ano} → fator {fator_anualizacao:.4f}x (12/{desp_mes})")
    df_desp_estrutura["VALOR"] = df_desp_estrutura["VALOR"] * fator_anualizacao

    df_desp_tot = (df_desp_estrutura.groupby(["NU_MATRICULA_EFPC", "SG_EFPC"])["VALOR"]
                   .sum().reset_index().rename(columns={"VALOR": "DESP_TOTAL"}))

    # Tiers + atributos de filtro
    t = df_tiers[["NU_MATRICULA_EFPC", "NM_RAZAO_SOCIAL", "Tier F3", "Tier Quando",
                  "BPO(S/N)", "Prestador", "Sistema", "Label"]].copy()
    t["Prestador"] = t["Prestador"].fillna("").astype(str).str.strip().replace({"nan": "", "0": "", "<NA>": ""})
    t["Sistema"]   = t["Sistema"].fillna("").astype(str).str.strip().replace({"nan": "", "0": "", "<NA>": ""})
    t["TIER_QUANDO"] = t["Tier Quando"].apply(lambda v: str(int(v)) if pd.notna(v) and str(v) not in ("", "nan") else "N/D")
    t["FIT_MARINA"]  = t["Label"].apply(lambda v: "S" if str(v).strip().upper() == "FIT MARINA" else "N")
    t["TEM_PRESTADOR"] = t["Prestador"].apply(lambda v: "S" if v else "N")

    master = df_pl.merge(df_desp_tot, on=["NU_MATRICULA_EFPC", "SG_EFPC"], how="left")
    master = master.merge(t, on="NU_MATRICULA_EFPC", how="left")
    master["DESP_TOTAL"]  = master["DESP_TOTAL"].fillna(0)
    master["DESP_PCT_PL"] = np.where(master["PL_valor"] > 0, master["DESP_TOTAL"] / master["PL_valor"] * 100, 0)

    # Participantes
    at = (df_dsi[df_dsi["NM_BENEFICIO"] == "Participantes Ativos"]
          .groupby("NU_MATRICULA_EFPC")["QT_PESSOAS"].sum().reset_index()
          .rename(columns={"QT_PESSOAS": "QT_ATIVOS"}))
    ass = (df_dsi[df_dsi["NM_BENEFICIO"].str.contains("Assistidos", na=False)]
           .groupby("NU_MATRICULA_EFPC")["QT_PESSOAS"].sum().reset_index()
           .rename(columns={"QT_PESSOAS": "QT_ASSISTIDOS"}))
    master = master.merge(at, on="NU_MATRICULA_EFPC", how="left").merge(ass, on="NU_MATRICULA_EFPC", how="left")
    master["QT_ATIVOS"]     = master["QT_ATIVOS"].fillna(0).astype(int)
    master["QT_ASSISTIDOS"] = master["QT_ASSISTIDOS"].fillna(0).astype(int)
    master["RAZAO_MATURIDADE"] = np.where(master["QT_ATIVOS"] > 0, master["QT_ASSISTIDOS"] / master["QT_ATIVOS"], 0)
    master = master.sort_values("PL_valor", ascending=False).reset_index(drop=True)

    for c, default in [("NM_RAZAO_SOCIAL", ""), ("Tier F3", "N/D"), ("BPO(S/N)", "N"),
                       ("Prestador", ""), ("Sistema", ""), ("TIER_QUANDO", "N/D"),
                       ("FIT_MARINA", "N"), ("TEM_PRESTADOR", "N")]:
        master[c] = master[c].fillna(default)

    # Contato (Cadastro EFPC — site, e-mail e telefone)
    df_cad_efpc = dfs["cad_efpc"]
    contato = df_cad_efpc[["NU_MATRICULA_EFPC", "TE_SITE", "TE_EMAIL", "NU_FONE", "NU_CNPJ"]].drop_duplicates("NU_MATRICULA_EFPC")
    contato = contato.rename(columns={"TE_SITE": "SITE", "TE_EMAIL": "EMAIL", "NU_FONE": "FONE", "NU_CNPJ": "CNPJ"})
    for c in ["SITE", "EMAIL", "FONE", "CNPJ"]:
        contato[c] = contato[c].fillna("").astype(str).str.strip().replace({"nan": ""})
    master = master.merge(contato, on="NU_MATRICULA_EFPC", how="left")
    for c in ["SITE", "EMAIL", "FONE", "CNPJ"]:
        master[c] = master[c].fillna("")

    # Faixa etária / evolução / investimentos / PGA
    df_faixa = df_dsi.groupby(["NU_MATRICULA_EFPC", "NM_FAIXA_ETARIA"])["QT_PESSOAS"].sum().reset_index()

    # Evolução histórica de Ativos/Assistidos — nível EFPC (toda a janela do EPB, não só a
    # competência mais recente). Formato longo (TIPO=ATIVOS/ASSISTIDOS) para alimentar tanto
    # o gráfico de evolução de ativos quanto o novo gráfico percentual ativos×assistidos.
    df_epb_efpc_hist = df_epb[(df_epb["TE_NIVEL"] == "TOTALIZADOR") & (df_epb["IN_ESTAT_PLAN_EFPC"] == "EFPC")]
    df_evo_ativos = (df_epb_efpc_hist[df_epb_efpc_hist["NM_BENEFICIO"] == "Participantes Ativos"]
                     .groupby(["NU_MATRICULA_EFPC", "SG_EFPC", "DT_COMPETENCIA"])["QT_ATUAL"].sum().reset_index())
    df_evo_ativos["TIPO"] = "ATIVOS"
    df_evo_assist = (df_epb_efpc_hist[df_epb_efpc_hist["NM_BENEFICIO"].str.contains("Aposentadoria", na=False)]
                     .groupby(["NU_MATRICULA_EFPC", "SG_EFPC", "DT_COMPETENCIA"])["QT_ATUAL"].sum().reset_index())
    df_evo_assist["TIPO"] = "ASSISTIDOS"
    df_evo = pd.concat([df_evo_ativos, df_evo_assist], ignore_index=True)
    inv_map = {1020301000000: "Títulos Públicos", 1020302000000: "Crédito Privado",
               1020303000000: "Renda Variável", 1020304000000: "Fundos",
               1020307000000: "Imóveis", 1020308000000: "Op. Participantes"}
    df_inv = df_cons[df_cons["NUM_CONTA"].isin(inv_map)].copy()
    df_inv["CLASSE"] = df_inv["NUM_CONTA"].map(inv_map)
    df_inv_grp = df_inv.groupby(["NU_MATRICULA_EFPC", "SG_EFPC", "CLASSE"])["VL_SALDO_FINAL"].sum().reset_index()

    # ── PLANOS ──
    df_bal_planos, df_cad_planos = dfs["planos"], dfs["cad_planos"]

    df_pl_plano = (df_bal_planos[df_bal_planos["NUM_CONTA"] == 2030000000000]
                   [["SG_EFPC", "NU_CNPB", "Textbox17", "SITUACAO", "VL_SALDO_FINAL"]]
                   .rename(columns={"VL_SALDO_FINAL": "PL_valor", "Textbox17": "MODALIDADE"}))

    df_plano_cad = (df_cad_planos[["NU_CNPB", "SG_PLANO", "NM_PLANO", "CS_PLANO_PATROC_INSTIT", "NU_CNPJ"]]
                    .drop_duplicates("NU_CNPB")
                    .rename(columns={"CS_PLANO_PATROC_INSTIT": "TIPO_PATROCINIO", "NU_CNPJ": "CNPJ"}))
    planos = df_pl_plano.merge(df_plano_cad, on="NU_CNPB", how="left")

    # Participantes por plano (nível "PLANO" do EPB, linhas totalizadoras) — snapshot da
    # competência mais recente da série (não soma ao longo da janela histórica)
    p_epb_hist = df_epb[(df_epb["IN_ESTAT_PLAN_EFPC"] == "PLANO") & (df_epb["TE_NIVEL"] == "TOTALIZADOR")]
    ultima_comp_epb = p_epb_hist["DT_COMPETENCIA"].max()
    p_epb = p_epb_hist[p_epb_hist["DT_COMPETENCIA"] == ultima_comp_epb]
    p_ativos = (p_epb[p_epb["NM_BENEFICIO"] == "Participantes Ativos"]
                .groupby("CNPB")["QT_ATUAL"].sum().reset_index()
                .rename(columns={"QT_ATUAL": "QT_ATIVOS", "CNPB": "NU_CNPB"}))
    p_assist = (p_epb[p_epb["NM_BENEFICIO"].str.contains("Aposentadoria", na=False)]
                .groupby("CNPB")["QT_ATUAL"].sum().reset_index()
                .rename(columns={"QT_ATUAL": "QT_ASSISTIDOS", "CNPB": "NU_CNPB"}))
    planos = planos.merge(p_ativos, on="NU_CNPB", how="left").merge(p_assist, on="NU_CNPB", how="left")
    planos["QT_ATIVOS"]    = planos["QT_ATIVOS"].fillna(0).astype(int)
    planos["QT_ASSISTIDOS"] = planos["QT_ASSISTIDOS"].fillna(0).astype(int)
    planos["QT_TOTAL"]     = planos["QT_ATIVOS"] + planos["QT_ASSISTIDOS"]

    # Evolução histórica de Ativos/Assistidos — nível PLANO (mesma janela de p_epb_hist,
    # mas sem recortar para a competência mais recente). Alimenta o filtro por plano e o
    # gráfico percentual ativos×assistidos quando a página Participantes está em modo Plano.
    p_evo_ativos = (p_epb_hist[p_epb_hist["NM_BENEFICIO"] == "Participantes Ativos"]
                    .groupby(["CNPB", "DT_COMPETENCIA"])["QT_ATUAL"].sum().reset_index()
                    .rename(columns={"CNPB": "NU_CNPB"}))
    p_evo_ativos["TIPO"] = "ATIVOS"
    p_evo_assist = (p_epb_hist[p_epb_hist["NM_BENEFICIO"].str.contains("Aposentadoria", na=False)]
                    .groupby(["CNPB", "DT_COMPETENCIA"])["QT_ATUAL"].sum().reset_index()
                    .rename(columns={"CNPB": "NU_CNPB"}))
    p_evo_assist["TIPO"] = "ASSISTIDOS"
    df_evo_planos = pd.concat([p_evo_ativos, p_evo_assist], ignore_index=True)

    # Liga ao cadastro de EFPC (matrícula) via sigla
    sg_efpc_map = master.drop_duplicates("SG_EFPC")[["SG_EFPC", "NU_MATRICULA_EFPC", "NM_RAZAO_SOCIAL", "Tier F3"]]
    planos = planos.merge(sg_efpc_map, on="SG_EFPC", how="left")
    planos["NU_MATRICULA_EFPC"] = planos["NU_MATRICULA_EFPC"].fillna(0).astype(int)
    planos["NM_PLANO"] = planos["NM_PLANO"].fillna(planos["SG_PLANO"])
    planos["MODALIDADE"] = planos["MODALIDADE"].fillna("N/D")
    planos["TIPO_PATROCINIO"] = planos["TIPO_PATROCINIO"].fillna("N/D")
    planos["CNPJ"] = planos["CNPJ"].fillna("")
    planos = planos.sort_values("PL_valor", ascending=False).reset_index(drop=True)

    # Quantidade e tipos de planos (BD/CD/CV) por EFPC
    planos_agg = (planos[planos["NU_MATRICULA_EFPC"] > 0]
                  .groupby("NU_MATRICULA_EFPC")
                  .agg(QT_PLANOS=("NU_CNPB", "count"),
                       TIPOS_PLANOS=("MODALIDADE", lambda s: ", ".join(sorted(set(s) - {"N/D"})) or "N/D"))
                  .reset_index())
    master = master.merge(planos_agg, on="NU_MATRICULA_EFPC", how="left")
    master["QT_PLANOS"] = master["QT_PLANOS"].fillna(0).astype(int)
    master["TIPOS_PLANOS"] = master["TIPOS_PLANOS"].fillna("N/D")

    # ── DIRIGENTES ──
    dd = dfs["dirigentes"].copy()
    dd = dd[dd["IN_ATIVO"].astype(str).str.strip().str.lower() == "sim"].copy()
    dd["DT_INI"] = pd.to_datetime(dd["DT_INICIO_MANDATO"], errors="coerce")
    dd["DT_FIM"] = pd.to_datetime(dd["DT_FIM_MANDATO"], errors="coerce")
    dd["DIAS_RESTANTES"] = (dd["DT_FIM"] - HOJE).dt.days
    def status(r):
        if pd.isna(r["DT_FIM"]): return "SEM_DATA"
        if r["DIAS_RESTANTES"] < 0: return "VENCIDO"
        if r["DIAS_RESTANTES"] <= 180: return "VENCENDO"
        return "VIGENTE"
    dd["STATUS_MANDATO"] = dd.apply(status, axis=1)
    sg_map = master.set_index("NU_MATRICULA_EFPC")["SG_EFPC"].to_dict()
    dd["SG"] = dd["NU_MATRICULA_EFPC"].map(sg_map).fillna(dd["SG_EFPC"])
    dir_out = pd.DataFrame({
        "nm":  dd["Nome Dirigente"].astype(str).str.title().str.strip(),
        "tp":  dd["TIPO_DIRIGENTE"].astype(str).str.strip(),
        "ef":  dd["SG"].astype(str).str.strip(),
        "mat": pd.to_numeric(dd["NU_MATRICULA_EFPC"], errors="coerce").fillna(0).astype(int),
        "ini": dd["DT_INI"].dt.strftime("%Y-%m-%d").fillna(""),
        "fim": dd["DT_FIM"].dt.strftime("%Y-%m-%d").fillna(""),
        "dias": dd["DIAS_RESTANTES"].fillna(99999).astype(int),
        "st":  dd["STATUS_MANDATO"],
        "pr":  dd["PRESIDENTE"].astype(str).str.strip().str.lower().eq("sim").map({True: "S", False: "N"}),
        "aetq": dd["AETQ"].astype(str).str.strip().str.lower().eq("sim").map({True: "S", False: "N"}),
        "rem": dd["CARGO_REMUNERADO"].astype(str).str.strip().str.lower().eq("sim").map({True: "S", False: "N"}),
        "gi":  dd["CS_GRAU_INSTRUCAO"].astype(str).str.replace("\xa0", " ").str.strip().str.title(),
    })

    def tj(df): return df.replace([np.inf, -np.inf], None).fillna(0).to_dict(orient="records")
    out = {
        "master": tj(master[["NU_MATRICULA_EFPC","SG_EFPC","NM_RAZAO_SOCIAL","CNPJ","PL_valor","DESP_TOTAL",
                             "DESP_PCT_PL","Tier F3","TIER_QUANDO","BPO(S/N)","Prestador","Sistema",
                             "TEM_PRESTADOR","FIT_MARINA","QT_ATIVOS","QT_ASSISTIDOS","RAZAO_MATURIDADE",
                             "QT_PLANOS","TIPOS_PLANOS","SITE","EMAIL","FONE"]]),
        "despesas_estrutura":  tj(df_desp_estrutura),
        "despesas_grupos_cfg": DESPESAS_GRUPOS,
        "despesas_meta":   {"mes": desp_mes, "ano": desp_ano, "fator": round(fator_anualizacao, 6)},
        "investimentos":   tj(df_inv_grp),
        "epb_evolucao":    tj(df_evo),
        "epb_evolucao_planos": tj(df_evo_planos),
        "faixa_etaria":    tj(df_faixa),
        "dirigentes":      tj(dir_out),
        "planos":          tj(planos[["NU_CNPB", "SG_PLANO", "NM_PLANO", "CNPJ", "MODALIDADE", "TIPO_PATROCINIO", "SITUACAO",
                                      "PL_valor", "SG_EFPC", "NU_MATRICULA_EFPC", "NM_RAZAO_SOCIAL",
                                      "Tier F3", "QT_ATIVOS", "QT_ASSISTIDOS", "QT_TOTAL"]]),
        "hoje":            HOJE.strftime("%Y-%m-%d"),
        "fichas":          FICHAS_EFPC,
        "fichas_data":     FICHAS_ATUALIZADO_EM,
        "meta_atualizacao": {"balancete_mes": desp_mes, "balancete_ano": desp_ano,
                              "investimentos_mes": None, "investimentos_ano": None},
    }

    # ── SÉRIE HISTÓRICA DE INVESTIMENTOS (Fundos Exclusivos) ──
    try:
        sg_map = master.set_index("NU_MATRICULA_EFPC")["SG_EFPC"].to_dict()
        serie_bruta = carregar_serie_historica(PASTA_FUNDOS_EXCLUSIVOS)
        hist = processar_serie_historica(serie_bruta, sg_map)
        out.update(hist)
        if hist["hist_competencias"]:
            ultima_comp_inv = int(max(hist["hist_competencias"]))
            out["meta_atualizacao"]["investimentos_mes"] = ultima_comp_inv % 100
            out["meta_atualizacao"]["investimentos_ano"] = ultima_comp_inv // 100
    except FileNotFoundError as e:
        print(f"  ⚠ Série histórica não encontrada, pulando: {e}\n")
        out.update({
            "hist_pl": [], "hist_classe": [], "hist_exclusivo": [],
            "hist_mercado": [], "hist_competencias": [],
        })

    # ── HISTÓRICO DE DESPESAS (múltiplas competências de balancetes consolidados) ──
    # A pasta de balancetes consolidados tratados acumula um arquivo por competência a
    # cada execução do pipeline (ver carregar_historico_despesas) — normalmente só 1-2
    # arquivos existem ainda, então esse gráfico começa esparso e cresce sozinho a cada
    # novo mês tratado, sem precisar mexer no gerador de novo.
    try:
        df_hist_desp = carregar_historico_despesas(os.path.dirname(ARQUIVOS["consolidado"]), "BALANCETES CONSOLIDADOS")
        out["hist_despesas"] = tj(df_hist_desp)
        out["hist_despesas_competencias"] = sorted(df_hist_desp["COMPETENCIA"].unique().tolist()) if len(df_hist_desp) else []
        print(f"  ✓ Histórico de despesas: {len(out['hist_despesas_competencias'])} competência(s) — {out['hist_despesas_competencias']}")
    except Exception as e:
        print(f"  ⚠ Histórico de despesas não carregado: {e}\n")
        out["hist_despesas"] = []
        out["hist_despesas_competencias"] = []

    print(f"  ✓ Master: {len(master)} EFPCs | Dirigentes ativos: {len(dir_out)}")
    print(f"  ✓ Mandatos — vigentes: {(dir_out['st']=='VIGENTE').sum()}, vencendo (≤180d): {(dir_out['st']=='VENCENDO').sum()}, vencidos: {(dir_out['st']=='VENCIDO').sum()}")
    print(f"  ✓ Planos: {len(planos)} (CD: {(planos['MODALIDADE']=='CD').sum()}, BD: {(planos['MODALIDADE']=='BD').sum()}, CV: {(planos['MODALIDADE']=='CV').sum()})\n")
    return out


# ─────────────────────────────────────────────────────────────────
# HTML — Design System Quando
# ─────────────────────────────────────────────────────────────────
def gerar_html(data):
    print("Gerando HTML...")
    data_js = json.dumps(data, ensure_ascii=False)
    html = HTML_TEMPLATE.replace("__DATA__", data_js)
    return html

def exportar_dados(data, caminho_saida):
    """
    Substitui gerar_html() no fluxo de publicação: em vez de colar o JSON
    dentro do HTML_TEMPLATE, salva só os dados processados como arquivo
    separado. O index.html (agora estático, fora deste script) faz fetch
    desse arquivo em tempo de execução.
    """
    print("Exportando dados...")
    with open(caminho_saida, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False)
    print(f"  ✓ Dados exportados: {caminho_saida}")

HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>EFPC Intelligence · Quando</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
:root{
  /* Quando Design System */
  --teal-50:#E8F7F7;--teal-100:#C8EBED;--teal-500:#1AAAB2;--teal-600:#169197;
  --teal-700:#12797E;--teal-800:#0F6165;--teal-900:#0C4C50;--teal-950:#174543;
  --navy:#153451;--navy-600:#16405B;--slate-700:#273444;--slate-900:#1A232E;
  --g800:#464646;--g700:#5F5F5F;--g600:#969696;--g500:#B3B3B3;--g400:#E0E0E0;
  --g300:#EBEBEB;--g200:#F5F5F5;--g150:#F6F6F6;--g100:#FAFAFA;--g50:#FBFBFB;
  --success:#22825D;--success-b:#30B783;--success-100:#CDEEE1;
  --danger:#B33E3E;--danger-300:#FF725E;--danger-100:#FBBCB3;
  --warning:#FF9A47;--info:#0788C9;
  --font:'Inter',sans-serif;--radius:12px;
  --shadow:0 1px 3px rgba(21,52,81,.08),0 4px 16px rgba(21,52,81,.06);
}
*{margin:0;padding:0;box-sizing:border-box}
body{background:var(--g100);color:var(--g800);font-family:var(--font);min-height:100vh;font-size:15px}
::-webkit-scrollbar{width:8px;height:8px}::-webkit-scrollbar-track{background:var(--g200)}
::-webkit-scrollbar-thumb{background:var(--g400);border-radius:4px}
::-webkit-scrollbar-thumb:hover{background:var(--g500)}
.layout{display:flex;min-height:100vh}
/* SIDEBAR — teal escuro como section divider */
.sidebar{width:232px;min-width:232px;background:var(--teal-900);display:flex;flex-direction:column;
  position:sticky;top:0;height:100vh;overflow-y:auto;z-index:100}
.sidebar ::-webkit-scrollbar-track{background:var(--teal-950)}
.sb-brand{padding:28px 24px 22px;border-bottom:1px solid rgba(255,255,255,.08)}
.sb-brand .logo{font-size:22px;font-weight:700;color:#fff;letter-spacing:-.5px}
.sb-brand .logo span{color:var(--teal-500)}
.sb-brand .sub{font-size:10px;color:rgba(255,255,255,.45);letter-spacing:.14em;text-transform:uppercase;margin-top:5px;font-weight:600}
.nav{padding:14px 0;flex:1}
.nav-sec{padding:14px 24px 6px;font-size:10px;letter-spacing:.14em;text-transform:uppercase;color:rgba(255,255,255,.35);font-weight:600}
.nav-it{display:flex;align-items:center;gap:11px;padding:10px 24px;cursor:pointer;font-size:14.5px;font-weight:500;
  color:rgba(255,255,255,.65);transition:.15s;border-left:3px solid transparent;user-select:none}
.nav-it:hover{color:#fff;background:rgba(255,255,255,.05)}
.nav-it.active{color:#fff;background:rgba(26,170,178,.18);border-left-color:var(--teal-500)}
.nav-it .ic{width:18px;text-align:center;font-size:15px;opacity:.8}
.sb-foot{padding:18px 24px;border-top:1px solid rgba(255,255,255,.08)}
.sb-foot .c{font-size:11px;color:rgba(255,255,255,.4);font-weight:500}
.main{flex:1;min-width:0;display:flex;flex-direction:column}
/* TOPBAR */
.topbar{background:#fff;border-bottom:1px solid var(--g300);padding:18px 32px 14px;position:sticky;top:0;z-index:50}
.topbar .eyebrow{font-size:12px;font-weight:600;letter-spacing:.08em;text-transform:uppercase;color:var(--teal-600);margin-bottom:3px}
.topbar h1{font-size:22px;font-weight:600;color:var(--navy);letter-spacing:-.3px}
.topbar .pg-updated{font-size:11.5px;color:var(--g500);margin-top:3px}
/* FILTER BAR */
.fbar{background:#fff;border-bottom:1px solid var(--g300);padding:12px 32px;display:flex;gap:22px;
  flex-wrap:wrap;align-items:center;position:sticky;top:73px;z-index:49}
.fgroup{display:flex;align-items:center;gap:7px;flex-wrap:wrap}
.fgroup .flabel{font-size:11.5px;font-weight:600;letter-spacing:.07em;text-transform:uppercase;color:var(--g600);margin-right:2px}
.chip{padding:4px 12px;border-radius:999px;font-size:12.5px;font-weight:600;border:1.5px solid var(--g300);
  cursor:pointer;transition:.13s;user-select:none;color:var(--g600);background:#fff}
.chip:hover{border-color:var(--teal-500);color:var(--teal-600)}
.chip.on{background:var(--teal-50);border-color:var(--teal-500);color:var(--teal-700)}
.chip.off{opacity:.45}
.fcount{margin-left:auto;font-size:13px;font-weight:600;color:var(--teal-600);background:var(--teal-50);
  padding:5px 14px;border-radius:999px;white-space:nowrap}
.freset{font-size:12.5px;font-weight:600;color:var(--g600);cursor:pointer;text-decoration:underline;white-space:nowrap}
.freset:hover{color:var(--danger-300)}
/* PAGES */
.page{display:none;padding:26px 32px;flex-direction:column;gap:20px}
.page.active{display:flex}
.efpc-bar{background:#fff;border:1px solid var(--g300);border-radius:var(--radius);box-shadow:var(--shadow);
  padding:14px 20px;display:flex;align-items:center;gap:14px;flex-wrap:wrap}
.efpc-bar label{font-size:12px;font-weight:600;letter-spacing:.06em;text-transform:uppercase;color:var(--g600)}
.inp,select{background:#fff;border:1.5px solid var(--g300);border-radius:8px;color:var(--g800);
  padding:8px 12px;font-size:14px;font-family:var(--font);outline:none;transition:.15s}
.inp:focus,select:focus{border-color:var(--teal-500);box-shadow:0 0 0 3px var(--teal-50)}
.inp{min-width:200px}select{min-width:240px;cursor:pointer}
.efpc-meta{margin-left:auto;display:flex;gap:20px;align-items:center;flex-wrap:wrap}
.em-it{display:flex;flex-direction:column;align-items:flex-end}
.em-it .l{font-size:10.5px;font-weight:600;color:var(--g500);text-transform:uppercase;letter-spacing:.08em}
.em-it .v{font-size:14.5px;color:var(--navy);font-weight:600}
/* KPI */
.kpi-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:16px}
.kpi{background:#fff;border:1px solid var(--g300);border-radius:var(--radius);box-shadow:var(--shadow);
  padding:20px 22px;display:flex;flex-direction:column;gap:7px;position:relative;overflow:hidden}
.kpi::before{content:'';position:absolute;top:0;left:0;right:0;height:3px;background:var(--teal-500)}
.kpi.teal::before{background:var(--teal-500)}.kpi.navy::before{background:var(--navy)}
.kpi.green::before{background:var(--success-b)}.kpi.red::before{background:var(--danger-300)}
.kpi.orange::before{background:var(--warning)}.kpi.blue::before{background:var(--info)}
.kpi-l{font-size:11.5px;color:var(--g600);font-weight:600;text-transform:uppercase;letter-spacing:.08em}
.kpi-v{font-size:32px;font-weight:700;color:var(--navy);line-height:1.05;letter-spacing:-.5px}
.kpi-s{font-size:12.5px;color:var(--g700);font-weight:500}
/* CHARTS */
.cg2{display:grid;grid-template-columns:1fr 1fr;gap:18px}
.cg21{display:grid;grid-template-columns:2fr 1fr;gap:18px}
@media(max-width:1100px){.cg2,.cg21{grid-template-columns:1fr}}
.cc{background:#fff;border:1px solid var(--g300);border-radius:var(--radius);box-shadow:var(--shadow);
  padding:22px;display:flex;flex-direction:column;gap:16px}
.cc-h{display:flex;align-items:baseline;gap:10px}
.cc-eyebrow{font-size:11px;font-weight:600;letter-spacing:.08em;text-transform:uppercase;color:var(--teal-600)}
.cc-t{font-size:16px;font-weight:600;color:var(--navy);letter-spacing:-.2px}
.cc-s{font-size:12px;color:var(--g600);margin-left:auto;font-weight:500}
.cw canvas{max-height:260px}.cw.tall canvas{max-height:340px}
/* LEGENDA CUSTOM — substitui a legenda nativa do Chart.js (que risca o texto ao ocultar
   uma fatia/série); aqui o item inteiro só "apaga" via opacidade, sem risco.
   .cwc dá ao canvas uma caixa com altura explícita e position:relative — sem isso, um
   <canvas> como item flex direto não tem altura estável para o Chart.js medir (é o que
   deixava os gráficos de rosca com legenda "estranhos"/mal dimensionados). */
.cw-leg-r{display:flex;align-items:center;gap:16px}
.cw-leg-r .cwc{flex:1;min-width:0}
.cw-leg-b{display:flex;flex-direction:column;gap:10px}
.cwc{position:relative;height:230px}
.cleg{display:flex;flex-wrap:wrap;gap:7px 16px}
.cleg-col{flex-direction:column;flex-wrap:nowrap;flex:0 0 auto;max-width:150px;max-height:230px;overflow-y:auto;padding-right:2px}
.cleg-it{display:flex;align-items:center;gap:7px;cursor:pointer;font-size:12.5px;font-weight:500;
  color:var(--g700);transition:opacity .15s ease;user-select:none}
.cleg-it:hover{color:var(--teal-700)}
.cleg-it.off{opacity:.35}
.cleg-dot{width:10px;height:10px;border-radius:50%;flex-shrink:0}
/* TABLE */
.tc{background:#fff;border:1px solid var(--g300);border-radius:var(--radius);box-shadow:var(--shadow);overflow:hidden}
.tc-h{padding:18px 22px;border-bottom:1px solid var(--g300);display:flex;align-items:center;gap:14px;flex-wrap:wrap}
.tc-t{font-size:16px;font-weight:600;color:var(--navy)}
.tc-search{margin-left:auto;width:220px}
.tw{overflow-x:auto;max-height:560px;overflow-y:auto}
table{width:100%;border-collapse:collapse;font-size:13.5px}
thead tr{background:var(--g150);position:sticky;top:0;z-index:5}
th{padding:11px 16px;text-align:left;color:var(--g600);font-size:11px;letter-spacing:.07em;font-weight:600;
  text-transform:uppercase;cursor:pointer;user-select:none;white-space:nowrap;border-bottom:1px solid var(--g300)}
th:hover{color:var(--teal-600)}
th.th-sort-active{color:var(--teal-700)}
td{padding:10px 16px;border-bottom:1px solid var(--g200);color:var(--g700);white-space:nowrap}
tbody tr:hover td{background:var(--teal-50);cursor:pointer}
tbody tr:last-child td{border-bottom:none}
.tag{display:inline-block;padding:3px 10px;border-radius:999px;font-size:11px;font-weight:600;letter-spacing:.02em}
.desp-grupo-row td{background:var(--g150);font-weight:600;color:var(--navy)}
.desp-grupo-row:hover td{background:var(--teal-50)}
.desp-caret{display:inline-block;width:14px;color:var(--teal-600);font-size:10px}
.desp-evo-ic{cursor:pointer;margin-left:6px;font-size:11px;opacity:.55;transition:.13s}
.desp-evo-ic:hover{opacity:1}
.desp-sub-row td{color:var(--g700)}
.desp-sub-row:hover td{cursor:default;background:transparent}
.desp-residual td{font-style:italic;color:var(--g600)}
.num{font-variant-numeric:tabular-nums;text-align:right;font-weight:500}
.num.pos{color:var(--success)}.num.neg{color:var(--danger-300)}.num.neu{color:var(--g600)}
.tabs{display:flex;gap:4px;background:var(--g200);padding:4px;border-radius:9px;width:fit-content;flex-wrap:wrap}
.tab{padding:6px 15px;border-radius:6px;font-size:13px;font-weight:600;cursor:pointer;color:var(--g600);transition:.13s}
.tab:hover{color:var(--navy)}
.tab.active{background:#fff;color:var(--teal-700);box-shadow:0 1px 3px rgba(21,52,81,.12)}
/* INSIGHT CARD */
.insight{background:var(--teal-50);border:1px solid var(--teal-100);border-left:4px solid var(--teal-500);
  border-radius:10px;padding:16px 20px;display:flex;gap:14px;align-items:flex-start}
.insight.warn{background:#FFF4EA;border-color:#FFE0C2;border-left-color:var(--warning)}
.insight.danger{background:#FDEEEC;border-color:var(--danger-100);border-left-color:var(--danger-300)}
.insight .ico{font-size:18px;line-height:1.3}
.insight .tt{font-size:14px;font-weight:600;color:var(--navy);margin-bottom:3px}
.insight .tx{font-size:13.5px;color:var(--g700);line-height:1.5}
.stbadge{display:inline-flex;align-items:center;gap:5px;padding:3px 10px;border-radius:999px;font-size:11.5px;font-weight:600}
.st-vigente{background:var(--success-100);color:var(--success)}
.st-vencendo{background:#FFF0E0;color:#C26A14}
.st-vencido{background:var(--danger-100);color:var(--danger)}
.st-semdata{background:var(--g200);color:var(--g600)}
.ac-wrap{position:relative;display:inline-flex;flex-direction:column}
.ac-list{position:absolute;top:calc(100% + 4px);left:0;min-width:300px;background:#fff;
  border:1.5px solid var(--teal-500);border-radius:10px;max-height:320px;overflow-y:auto;
  z-index:500;box-shadow:0 4px 20px rgba(21,52,81,.15);display:none}
.ac-list.open{display:block}
.ac-item{display:flex;align-items:center;padding:9px 14px;cursor:pointer;font-size:14px;color:var(--g700);transition:.1s;border-bottom:1px solid var(--g200)}
.ac-item:last-child{border-bottom:none}
.ac-item:hover{background:var(--teal-50);color:var(--teal-700);font-weight:500}
.ac-item.checked{background:var(--teal-50);color:var(--teal-700);font-weight:600}
.ac-item input[type=checkbox]{margin-right:10px;accent-color:var(--teal-500);cursor:pointer;flex-shrink:0}
.ac-caret{display:inline-block;width:14px;text-align:center;color:var(--teal-600);font-size:9px;cursor:pointer;flex-shrink:0;margin-right:2px}
.ac-caret-empty{visibility:hidden;cursor:default}
.ac-sub{padding-left:38px;font-size:12.5px;color:var(--g600);background:var(--g100)}
.ac-sub.checked{color:var(--teal-700);background:var(--teal-50)}
.ac-empty{padding:9px 14px;font-size:13px;color:var(--g500);font-style:italic}
.ms-chips{display:flex;gap:6px;flex-wrap:wrap;align-items:center}
.ms-chip{display:inline-flex;align-items:center;gap:6px;background:var(--teal-50);border:1px solid var(--teal-100);
  color:var(--teal-700);font-size:12.5px;font-weight:600;padding:4px 8px 4px 11px;border-radius:999px}
.ms-x{cursor:pointer;color:var(--teal-600);font-size:10px;line-height:1}
.ms-x:hover{color:var(--danger-300)}
/* FICHA EFPC — MODAL */
.ef-overlay{display:none;position:fixed;inset:0;background:rgba(21,52,81,.55);backdrop-filter:blur(3px);
  align-items:center;justify-content:center;z-index:1000;padding:24px}
.ef-overlay.open{display:flex}
.ef-modal{background:#fff;border-radius:var(--radius);box-shadow:0 12px 48px rgba(21,52,81,.35);
  width:100%;max-width:900px;max-height:88vh;display:flex;flex-direction:column;overflow:hidden}
.ef-head{padding:22px 26px 16px;border-bottom:1px solid var(--g300);display:flex;align-items:flex-start;
  justify-content:space-between;gap:16px}
.ef-eyebrow{font-size:12px;font-weight:600;letter-spacing:.08em;text-transform:uppercase;color:var(--teal-600);margin-bottom:4px}
.ef-title{font-size:20px;font-weight:600;color:var(--navy);letter-spacing:-.3px;margin-bottom:8px}
.ef-badges{display:flex;gap:6px;flex-wrap:wrap}
.ef-close{cursor:pointer;font-size:16px;color:var(--g600);width:30px;height:30px;display:flex;
  align-items:center;justify-content:center;border-radius:8px;transition:.15s;flex-shrink:0}
.ef-close:hover{background:var(--g200);color:var(--navy)}
.ef-tabs{margin:16px 26px 0}
.ef-body{padding:20px 26px 26px;overflow-y:auto;flex:1;display:flex;flex-direction:column;gap:18px}
.ef-pane{display:none;flex-direction:column;gap:18px}
.ef-pane.active{display:flex}
.ef-section-t{font-size:12px;font-weight:600;letter-spacing:.07em;text-transform:uppercase;color:var(--g600);margin-bottom:8px}
.ef-desc-text{font-size:14px;color:var(--g700);line-height:1.6}
.ef-fonte{font-size:12px;color:var(--g500);font-style:italic;border-top:1px solid var(--g200);padding-top:12px}
.ef-open{cursor:pointer;color:var(--teal-600)}
.ef-open:hover{text-decoration:underline}
.ef-link{font-size:13px;font-weight:600;color:var(--teal-600);cursor:pointer;text-decoration:none;white-space:nowrap}
.ef-link:hover{text-decoration:underline}
/* DADOS MANUAIS — EDIÇÃO (desativado, ver comentário no bloco HTML/JS correspondente)
.man-cell{border:1px solid transparent;background:transparent;font-family:var(--font);font-size:12.5px;
  color:var(--g700);padding:5px 7px;width:100%;min-width:110px;border-radius:5px;outline:none}
.man-cell:hover{border-color:var(--g300)}
.man-cell:focus{border-color:var(--teal-500);background:#fff;box-shadow:0 0 0 2px var(--teal-50)}
.man-del{color:var(--danger-300);cursor:pointer;font-size:11px;font-weight:600;white-space:nowrap}
.man-del:hover{text-decoration:underline}
*/
</style>
</head>
<body>
<div class="layout">
<nav class="sidebar">
  <div class="sb-brand">
    <div class="logo">EFPC<span>intel</span></div>
    <div class="sub" id="sb-sub">Quando · Previc</div>
  </div>
  <div class="nav">
    <div class="nav-sec">Visão Geral</div>
    <div class="nav-it active" data-page="overview" onclick="navTo('overview',this)"><span class="ic">◈</span> Painel Geral</div>
    <div class="nav-it" data-page="ranking" onclick="navTo('ranking',this)"><span class="ic">▤</span> Ranking EFPCs</div>
    <div class="nav-sec">Análise Financeira</div>
    <div class="nav-it" data-page="despesas" onclick="navTo('despesas',this)"><span class="ic">◉</span> Despesas</div>
    <div class="nav-it" data-page="investimentos" onclick="navTo('investimentos',this)"><span class="ic">◎</span> Investimentos</div>
    <div class="nav-sec">Pessoas</div>
    <div class="nav-it" data-page="participantes" onclick="navTo('participantes',this)"><span class="ic">◌</span> Participantes</div>
    <div class="nav-it" data-page="dirigentes" onclick="navTo('dirigentes',this)"><span class="ic">♟</span> Dirigentes</div>
    <div class="nav-sec">Planos</div>
    <div class="nav-it" data-page="planos" onclick="navTo('planos',this)"><span class="ic">▦</span> Planos por EFPC</div>
    <!-- Dados Manuais desativado a pedido do usuário — ver bloco comentado em page-manuais e nas funções man*
    <div class="nav-sec">Dados</div>
    <div class="nav-it" data-page="manuais" onclick="navTo('manuais',this)"><span class="ic">✎</span> Dados Manuais</div>
    -->
  </div>
  <div class="sb-foot">
    <div class="c"> </div>
    <div class="c" style="margin-top:3px" id="side-count">—</div>
  </div>
</nav>
<div class="main">
  <div class="topbar">
    <div class="eyebrow" id="pg-eyebrow">Visão Geral</div>
    <h1 id="pg-title">Painel Geral</h1>
    <div class="pg-updated" id="pg-updated"></div>
  </div>
  <!-- FILTROS GLOBAIS -->
  <div class="fbar">
    <div class="fgroup"><span class="flabel">Tier F3</span>
      <span class="chip on" data-f="f3" data-v="CONSULTING" onclick="tg(this)">Consulting</span>
      <span class="chip on" data-f="f3" data-v="OCIO I" onclick="tg(this)">OCIO I</span>
      <span class="chip on" data-f="f3" data-v="OCIO II" onclick="tg(this)">OCIO II</span>
      <span class="chip on" data-f="f3" data-v="OCIO III" onclick="tg(this)">OCIO III</span>
      <span class="chip on" data-f="f3" data-v="OCIO IV" onclick="tg(this)">OCIO IV</span>
      <span class="chip on" data-f="f3" data-v="OCIO V" onclick="tg(this)">OCIO V</span>
    </div>
    <div class="fgroup"><span class="flabel">Marina</span>
      <span class="chip on" data-f="tq" data-v="1" onclick="tg(this)">1</span>
      <span class="chip on" data-f="tq" data-v="2" onclick="tg(this)">2</span>
      <span class="chip on" data-f="tq" data-v="3" onclick="tg(this)">3</span>
      <span class="chip on" data-f="tq" data-v="N/D" onclick="tg(this)">N/D</span>
    </div>
    <div class="fgroup"><span class="flabel">BPO</span>
      <span class="chip on" data-f="bpo" data-v="S" onclick="tg(this)">Sim</span>
      <span class="chip on" data-f="bpo" data-v="N" onclick="tg(this)">Não</span>
    </div>
    <span class="freset" onclick="resetF()">limpar</span>
    <div class="fcount" id="fcount">—</div>
  </div>

  <!-- OVERVIEW -->
  <div class="page active" id="page-overview">
    <div class="kpi-grid" id="kpi-overview"></div>
    <div class="tabs" id="disp-tabs">
      <div class="tab active" onclick="setDispDim('tierf3',this)">Tier F3</div>
      <div class="tab" onclick="setDispDim('tierq',this)">Marina</div>
      <div class="tab" onclick="setDispDim('bpo',this)">BPO</div>
      <div class="tab" onclick="setDispDim('prestador',this)">Prestador</div>
    </div>
    <div class="cg2">
      <div class="cc"><div class="cc-h"><div class="cc-t" id="disp-t1">Distribuição por Tier F3</div><div class="cc-s">por Ativos Sob Gestão</div></div><div class="cw cw-leg-r"><div class="cwc"><canvas id="ch-tier-pl"></canvas></div><div class="cleg cleg-col" id="ch-tier-pl-leg"></div></div></div>
      <div class="cc"><div class="cc-h"><div class="cc-t" id="disp-t2">Despesa Total por Tier F3</div><div class="cc-s">R$ milhões</div></div><div class="cw"><canvas id="ch-tier-desp"></canvas></div></div>
    </div>
    <div class="cg21">
      <div class="cc"><div class="cc-h"><div class="cc-t">Top 15 EFPCs · Ativos Sob Gestão</div><div class="cc-s">R$ bilhões</div></div><div class="cw tall"><canvas id="ch-top-pl"></canvas></div></div>
      <div class="cc"><div class="cc-h"><div class="cc-t">Despesa % PL</div><div class="cc-s">histograma</div></div><div class="cw tall"><canvas id="ch-hist"></canvas></div></div>
    </div>
  </div>

  <!-- RANKING -->
  <div class="page" id="page-ranking">
    <div class="tc">
      <div class="tc-h">
        <div class="tc-t">Ranking de EFPCs <span class="fcount" id="rank-count" style="margin-left:6px">—</span></div>
        <input class="inp tc-search" id="rank-q" placeholder="Buscar EFPC..." oninput="rRanking()">
        <span class="chip" onclick="exportTable('ranking','csv')">⬇ CSV</span>
        <span class="chip" onclick="exportTable('ranking','xlsx')">⬇ XLSX</span>
      </div>
      <div class="tw"><table><thead><tr id="rank-thead-row">
        <th style="cursor:default">#</th>
        <th data-key="SG_EFPC" data-label="EFPC" onclick="setRank('SG_EFPC',this)">EFPC</th>
        <th data-key="NM_RAZAO_SOCIAL" data-label="Razão Social" onclick="setRank('NM_RAZAO_SOCIAL',this)">Razão Social</th>
        <th data-key="CNPJ" data-label="CNPJ" onclick="setRank('CNPJ',this)">CNPJ</th>
        <th data-key="Tier F3" data-label="Tier F3" onclick="setRank('Tier F3',this)">Tier F3</th>
        <th data-key="TIER_QUANDO" data-label="Marina" onclick="setRank('TIER_QUANDO',this)">Marina</th>
        <th data-key="BPO(S/N)" data-label="BPO" onclick="setRank('BPO(S/N)',this)">BPO</th>
        <th data-key="Prestador" data-label="Prestador" onclick="setRank('Prestador',this)">Prestador</th>
        <th data-key="Sistema" data-label="Sistema" onclick="setRank('Sistema',this)">Sistema</th>
        <th class="num" data-key="PL_valor" data-label="PL" onclick="setRank('PL_valor',this)">PL</th>
        <th class="num" data-key="DESP_TOTAL" data-label="Despesas" onclick="setRank('DESP_TOTAL',this)">Despesas</th>
        <th class="num" data-key="DESP_PCT_PL" data-label="Desp%PL" onclick="setRank('DESP_PCT_PL',this)">Desp%PL</th>
        <th class="num" data-key="QT_ATIVOS" data-label="Ativos" onclick="setRank('QT_ATIVOS',this)">Ativos</th>
        <th class="num" data-key="RAZAO_MATURIDADE" data-label="Maturidade" onclick="setRank('RAZAO_MATURIDADE',this)">Maturidade</th>
        <th class="num" data-key="QT_PLANOS" data-label="Planos" onclick="setRank('QT_PLANOS',this)">Planos</th>
        <th style="cursor:default">Contato</th>
      </tr></thead><tbody id="rank-body"></tbody></table></div>
    </div>
  </div>

  <!-- DESPESAS -->
  <div class="page" id="page-despesas">
    <div class="efpc-bar"><label>Buscar EFPC</label>
      <div class="ac-wrap">
        <input class="inp" id="desp-q" placeholder="Marque uma ou mais EFPCs..." oninput="msRender('desp')" onfocus="msRender('desp')" autocomplete="off">
        <div class="ac-list" id="desp-q-ac"></div>
      </div>
      <div class="ms-chips" id="desp-chips"></div>
      <span class="freset" id="desp-clear" style="display:none" onclick="msClear('desp')">limpar seleção</span>
      <a class="ef-link" id="desp-ficha-link" style="display:none" onclick="openFicha([...MS.desp][0])">🔍 Ver ficha completa</a>
      <div class="efpc-meta" id="desp-meta"></div>
    </div>
    <div id="desp-ann-note" style="font-size:11.5px;color:var(--g500);font-style:italic"></div>
    <div class="kpi-grid" id="kpi-desp"></div>
    <div class="cc">
      <div class="cc-h"><div class="cc-t">Comparação com o Tier</div><div class="cc-s" id="desp-tier-sub">Desp%PL vs. outras EFPCs do mesmo Tier</div>
        <div class="tabs" id="desp-cmp-tabs" style="margin-left:8px;align-self:center">
          <div class="tab active" data-cmp="f3" onclick="setDespCmp('f3',this)">Tier F3</div>
          <div class="tab" data-cmp="tq" onclick="setDespCmp('tq',this)">Marina</div>
        </div>
      </div>
      <div id="desp-tier-empty" style="display:none;color:var(--g500);font-size:12.5px;font-style:italic">Selecione uma única EFPC (acima) para comparar o Desp%PL dela com as demais do mesmo Tier.</div>
      <div id="desp-tier-body" style="display:flex;flex-direction:column;gap:14px">
        <div class="insight" id="desp-tier-insight" style="display:none"></div>
        <div class="cw tall"><canvas id="ch-desp-tier"></canvas></div>
      </div>
    </div>
    <div class="tc">
      <div class="tc-h"><div class="tc-t">Despesas por Grupo</div><div class="cc-s" style="margin-left:auto">clique no grupo para ver as subcontas · clique no 📈 de uma conta para ver a evolução dela · selecione 1 EFPC para comparar com a mediana do Tier</div>
        <span class="chip" onclick="exportDespesasXlsx()">⬇ XLSX · Breakdown por EFPC</span>
      </div>
      <div class="tw">
        <table>
          <thead><tr><th>Conta</th><th class="num">Valor</th><th class="num">% do Grupo</th><th class="num">% da Despesa Total</th><th class="num">% do AuM</th><th class="num" id="desp-th-mediana">Mediana Tier F3 (%AuM)</th><th class="num">Vs. Mediana</th></tr></thead>
          <tbody id="desp-grupos-tbody"></tbody>
        </table>
      </div>
    </div>
    <div class="cc" id="desp-evo-card" style="display:none">
      <div class="cc-h"><div class="cc-t" id="desp-evo-t">Evolução da Despesa</div><div class="cc-s" id="desp-evo-s">valores anualizados · por competência do balancete</div></div>
      <div id="desp-evo-empty" style="display:none;color:var(--g500);font-size:12.5px;font-style:italic">Ainda não há competências anteriores tratadas para comparar — este gráfico cresce sozinho a cada balancete consolidado novo processado.</div>
      <div class="cw tall"><canvas id="ch-desp-evo"></canvas></div>
    </div>
  </div>

  <!-- INVESTIMENTOS (alocação atual + série histórica) -->
  <div class="page" id="page-investimentos">
    <div class="efpc-bar"><label>Buscar EFPC</label>
      <div class="ac-wrap">
        <input class="inp" id="inv-q" placeholder="Marque uma ou mais EFPCs..." oninput="msRender('inv')" onfocus="msRender('inv')" autocomplete="off">
        <div class="ac-list" id="inv-q-ac"></div>
      </div>
      <div class="ms-chips" id="inv-chips"></div>
      <span class="freset" id="inv-clear" style="display:none" onclick="msClear('inv')">limpar seleção</span>
      <a class="ef-link" id="inv-ficha-link" style="display:none" onclick="openFicha([...MS.inv][0])">🔍 Ver ficha completa</a>
      <span class="chip" onclick="exportInvestimentosXlsx()">⬇ XLSX · Breakdown por EFPC</span>
      <div class="efpc-meta" id="inv-meta"></div>
    </div>
    <div id="hist-insights" style="display:flex;flex-direction:column;gap:10px"></div>
    <div class="kpi-grid" id="kpi-hist"></div>
    <div class="cg2">
      <div class="cc"><div class="cc-h"><div class="cc-t">Alocação por Classe</div></div><div class="cw cw-leg-r"><div class="cwc"><canvas id="ch-ipie"></canvas></div><div class="cleg cleg-col" id="ch-ipie-leg"></div></div></div>
      <div class="cc"><div class="cc-h"><div class="cc-t">Alocação Relativa</div><div class="cc-s">% do total</div></div><div class="cw"><canvas id="ch-ibar"></canvas></div></div>
    </div>
    <div class="cc"><div class="cc-h"><div class="cc-t">Composição · Top 20 por PL</div><div class="cc-s">% empilhado</div></div><div class="cw tall"><canvas id="ch-istk"></canvas></div></div>
    <div style="margin:6px 2px -2px;font-size:11px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:var(--g600)">Série Histórica</div>
    <div class="cc"><div class="cc-h"><div class="cc-t">Evolução AuM</div><div class="cc-s">série mensal · R$</div></div><div class="cw tall"><canvas id="ch-hist-pl"></canvas></div></div>
  </div>

  <!-- PARTICIPANTES -->
  <div class="page" id="page-participantes">
    <div class="efpc-bar"><label>Buscar EFPC</label>
      <div class="ac-wrap">
        <input class="inp" id="part-q" placeholder="Busque uma EFPC — clique em ▸ para escolher planos específicos dela..." style="min-width:340px" oninput="partRenderAcList()" onfocus="partRenderAcList()" autocomplete="off">
        <div class="ac-list" id="part-q-ac"></div>
      </div>
      <div class="ms-chips" id="part-chips"></div>
      <span class="freset" id="part-clear" style="display:none" onclick="partClearSel()">limpar seleção</span>
      <div class="fgroup" style="margin-left:auto"><span class="flabel">Modalidade</span>
        <span class="chip on" data-mod="CD" onclick="tgModPart(this)">CD</span>
        <span class="chip on" data-mod="BD" onclick="tgModPart(this)">BD</span>
        <span class="chip on" data-mod="CV" onclick="tgModPart(this)">CV</span>
      </div>
      <a class="ef-link" id="part-ficha-link" style="display:none" onclick="openFicha([...MS.part][0])">🔍 Ver ficha completa</a>
    </div>
    <div class="kpi-grid" id="kpi-part"></div>
    <div class="cg2">
      <div class="cc"><div class="cc-h"><div class="cc-t">Distribuição Etária</div></div><div class="cw tall"><canvas id="ch-pira"></canvas></div></div>
      <div class="cc"><div class="cc-h" id="tipo-h"><div class="cc-t">Ativos × Assistidos</div><div class="cc-s" id="tipo-s">busque uma EFPC para ver por plano</div></div><div class="cw cw-leg-b tall"><div class="cwc"><canvas id="ch-tipo"></canvas></div><div class="cleg" id="ch-tipo-leg"></div></div></div>
    </div>
    <div class="cg2">
      <div class="cc"><div class="cc-h" id="matu-h"><div class="cc-t">Mais Participantes · Top 15</div><div class="cc-s">ativos + assistidos</div></div><div class="cw tall"><canvas id="ch-matu"></canvas></div></div>
      <div class="cc"><div class="cc-h"><div class="cc-t">Evolução de Participantes</div><div class="cc-s">Ativos, Assistidos e Total · anual</div></div><div class="cw tall"><canvas id="ch-evo"></canvas></div></div>
    </div>
  </div>

  <!-- DIRIGENTES -->
  <div class="page" id="page-dirigentes">
    <div class="tc">
      <div class="tc-h">
        <div class="tc-t">Dirigentes</div>
        <div class="tabs">
          <div class="tab active" onclick="setDirSt('todos',this)">Todos</div>
          <div class="tab" onclick="setDirSt('VIGENTE',this)">Vigentes</div>
          <div class="tab" onclick="setDirSt('VENCENDO',this)">Vencendo ≤180d</div>
          <div class="tab" onclick="setDirSt('VENCIDO',this)">Vencidos</div>
          <div class="tab" onclick="setDirSt('SEM_DATA',this)">Sem data</div>
        </div>
        <select id="dir-tipo" onchange="rDirTable()" style="min-width:200px">
          <option value="">Todos os cargos</option>
        </select>
        <input class="inp tc-search" id="dir-q" placeholder="Buscar nome ou EFPC..." oninput="rDirTable()">
      </div>
      <div class="tw"><table><thead><tr>
        <th>Nome</th><th>EFPC</th><th>Cargo</th><th>Pres.</th><th>AETQ</th><th>Rem.</th>
        <th>Início</th><th>Fim Mandato</th><th class="num">Dias Rest.</th><th>Status</th>
      </tr></thead><tbody id="dir-body"></tbody></table></div>
    </div>
  </div>

  <!-- PLANOS -->
  <div class="page" id="page-planos">
    <div class="kpi-grid" id="kpi-planos"></div>
    <div class="cg2">
      <div class="cc"><div class="cc-h"><div class="cc-t">PL por Modalidade</div><div class="cc-s">R$ bilhões</div></div><div class="cw"><canvas id="ch-pmod"></canvas></div></div>
      <div class="cc"><div class="cc-h"><div class="cc-t">Planos por Modalidade</div><div class="cc-s">quantidade</div></div><div class="cw cw-leg-r"><div class="cwc"><canvas id="ch-pmodqt"></canvas></div><div class="cleg cleg-col" id="ch-pmodqt-leg"></div></div></div>
    </div>
    <div class="tc">
      <div class="tc-h">
        <div class="tc-t">Planos de Benefícios <span class="fcount" id="plano-count" style="margin-left:6px">—</span></div>
        <span class="chip on" data-mod="CD" onclick="tgMod(this)">CD</span>
        <span class="chip on" data-mod="BD" onclick="tgMod(this)">BD</span>
        <span class="chip on" data-mod="CV" onclick="tgMod(this)">CV</span>
        <input class="inp tc-search" id="plano-q" placeholder="Buscar plano ou EFPC..." oninput="rPlanos()">
        <span class="chip" onclick="exportTable('planos','csv')">⬇ CSV</span>
        <span class="chip" onclick="exportTable('planos','xlsx')">⬇ XLSX</span>
      </div>
      <div class="tw"><table><thead><tr id="plano-thead-row">
        <th style="cursor:default">#</th>
        <th data-key="NM_PLANO" data-label="Plano" onclick="setPlanoRank('NM_PLANO')">Plano</th>
        <th data-key="CNPJ" data-label="CNPJ" onclick="setPlanoRank('CNPJ')">CNPJ</th>
        <th data-key="SG_EFPC" data-label="EFPC" onclick="setPlanoRank('SG_EFPC')">EFPC</th>
        <th data-key="MODALIDADE" data-label="Modalidade" onclick="setPlanoRank('MODALIDADE')">Modalidade</th>
        <th data-key="TIPO_PATROCINIO" data-label="Patrocínio" onclick="setPlanoRank('TIPO_PATROCINIO')">Patrocínio</th>
        <th data-key="SITUACAO" data-label="Situação" onclick="setPlanoRank('SITUACAO')">Situação</th>
        <th class="num" data-key="PL_valor" data-label="PL" onclick="setPlanoRank('PL_valor')">PL</th>
        <th class="num" data-key="QT_ATIVOS" data-label="Ativos" onclick="setPlanoRank('QT_ATIVOS')">Ativos</th>
        <th class="num" data-key="QT_ASSISTIDOS" data-label="Assistidos" onclick="setPlanoRank('QT_ASSISTIDOS')">Assistidos</th>
        <th class="num" data-key="QT_TOTAL" data-label="Total Part." onclick="setPlanoRank('QT_TOTAL')">Total Part.</th>
      </tr></thead><tbody id="plano-body"></tbody></table></div>
    </div>
  </div>

  <!-- DADOS MANUAIS — desativado a pedido do usuário (não ficou como esperado); mantido comentado para retomar depois
  <div class="page" id="page-manuais">
    <div class="insight">
      <div class="ico">✎</div>
      <div>
        <div class="tt">Edição direta dos arquivos-fonte</div>
        <div class="tx">Clique em "Selecionar arquivo…" e escolha o Excel correspondente (Tiers.xlsx ou Classificação de
        Dados no Balancete.xlsx, dentro da pasta "Dados Manuais"). As alterações são salvas direto no arquivo original —
        não é preciso baixar nem substituir nada manualmente. Funciona apenas no Google Chrome ou Microsoft Edge.
        Depois de salvar, rode o gerador do dashboard novamente para os números refletirem a mudança.</div>
      </div>
    </div>
    <div class="tc">
      <div class="tc-h">
        <div class="tabs" id="man-tabs">
          <div class="tab active" onclick="manSetTab('tiers',this)">Tiers</div>
          <div class="tab" onclick="manSetTab('classificacao',this)">Classificação de Dados no Balancete</div>
        </div>
        <span class="chip" onclick="manAbrir()">📂 Selecionar arquivo…</span>
        <span class="chip" id="man-addrow-btn" style="display:none" onclick="manAddRow()">+ Linha</span>
        <input class="inp tc-search" id="man-q" placeholder="Buscar..." style="display:none" oninput="manRender()">
        <span class="chip" id="man-save-btn" style="display:none;background:var(--teal-500);color:#fff;border-color:var(--teal-500)" onclick="manSalvar()">💾 Salvar no arquivo</span>
        <div id="man-status" style="margin-left:auto;font-size:12px;color:var(--g600);font-weight:600"></div>
      </div>
      <div class="tw" id="man-table-wrap">
        <div style="padding:40px;text-align:center;color:var(--g500);font-size:13px">Nenhum arquivo carregado. Clique em "Selecionar arquivo…" e escolha o Excel correspondente.</div>
      </div>
    </div>
  </div>
  -->

</div></div>

<!-- FICHA EFPC — MODAL -->
<div class="ef-overlay" id="ef-overlay" onclick="if(event.target===this)closeFicha()">
  <div class="ef-modal">
    <div class="ef-head">
      <div>
        <div class="ef-eyebrow" id="ef-eyebrow">—</div>
        <div class="ef-title" id="ef-title">—</div>
        <div class="ef-badges" id="ef-badges"></div>
      </div>
      <div class="ef-close" onclick="closeFicha()">✕</div>
    </div>
    <div class="ef-tabs tabs">
      <div class="tab active" onclick="efTab('desc',this)">Descrição</div>
      <div class="tab" onclick="efTab('inv',this)">Investimentos</div>
      <div class="tab" onclick="efTab('desp',this)">Despesas</div>
      <div class="tab" onclick="efTab('planos',this)">Planos</div>
    </div>
    <div class="ef-body">
      <div class="ef-pane active" id="ef-pane-desc">
        <div class="kpi-grid" id="ef-kpi"></div>
        <div>
          <div class="ef-section-t">Patrocinadoras</div>
          <div id="ef-patro"></div>
        </div>
        <div>
          <div class="ef-section-t">Contato</div>
          <div id="ef-contato" style="font-size:12.5px;color:var(--g700);line-height:1.8"></div>
        </div>
        <div>
          <div class="ef-section-t">Descrição</div>
          <p class="ef-desc-text" id="ef-desc-text"></p>
        </div>
        <div class="ef-fonte" id="ef-fonte"></div>
      </div>
      <div class="ef-pane" id="ef-pane-inv">
        <div class="cg2">
          <div class="cc"><div class="cc-h"><div class="cc-t">Alocação por Classe</div></div><div class="cw cw-leg-r"><div class="cwc"><canvas id="ef-ch-ipie"></canvas></div><div class="cleg cleg-col" id="ef-ch-ipie-leg"></div></div></div>
          <div class="cc"><div class="cc-h"><div class="cc-t">Alocação Relativa</div><div class="cc-s">% do total</div></div><div class="cw"><canvas id="ef-ch-ibar"></canvas></div></div>
        </div>
      </div>
      <div class="ef-pane" id="ef-pane-desp">
        <div class="cc">
          <div class="cc-h"><div class="cc-t">Despesas por Grupo</div></div>
          <div class="tw" style="max-height:240px">
            <table>
              <thead><tr><th>Conta</th><th class="num">Valor</th><th class="num">% Grupo</th><th class="num">% Total</th><th class="num">% AuM</th><th class="num">Mediana Tier</th><th class="num">Vs. Mediana</th></tr></thead>
              <tbody id="ef-desp-tbody"></tbody>
            </table>
          </div>
        </div>
      </div>
      <div class="ef-pane" id="ef-pane-planos">
        <div class="kpi-grid" id="ef-kpi-planos"></div>
        <div class="tc" style="box-shadow:none;border-radius:8px">
          <div class="tw" style="max-height:320px">
            <table>
              <thead><tr>
                <th>Plano</th><th>Modalidade</th><th>Patrocínio</th><th>Situação</th>
                <th class="num">PL</th><th class="num">Ativos</th><th class="num">Assistidos</th><th class="num">Total Part.</th>
              </tr></thead>
              <tbody id="ef-planos-tbody"></tbody>
            </table>
          </div>
        </div>
      </div>
    </div>
  </div>
</div>

<script>
const RAW=__DATA__;
// ── PALETA QUANDO ──
const TC={'CONSULTING':'#1AAAB2','OCIO I':'#0788C9','OCIO II':'#30B783','OCIO III':'#FF9A47','OCIO IV':'#153451','OCIO V':'#B3B3B3','N/D':'#E0E0E0'};
const IC={'Títulos Públicos':'#1AAAB2','Crédito Privado':'#0788C9','Renda Variável':'#FF9A47','Fundos':'#153451','Imóveis':'#30B783','Op. Participantes':'#FF725E'};
const INV_CLASSES=Object.keys(IC); // domínio fixo — mantém as fatias do doughnut estáveis (sem reordenar/sumir) entre filtros
const MC={'CD':'#1AAAB2','BD':'#FF725E','CV':'#0788C9','N/D':'#B3B3B3'};
const CATC=['#0788C9','#1AAAB2','#30B783','#FF9A47','#FF725E','#153451','#F59E0B','#A855F7','#EC4899','#84CC16','#6366F1','#14B8A6'];
function catColor(l){let h=0;for(let i=0;i<l.length;i++)h=(h*31+l.charCodeAt(i))|0;return CATC[Math.abs(h)%CATC.length];}
// ── DESPESAS · tabela accordion (2 grupos + subcontas/residual) ──
let despAbertos=new Set();
let despRowsCache={};
// Mediana, item a item, das EFPCs do mesmo Tier F3 — permite comparar não só a despesa
// total, mas cada grupo/subconta (ex.: Mão-de-Obra Temporária) contra o Tier. Chave
// '<grupo>' = total do grupo; '<grupo>||<item>' = subconta/residual específica. EFPCs do
// tier sem lançamento numa conta entram como 0% (contam na mediana do tier como um todo).
function tierMedianData(tier,field){
  const getT=field==='TIER_QUANDO'?(d=>String(d.TIER_QUANDO||'N/D')):(d=>d['Tier F3']||'N/D');
  const peers=RAW.master.filter(d=>getT(d)===tier&&d.PL_valor>0);
  if(!peers.length)return{};
  const plByMat=new Map(peers.map(d=>[d.NU_MATRICULA_EFPC,d.PL_valor]));
  const porPeer={};
  peers.forEach(p=>{porPeer[p.NU_MATRICULA_EFPC]={};});
  RAW.despesas_estrutura.forEach(d=>{
    const acc=porPeer[d.NU_MATRICULA_EFPC];
    if(!acc)return;
    const key=d.GRUPO+'||'+d.ITEM;
    acc[key]=(acc[key]||0)+d.VALOR;
    acc[d.GRUPO]=(acc[d.GRUPO]||0)+d.VALOR;
  });
  const chaves=new Set();
  Object.values(porPeer).forEach(acc=>Object.keys(acc).forEach(k=>chaves.add(k)));
  const medianas={};
  chaves.forEach(key=>{
    const arr=peers.map(p=>Math.abs(porPeer[p.NU_MATRICULA_EFPC][key]||0)/plByMat.get(p.NU_MATRICULA_EFPC)*100).sort((a,b)=>a-b);
    const n=arr.length;
    medianas[key]=n%2?arr[(n-1)/2]:(arr[n/2-1]+arr[n/2])/2;
  });
  return medianas;
}
function renderDespGrupos(containerId,rows,aum,tierMedian,mats){
  despRowsCache[containerId]={rows,aum,tierMedian,mats};
  // Ícone de evolução (📈) só existe na tabela principal de Despesas — é lá que fica o
  // card com o gráfico de destino (ver showDespEvolucao); a tabela da ficha (popup) não
  // ganhou esse card para não sobrecarregar o modal.
  const podeEvo=containerId==='desp-grupos-tbody';
  const evoIcGrupo=grupo=>podeEvo?`<span class="desp-evo-ic" title="Ver evolução histórica" onclick="event.stopPropagation();showDespEvolucao('${containerId}','${grupo}',null)">📈</span>`:'';
  const evoIcItem=(grupo,item)=>podeEvo?`<span class="desp-evo-ic" title="Ver evolução histórica" onclick="event.stopPropagation();showDespEvolucao('${containerId}','${grupo}','${item.replace(/'/g,"\\'")}')">📈</span>`:'';
  const totalGeral=rows.reduce((s,d)=>s+d.VALOR,0);
  const vsMedTd=(pct,key)=>{
    const med=tierMedian?tierMedian[key]:null;
    if(med==null)return'<td class="num">—</td><td class="num">—</td>';
    const diff=pct-med;
    const cls=diff>0.005?'neg':diff<-0.005?'pos':'neu';
    return`<td class="num">${med.toFixed(3)}%</td><td class="num ${cls}">${diff>0?'+':''}${diff.toFixed(3)}pp</td>`;
  };
  let html='';
  RAW.despesas_grupos_cfg.forEach(cfg=>{
    const grpRows=rows.filter(d=>d.GRUPO===cfg.grupo);
    const agg={};
    grpRows.forEach(d=>{agg[d.ITEM]=(agg[d.ITEM]||0)+d.VALOR;});
    const nomes=cfg.itens.map(it=>it.nome).concat([cfg.residual]);
    const totalGrupo=nomes.reduce((s,n)=>s+(agg[n]||0),0);
    const key=containerId+'::'+cfg.grupo;
    const aberto=despAbertos.has(key);
    const pctAumGrupo=aum>0?Math.abs(totalGrupo)/aum*100:0;
    html+=`<tr class="desp-grupo-row" onclick="despToggle('${key}')">
      <td><span class="desp-caret">${aberto?'▾':'▸'}</span> ${cfg.grupo}${evoIcGrupo(cfg.grupo)}</td>
      <td class="num">R$ ${fmt.brl(Math.abs(totalGrupo))}</td>
      <td class="num">—</td>
      <td class="num">${totalGeral!==0?(totalGrupo/totalGeral*100).toFixed(1):'0.0'}%</td>
      <td class="num">${pctAumGrupo.toFixed(3)}%</td>
      ${vsMedTd(pctAumGrupo,cfg.grupo)}
    </tr>`;
    if(aberto){
      nomes.forEach((nome,i)=>{
        const v=agg[nome]||0;
        const residual=i===nomes.length-1;
        const pctAumItem=aum>0?Math.abs(v)/aum*100:0;
        html+=`<tr class="desp-sub-row${residual?' desp-residual':''}">
          <td style="padding-left:38px">${nome}${evoIcItem(cfg.grupo,nome)}</td>
          <td class="num">R$ ${fmt.brl(Math.abs(v))}</td>
          <td class="num">${totalGrupo!==0?(v/totalGrupo*100).toFixed(1):'0.0'}%</td>
          <td class="num">${totalGeral!==0?(v/totalGeral*100).toFixed(1):'0.0'}%</td>
          <td class="num">${pctAumItem.toFixed(3)}%</td>
          ${vsMedTd(pctAumItem,cfg.grupo+'||'+nome)}
        </tr>`;
      });
    }
  });
  document.getElementById(containerId).innerHTML=html;
}
function despToggle(key){
  if(despAbertos.has(key))despAbertos.delete(key);else despAbertos.add(key);
  const containerId=key.split('::')[0];
  const cache=despRowsCache[containerId]||{rows:[],aum:0,tierMedian:null,mats:new Set()};
  renderDespGrupos(containerId,cache.rows,cache.aum,cache.tierMedian,cache.mats);
}
// Evolução histórica de uma conta/grupo de despesa — usa a mesma seleção de EFPCs
// (mats) da tabela que originou o clique. Cada competência em RAW.hist_despesas já vem
// anualizada com o fator próprio dela (ver carregar_historico_despesas no gerador
// Python), então comparar entre competências compara taxas equivalentes a 12 meses.
function showDespEvolucao(containerId,grupo,item){
  const cache=despRowsCache[containerId];
  if(!cache||!cache.mats)return;
  const mats=cache.mats;
  const comps=RAW.hist_despesas_competencias||[];
  const card=document.getElementById('desp-evo-card'),empty=document.getElementById('desp-evo-empty');
  card.style.display='flex';
  document.getElementById('desp-evo-t').textContent=`Evolução · ${grupo}${item?' · '+item:''}`;
  if(!comps.length){
    empty.style.display='block';
    document.getElementById('ch-desp-evo').style.display='none';
  }else{
    empty.style.display='none';
    document.getElementById('ch-desp-evo').style.display='block';
    const rowsHist=RAW.hist_despesas.filter(d=>mats.has(String(d.NU_MATRICULA_EFPC))&&d.GRUPO===grupo&&(item?d.ITEM===item:true));
    const serie=comps.map(c=>rowsHist.filter(d=>d.COMPETENCIA===c).reduce((s,d)=>s+Math.abs(d.VALOR),0));
    mk('ch-desp-evo','bar',{labels:comps.map(compLabel),datasets:[{data:serie,backgroundColor:'#1AAAB2',borderRadius:5}]},{plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>' R$ '+fmt.brl(c.raw)}}},scales:{y:{ticks:{callback:v=>fmt.brl(v)}}}});
  }
  card.scrollIntoView({behavior:'smooth',block:'nearest'});
}
const fmt={brl:v=>{const a=Math.abs(v);if(a>=1e12)return(v/1e12).toFixed(2)+'T';if(a>=1e9)return(v/1e9).toFixed(2)+'B';if(a>=1e6)return(v/1e6).toFixed(1)+'M';if(a>=1e3)return(v/1e3).toFixed(0)+'K';return v.toFixed(0)},num:v=>v.toLocaleString('pt-BR'),dt:s=>s?s.split('-').reverse().join('/'):'—'};
// ── FILTROS GLOBAIS ──
const F={f3:new Set(['CONSULTING','OCIO I','OCIO II','OCIO III','OCIO IV','OCIO V','N/D']),
  tq:new Set(['1','2','3','N/D']),bpo:new Set(['S','N'])};
function tg(el){
  const f=el.dataset.f,v=el.dataset.v;
  if(F[f].has(v)){F[f].delete(v);el.classList.remove('on');el.classList.add('off');}
  else{F[f].add(v);el.classList.add('on');el.classList.remove('off');}
  refresh();
}
function resetF(){
  F.f3=new Set(['CONSULTING','OCIO I','OCIO II','OCIO III','OCIO IV','OCIO V','N/D']);
  F.tq=new Set(['1','2','3','N/D']);F.bpo=new Set(['S','N']);
  document.querySelectorAll('.chip').forEach(c=>{c.classList.add('on');c.classList.remove('off');});
  refresh();
}
function mf(){return RAW.master.filter(d=>
  F.f3.has(d['Tier F3']||'N/D')&&F.tq.has(String(d.TIER_QUANDO||'N/D'))&&
  F.bpo.has(d['BPO(S/N)']||'N'));}
function mfMats(){return new Set(mf().map(d=>String(d.NU_MATRICULA_EFPC)));}
function refresh(){
  const m=mf();
  document.getElementById('fcount').textContent=m.length+' de '+RAW.master.length+' EFPCs';
  document.getElementById('side-count').textContent=m.length+' de '+RAW.master.length+' EFPCs no filtro';
  const ap=document.querySelector('.page.active');
  if(ap)renderPage(ap.id.replace('page-',''));
}
// ── NAV ──
const PT={overview:['Visão Geral','Painel Geral'],ranking:['Visão Geral','Ranking EFPCs'],
  despesas:['Análise Financeira','Despesas'],investimentos:['Análise Financeira','Investimentos & Série Histórica'],
  participantes:['Pessoas','Participantes'],dirigentes:['Pessoas','Dirigentes & Governança'],
  planos:['Planos','Planos por EFPC']/*,manuais:['Dados','Editar Dados Manuais']*/};
// Data de referência mostrada no topo de cada aba: a maioria das abas deriva do
// balancete consolidado mais recente; só Investimentos usa a série de Fundos
// Exclusivos, que é uma fonte separada e pode estar numa competência diferente.
function compTxt(mes,ano){return(mes&&ano)?String(mes).padStart(2,'0')+'/'+ano:'—';}
function pgUpdatedTxt(p){
  const M=RAW.meta_atualizacao||{};
  return p==='investimentos'
    ?'Fundos Exclusivos atualizados até '+compTxt(M.investimentos_mes,M.investimentos_ano)
    :'Balancete atualizado até '+compTxt(M.balancete_mes,M.balancete_ano);
}
function navTo(p,el){
  document.querySelectorAll('.page').forEach(x=>x.classList.remove('active'));
  document.querySelectorAll('.nav-it').forEach(x=>x.classList.remove('active'));
  document.getElementById('page-'+p).classList.add('active');
  if(el)el.classList.add('active');
  document.getElementById('pg-eyebrow').textContent=PT[p][0];
  document.getElementById('pg-title').textContent=PT[p][1];
  document.getElementById('pg-updated').textContent=pgUpdatedTxt(p);
  renderPage(p);
}
function renderPage(p){
  ({overview:rOverview,ranking:rRanking,despesas:rDespesas,investimentos:rInvest,
    participantes:rPart,dirigentes:rDirTable,planos:rPlanos/*,manuais:rManuais*/})[p]();
}
// ── CHART BASE ──
const charts={};
function dk(id){if(charts[id]){charts[id].destroy();delete charts[id];}}
const CD={responsive:true,maintainAspectRatio:true,animation:{duration:450,easing:'easeOutQuart'},
  plugins:{legend:{labels:{color:'#5F5F5F',font:{family:'Inter',size:12,weight:'500'},boxWidth:10,usePointStyle:true,pointStyle:'circle'}},
  tooltip:{backgroundColor:'#153451',borderColor:'#153451',titleColor:'#fff',bodyColor:'rgba(255,255,255,.85)',
    titleFont:{family:'Inter',size:13,weight:'600'},bodyFont:{family:'Inter',size:12.5},padding:10,cornerRadius:8}},
  scales:{x:{ticks:{color:'#969696',font:{family:'Inter',size:11.5}},grid:{color:'rgba(21,52,81,.05)'}},
    y:{ticks:{color:'#969696',font:{family:'Inter',size:11.5}},grid:{color:'rgba(21,52,81,.05)'}}}};
function mk(id,type,data,opts,legendId){
  const c=document.getElementById(id);if(!c)return;
  opts=opts||{};
  const persistOpt=!!opts.persist;
  // legendId presente: usamos a legenda HTML custom (ver renderLegend) em vez da legenda
  // nativa do Chart.js — desliga o desenho nativo no canvas. maintainAspectRatio:false
  // porque esses canvases agora vivem dentro de uma caixa .cwc com altura própria
  // (ver CSS); deixar o Chart.js calcular a própria proporção nesse layout flex é o que
  // fazia os gráficos de rosca renderizarem mal dimensionados.
  if(legendId)opts=dm({maintainAspectRatio:false,plugins:{legend:{display:false}}},opts);
  const options=dm(JSON.parse(JSON.stringify(CD)),opts);
  // Doughnut/pie: mantém a instância e atualiza os dados no lugar, para que o Chart.js
  // anime a transição (fatias encolhendo/"apagando") em vez do corte abrupto de
  // destruir e recriar o gráfico a cada filtro. Categorias com domínio estável (ver
  // INV_CLASSES / dim.order / mods fixos) garantem que o índice de cada fatia não mude.
  // Gráficos de barra com domínio de categorias que pode encolher (ex.: ch-pmod) também
  // podem optar por esse comportamento passando {persist:true}.
  if((type==='doughnut'||type==='pie'||persistOpt)&&charts[id]&&charts[id].config.type===type){
    charts[id].data=data;
    charts[id].options=options;
    charts[id].update();
  }else{
    dk(id);
    charts[id]=new Chart(c,{type,data,options});
  }
  if(legendId)renderLegend(id,legendId);
}
// ── LEGENDA CUSTOM ──
// Substitui a legenda nativa do Chart.js (que risca o texto do item ao ocultá-lo) por uma
// legenda HTML: ao clicar, o item inteiro (texto + marcador de cor) só "apaga" via
// opacidade — ver .cleg-it.off no CSS.
function renderLegend(chartId,legendId){
  const chart=charts[chartId],el=document.getElementById(legendId);
  if(!chart||!el)return;
  const items=chart.options.plugins.legend.labels.generateLabels(chart);
  el.innerHTML=items.map(it=>{
    const idx=it.index!==undefined?it.index:it.datasetIndex;
    return`<span class="cleg-it${it.hidden?' off':''}" onclick="legendToggle('${chartId}','${legendId}',${idx})">
      <span class="cleg-dot" style="background:${it.fillStyle}"></span>${it.text}</span>`;
  }).join('');
}
function legendToggle(chartId,legendId,idx){
  const chart=charts[chartId];if(!chart)return;
  if(chart.config.type==='doughnut'||chart.config.type==='pie'||chart.config.type==='polarArea'){
    chart.toggleDataVisibility(idx);
  }else{
    const meta=chart.getDatasetMeta(idx);
    meta.hidden=meta.hidden===null?!chart.data.datasets[idx].hidden:null;
  }
  chart.update();
  renderLegend(chartId,legendId);
}
function dm(a,b){for(const k in b){if(b[k]&&typeof b[k]==='object'&&!Array.isArray(b[k])){a[k]=a[k]||{};dm(a[k],b[k]);}else a[k]=b[k];}return a;}
function ns(){return{scales:{x:{display:false},y:{display:false}}};}
// ── SELECTS ──
function popSel(){
  const tipos=[...new Set(RAW.dirigentes.map(d=>d.tp))].sort();
  const dt=document.getElementById('dir-tipo');
  tipos.forEach(t=>{const o=document.createElement('option');o.value=t;o.textContent=t;dt.appendChild(o);});
}
// ── BUSCA EFPC · MULTI-SELEÇÃO (Despesas / Investimentos [aloc. atual + série histórica] / Participantes) ──
// Cada página mantém um Set de matrículas marcadas via checkbox. Vazio = comportamento
// antigo de "Todas" (respeitando os filtros gerais). Digitar no campo só filtra a lista
// visível — não altera a seleção — e "limpar seleção" desmarca tudo de uma vez, em vez de
// precisar remover EFPC por EFPC.
const MS={desp:new Set(),inv:new Set(),part:new Set()};
const MS_FN={desp:'rDespesas',inv:'rInvest',part:'rPart'};
function msRenderList(page){
  const q=(document.getElementById(page+'-q').value||'').toLowerCase().trim();
  const ac=document.getElementById(page+'-q-ac');
  const all=RAW.master.slice().sort((a,b)=>a.SG_EFPC.localeCompare(b.SG_EFPC));
  const matches=q?all.filter(d=>d.SG_EFPC.toLowerCase().includes(q)||(d.NM_RAZAO_SOCIAL||'').toLowerCase().includes(q)):all;
  if(!matches.length){ac.innerHTML='<div class="ac-empty">Nenhuma EFPC encontrada</div>';return;}
  ac.innerHTML=matches.slice(0,300).map(d=>{
    const checked=MS[page].has(d.NU_MATRICULA_EFPC);
    return`<label class="ac-item${checked?' checked':''}"><input type="checkbox" ${checked?'checked':''} onclick="msToggle('${page}',${d.NU_MATRICULA_EFPC},this.checked)"> ${d.SG_EFPC}</label>`;
  }).join('');
}
function msRender(page){
  msRenderList(page);
  document.getElementById(page+'-q-ac').classList.add('open');
}
function msToggle(page,mat,checked){
  if(checked)MS[page].add(mat);else MS[page].delete(mat);
  if(page==='part'){partRenderAcList();partRenderChips();}
  else{msRenderList(page);msRenderChips(page);}
  window[MS_FN[page]]();
}
function msRenderChips(page){
  const chips=document.getElementById(page+'-chips');
  const clr=document.getElementById(page+'-clear');
  const sel=[...MS[page]];
  if(!sel.length){chips.innerHTML='';clr.style.display='none';return;}
  clr.style.display='inline';
  chips.innerHTML=sel.map(mat=>{
    const d=RAW.master.find(x=>x.NU_MATRICULA_EFPC===mat);
    return`<span class="ms-chip">${d?d.SG_EFPC:mat}<span class="ms-x" onclick="msToggle('${page}',${mat},false)">✕</span></span>`;
  }).join('');
}
function msClear(page){
  MS[page].clear();
  document.getElementById(page+'-q').value='';
  msRenderList(page);
  msRenderChips(page);
  window[MS_FN[page]]();
}
// ── BUSCA EFPC COM DRILL-DOWN DE PLANOS (só a página Participantes) ──
// Um único campo de busca: cada EFPC do dropdown tem uma seta (▸/▾) que expande, ali
// mesmo, a lista de planos daquela EFPC — sem precisar de um segundo campo de busca.
// Marcar a EFPC filtra pelo total dela; expandir e marcar plano(s) específicos filtra só
// por eles (MSPL), independente da EFPC estar marcada ou não.
const MSPL={part:new Set()};
const MSPL_FN={part:'rPart'};
const PLANO_BY_CNPB=new Map(RAW.planos.map(p=>[p.NU_CNPB,p]));
const EFPC_COM_PLANOS=new Set(RAW.planos.map(p=>p.NU_MATRICULA_EFPC));
let partExpand=new Set();
function mspToggleExpand(mat){
  if(partExpand.has(mat))partExpand.delete(mat);else partExpand.add(mat);
  partRenderAcList();
}
function partRenderAcList(){
  const q=(document.getElementById('part-q').value||'').toLowerCase().trim();
  const ac=document.getElementById('part-q-ac');
  const all=RAW.master.slice().sort((a,b)=>a.SG_EFPC.localeCompare(b.SG_EFPC));
  const matches=q?all.filter(d=>d.SG_EFPC.toLowerCase().includes(q)||(d.NM_RAZAO_SOCIAL||'').toLowerCase().includes(q)):all;
  ac.classList.add('open');
  if(!matches.length){ac.innerHTML='<div class="ac-empty">Nenhuma EFPC encontrada</div>';return;}
  const planCountByMat={};
  MSPL.part.forEach(cnpb=>{const p=PLANO_BY_CNPB.get(cnpb);if(p)planCountByMat[p.NU_MATRICULA_EFPC]=(planCountByMat[p.NU_MATRICULA_EFPC]||0)+1;});
  ac.innerHTML=matches.slice(0,300).map(d=>{
    const checked=MS.part.has(d.NU_MATRICULA_EFPC);
    const temPlanos=EFPC_COM_PLANOS.has(d.NU_MATRICULA_EFPC);
    const expanded=partExpand.has(d.NU_MATRICULA_EFPC);
    const nPl=planCountByMat[d.NU_MATRICULA_EFPC]||0;
    let row=`<label class="ac-item${checked?' checked':''}">`+
      `<span class="ac-caret${temPlanos?'':' ac-caret-empty'}" ${temPlanos?`onclick="event.preventDefault();event.stopPropagation();mspToggleExpand(${d.NU_MATRICULA_EFPC})"`:''}>${temPlanos?(expanded?'▾':'▸'):''}</span>`+
      `<input type="checkbox" ${checked?'checked':''} onclick="msToggle('part',${d.NU_MATRICULA_EFPC},this.checked)"> ${d.SG_EFPC}`+
      `${nPl?`<span style="margin-left:6px;color:var(--teal-600);font-size:10px;font-weight:600">${nPl} plano${nPl>1?'s':''}</span>`:''}`+
      `</label>`;
    if(expanded){
      const planos=RAW.planos.filter(p=>p.NU_MATRICULA_EFPC===d.NU_MATRICULA_EFPC);
      row+=planos.map(p=>{
        const pchecked=MSPL.part.has(p.NU_CNPB);
        return`<label class="ac-item ac-sub${pchecked?' checked':''}"><input type="checkbox" ${pchecked?'checked':''} onclick="msplToggle('part',${p.NU_CNPB},this.checked)"> ${p.NM_PLANO||p.SG_PLANO}</label>`;
      }).join('');
    }
    return row;
  }).join('');
}
function msplToggle(page,cnpb,checked){
  if(checked)MSPL[page].add(cnpb);else MSPL[page].delete(cnpb);
  partRenderAcList();
  partRenderChips();
  window[MSPL_FN[page]]();
}
function partRenderChips(){
  const chips=document.getElementById('part-chips');
  const clr=document.getElementById('part-clear');
  const efs=[...MS.part],pls=[...MSPL.part];
  if(!efs.length&&!pls.length){chips.innerHTML='';clr.style.display='none';return;}
  clr.style.display='inline';
  const efChips=efs.map(mat=>{
    const d=RAW.master.find(x=>x.NU_MATRICULA_EFPC===mat);
    return`<span class="ms-chip">${d?d.SG_EFPC:mat}<span class="ms-x" onclick="msToggle('part',${mat},false)">✕</span></span>`;
  });
  const plChips=pls.map(cnpb=>{
    const p=PLANO_BY_CNPB.get(cnpb);
    return`<span class="ms-chip" style="background:var(--g150);border-color:var(--g300);color:var(--g700)">${p?(p.SG_EFPC+' · '+(p.SG_PLANO||p.NM_PLANO)):cnpb}<span class="ms-x" onclick="msplToggle('part',${cnpb},false)">✕</span></span>`;
  });
  chips.innerHTML=efChips.concat(plChips).join('');
}
function partClearSel(){
  MS.part.clear();MSPL.part.clear();partExpand.clear();
  document.getElementById('part-q').value='';
  partRenderAcList();
  partRenderChips();
  rPart();
}
function tierTag(t){const c=TC[t]||'#B3B3B3';return`<span class="tag" style="background:${c}1A;color:${c==='#B3B3B3'||c==='#E0E0E0'?'#5F5F5F':c};border:1px solid ${c}55">${t||'N/D'}</span>`;}
function contatoHtml(d){
  const parts=[];
  if(d.SITE)parts.push(`<a href="${/^https?:\/\//i.test(d.SITE)?d.SITE:'http://'+d.SITE}" target="_blank" rel="noopener" title="${d.SITE}" style="color:var(--teal-600);text-decoration:none">🌐</a>`);
  if(d.EMAIL)parts.push(`<a href="mailto:${d.EMAIL}" title="${d.EMAIL}" style="color:var(--teal-600);text-decoration:none">✉</a>`);
  if(d.FONE)parts.push(`<span title="${d.FONE}" style="color:var(--g600)">☎</span>`);
  return parts.length?parts.join(' '):'<span style="color:var(--g400)">—</span>';
}
function contatoFullHtml(d){
  const parts=[];
  if(d.SITE)parts.push(`🌐 <a href="${/^https?:\/\//i.test(d.SITE)?d.SITE:'http://'+d.SITE}" target="_blank" rel="noopener" style="color:var(--teal-600)">${d.SITE}</a>`);
  if(d.EMAIL)parts.push(`✉ <a href="mailto:${d.EMAIL}" style="color:var(--teal-600)">${d.EMAIL}</a>`);
  if(d.FONE)parts.push(`☎ ${d.FONE}`);
  return parts.length?parts.join('<br>'):'<span style="color:var(--g500);font-style:italic">Sem dados de contato cadastrados</span>';
}
// ── FICHA EFPC (MODAL) ──
let efMat=null;
function openFicha(mat){
  mat=+mat;if(!mat)return;
  const ef=RAW.master.find(d=>d.NU_MATRICULA_EFPC===mat);
  if(!ef)return;
  efMat=mat;
  const fc=RAW.fichas[mat];
  document.getElementById('ef-eyebrow').textContent=`Tier ${ef['Tier F3']||'N/D'} · Marina ${ef.TIER_QUANDO==='N/D'?'—':ef.TIER_QUANDO}`;
  document.getElementById('ef-title').textContent=`${ef.SG_EFPC} — ${ef.NM_RAZAO_SOCIAL||''}`;
  document.getElementById('ef-badges').innerHTML=`
  <span class="tag" style="background:${ef['BPO(S/N)']==='S'?'var(--success-100)':'var(--g200)'};color:${ef['BPO(S/N)']==='S'?'var(--success)':'var(--g600)'}">${ef['BPO(S/N)']==='S'?'Com BPO':'Sem BPO'}</span>
  <span class="tag" style="background:${ef.FIT_MARINA==='S'?'var(--teal-50)':'var(--g200)'};color:${ef.FIT_MARINA==='S'?'var(--teal-700)':'var(--g600)'}">${ef.FIT_MARINA==='S'?'★ Fit Marina':'Sem Fit Marina'}</span>
  ${ef.Prestador?`<span class="tag" style="background:var(--g200);color:var(--g700)">${ef.Prestador}${ef.Sistema?' · '+ef.Sistema:''}</span>`:''}`;
  document.getElementById('ef-kpi').innerHTML=`
  <div class="kpi teal"><div class="kpi-l">Ativos Sob Gestão</div><div class="kpi-v">R$ ${fmt.brl(ef.PL_valor)}</div></div>
  <div class="kpi ${ef.DESP_PCT_PL>15?'red':ef.DESP_PCT_PL<5?'green':'orange'}"><div class="kpi-l">Desp%PL</div><div class="kpi-v">${ef.DESP_PCT_PL.toFixed(2)}%</div></div>
  <div class="kpi navy"><div class="kpi-l">Ativos</div><div class="kpi-v">${fmt.num(ef.QT_ATIVOS)}</div></div>
  <div class="kpi green"><div class="kpi-l">Assistidos</div><div class="kpi-v">${fmt.num(ef.QT_ASSISTIDOS)}</div></div>
  <div class="kpi blue"><div class="kpi-l">Planos</div><div class="kpi-v">${fmt.num(ef.QT_PLANOS||0)}</div><div class="kpi-s">${ef.TIPOS_PLANOS||'—'}</div></div>`;
  document.getElementById('ef-contato').innerHTML=contatoFullHtml(ef);
  if(fc){
    document.getElementById('ef-patro').innerHTML=(fc.patrocinadoras||[]).map(p=>`<span class="tag" style="background:var(--teal-50);color:var(--teal-700);border:1px solid var(--teal-100);margin:2px 4px 2px 0">${p}</span>`).join('');
    document.getElementById('ef-desc-text').textContent=fc.descricao||'';
    document.getElementById('ef-fonte').textContent=fc.fonte?`Fonte: ${fc.fonte} · base atualizada em ${fmt.dt(RAW.fichas_data)}`:'';
  }else{
    document.getElementById('ef-patro').innerHTML='<span style="color:var(--g500);font-style:italic">Ficha ainda não pesquisada para esta EFPC</span>';
    document.getElementById('ef-desc-text').textContent='';
    document.getElementById('ef-fonte').textContent='';
  }
  efTab('desc',document.querySelector('.ef-tabs .tab'));
  document.getElementById('ef-overlay').classList.add('open');
}
function closeFicha(){document.getElementById('ef-overlay').classList.remove('open');efMat=null;}
function efTab(name,el){
  document.querySelectorAll('.ef-tabs .tab').forEach(t=>t.classList.remove('active'));
  if(el)el.classList.add('active');
  document.querySelectorAll('.ef-pane').forEach(p=>p.classList.remove('active'));
  document.getElementById('ef-pane-'+name).classList.add('active');
  if(!efMat)return;
  if(name==='inv')efRenderInv(efMat);
  if(name==='desp')efRenderDesp(efMat,RAW.master.find(d=>d.NU_MATRICULA_EFPC===efMat));
  if(name==='planos')efRenderPlanos(efMat);
}
function efRenderInv(mat){
  const byClass={};INV_CLASSES.forEach(c=>byClass[c]=0);
  RAW.investimentos.filter(d=>d.NU_MATRICULA_EFPC===mat).forEach(d=>{byClass[d.CLASSE]=(byClass[d.CLASSE]||0)+d.VL_SALDO_FINAL;});
  const cl=INV_CLASSES,dataVals=cl.map(c=>byClass[c]||0),tot=dataVals.reduce((s,v)=>s+v,0);
  mk('ef-ch-ipie','doughnut',{labels:cl,datasets:[{data:dataVals.map(v=>v/1e9),backgroundColor:cl.map(c=>IC[c]),borderColor:'#fff',borderWidth:2}]},{...ns(),cutout:'62%',plugins:{tooltip:{callbacks:{label:c=>` ${c.label}: R$ ${c.raw.toFixed(1)}B`}}}},'ef-ch-ipie-leg');
  mk('ef-ch-ibar','bar',{labels:cl,datasets:[{data:dataVals.map(v=>tot>0?v/tot*100:0),backgroundColor:cl.map(c=>IC[c]),borderRadius:5}]},{indexAxis:'y',plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>` ${c.raw.toFixed(1)}%`}}},scales:{x:{ticks:{callback:v=>v+'%'}}}});
}
function efRenderDesp(mat,ef){
  renderDespGrupos('ef-desp-tbody',RAW.despesas_estrutura.filter(d=>d.NU_MATRICULA_EFPC===mat),ef?ef.PL_valor:0,ef?tierMedianData(ef['Tier F3']||'N/D'):null,new Set([String(mat)]));
}
function efRenderPlanos(mat){
  const d=RAW.planos.filter(x=>x.NU_MATRICULA_EFPC===mat).slice().sort((a,b)=>b.PL_valor-a.PL_valor);
  const totPL=d.reduce((s,x)=>s+x.PL_valor,0),totAt=d.reduce((s,x)=>s+x.QT_ATIVOS,0),totAs=d.reduce((s,x)=>s+x.QT_ASSISTIDOS,0);
  document.getElementById('ef-kpi-planos').innerHTML=`
  <div class="kpi teal"><div class="kpi-l">Planos</div><div class="kpi-v">${fmt.num(d.length)}</div></div>
  <div class="kpi blue"><div class="kpi-l">Patrimônio dos Planos</div><div class="kpi-v">R$ ${fmt.brl(totPL)}</div></div>
  <div class="kpi green"><div class="kpi-l">Participantes Ativos</div><div class="kpi-v">${fmt.num(totAt)}</div></div>
  <div class="kpi orange"><div class="kpi-l">Assistidos</div><div class="kpi-v">${fmt.num(totAs)}</div></div>`;
  document.getElementById('ef-planos-tbody').innerHTML=d.length?d.map(x=>`<tr>
  <td style="max-width:220px;overflow:hidden;text-overflow:ellipsis;font-weight:600">${x.NM_PLANO||x.SG_PLANO||'—'}</td>
  <td>${modTag(x.MODALIDADE)}</td>
  <td>${patrocinioTag(x.TIPO_PATROCINIO)}</td>
  <td style="color:var(--g600);font-size:11px">${x.SITUACAO||'—'}</td>
  <td class="num">R$ ${fmt.brl(x.PL_valor)}</td>
  <td class="num">${fmt.num(x.QT_ATIVOS)}</td>
  <td class="num">${fmt.num(x.QT_ASSISTIDOS)}</td>
  <td class="num">${fmt.num(x.QT_TOTAL)}</td>
  </tr>`).join(''):'<tr><td colspan="8" style="text-align:center;color:var(--g500);font-style:italic;padding:24px">Nenhum plano cadastrado para esta EFPC</td></tr>';
}
// ── OVERVIEW ──
// Domínio fixo de prestadores (todos os nomes já vistos no cadastro, não só os do filtro
// atual) — mantém as fatias do doughnut estáveis entre filtros em vez de somem/aparecem.
const PRESTADOR_ORDER=(()=>{
  const tot={};
  RAW.master.forEach(d=>{const l=d.Prestador||'Sem Prestador';tot[l]=(tot[l]||0)+d.PL_valor;});
  return Object.keys(tot).sort((a,b)=>{
    if(a==='Sem Prestador')return 1;
    if(b==='Sem Prestador')return -1;
    return tot[b]-tot[a];
  });
})();
const DISP_DIMS={
  tierf3:{name:'Tier F3',get:d=>d['Tier F3']||'N/D',color:l=>TC[l]||'#B3B3B3',order:['CONSULTING','OCIO I','OCIO II','OCIO III','OCIO IV','OCIO V','N/D']},
  tierq:{name:'Marina',get:d=>String(d.TIER_QUANDO||'N/D'),color:l=>({'1':'#1AAAB2','2':'#0788C9','3':'#FF9A47','N/D':'#B3B3B3'}[l]||'#B3B3B3'),order:['1','2','3','N/D']},
  bpo:{name:'BPO',get:d=>d['BPO(S/N)']==='S'?'Com BPO':'Sem BPO',color:l=>l==='Com BPO'?'#22825D':'#B3B3B3',order:['Com BPO','Sem BPO']},
  prestador:{name:'Prestador',get:d=>d.Prestador||'Sem Prestador',color:l=>l==='Sem Prestador'?'#B3B3B3':catColor(l),order:PRESTADOR_ORDER},
};
let dispDim='tierf3';
function setDispDim(k,el){dispDim=k;document.querySelectorAll('#disp-tabs .tab').forEach(t=>t.classList.remove('active'));if(el)el.classList.add('active');renderDisp();}
function renderDisp(){
  const m=mf();
  const dim=DISP_DIMS[dispDim];
  // domínio de categorias sempre fixo (dim.order) — categorias sem membros no filtro atual
  // aparecem com valor 0 e a fatia "apaga" suavemente, em vez de sumir do gráfico
  const labels=dim.order;
  document.getElementById('disp-t1').textContent=`Distribuição por ${dim.name}`;
  document.getElementById('disp-t2').textContent=`Despesa Total por ${dim.name}`;
  mk('ch-tier-pl','doughnut',{labels,datasets:[{data:labels.map(l=>m.filter(d=>dim.get(d)===l).reduce((s,d)=>s+d.PL_valor,0)/1e9),backgroundColor:labels.map(l=>dim.color(l)),borderColor:'#fff',borderWidth:2}]},{...ns(),cutout:'62%',plugins:{tooltip:{callbacks:{label:c=>` ${c.label}: R$ ${c.raw.toFixed(1)}B`}}}},'ch-tier-pl-leg');
  mk('ch-tier-desp','bar',{labels,datasets:[{data:labels.map(l=>m.filter(d=>dim.get(d)===l).reduce((s,d)=>s+Math.abs(d.DESP_TOTAL),0)/1e6),backgroundColor:labels.map(l=>dim.color(l)),borderRadius:6}]},{plugins:{legend:{display:false}},scales:{y:{ticks:{callback:v=>fmt.brl(v*1e6)}}}});
}
function rOverview(){
  const m=mf();
  const tPL=m.reduce((s,d)=>s+d.PL_valor,0),tD=m.reduce((s,d)=>s+d.DESP_TOTAL,0);
  const tA=m.reduce((s,d)=>s+d.QT_ATIVOS,0),tAs=m.reduce((s,d)=>s+d.QT_ASSISTIDOS,0);
  const fit=m.filter(d=>d.FIT_MARINA==='S').length;
  document.getElementById('kpi-overview').innerHTML=`
  <div class="kpi navy"><div class="kpi-l">EFPCs no Filtro</div><div class="kpi-v">${m.length}</div><div class="kpi-s">de ${RAW.master.length} no total</div></div>
  <div class="kpi teal"><div class="kpi-l">Ativos Sob Gestão</div><div class="kpi-v">R$ ${fmt.brl(tPL)}</div></div>
  <div class="kpi red"><div class="kpi-l">Despesas Totais</div><div class="kpi-v">R$ ${fmt.brl(Math.abs(tD))}</div></div>
  <div class="kpi green"><div class="kpi-l">Participantes Ativos</div><div class="kpi-v">${fmt.num(tA)}</div><div class="kpi-s">${fmt.num(tAs)} assistidos</div></div>
  <div class="kpi orange"><div class="kpi-l">Fit Marina</div><div class="kpi-v">${fit}</div><div class="kpi-s">${m.filter(d=>d['BPO(S/N)']==='S').length} com BPO</div></div>`;
  document.getElementById('fcount').textContent=m.length+' de '+RAW.master.length+' EFPCs';
  document.getElementById('side-count').textContent=m.length+' de '+RAW.master.length+' EFPCs no filtro';
  renderDisp();
  const t15=m.slice().sort((a,b)=>b.PL_valor-a.PL_valor).slice(0,15);
  mk('ch-top-pl','bar',{labels:t15.map(d=>d.SG_EFPC),datasets:[{data:t15.map(d=>d.PL_valor/1e9),backgroundColor:t15.map(d=>TC[d['Tier F3']||'N/D']),borderRadius:5}]},{indexAxis:'y',plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>` R$ ${c.raw.toFixed(2)}B`}}},scales:{x:{ticks:{callback:v=>v+'B'}}}});
  const pcts=m.filter(d=>d.PL_valor>0&&Math.abs(d.DESP_PCT_PL)<50).map(d=>d.DESP_PCT_PL);
  const bins=Array.from({length:6},(_,i)=>-12+i*2);
  mk('ch-hist','bar',{labels:bins.map(b=>b+'%'),datasets:[{data:bins.map(b=>pcts.filter(p=>p>=b&&p<b+2).length),backgroundColor:'#FF725E',borderRadius:4}]},{plugins:{legend:{display:false}}});
}
// ── RANKING ──
// Campos de texto (ordenação alfabética); qualquer outro campo é tratado como numérico
// (ordenação por magnitude, como já era o comportamento das abas de ranking).
const RANK_STR_FIELDS=new Set(['SG_EFPC','NM_RAZAO_SOCIAL','CNPJ','Tier F3','TIER_QUANDO','BPO(S/N)','Prestador','Sistema']);
let rankKey='PL_valor';
let rankDir='desc'; // 'desc' = maior/Z→A primeiro, 'asc' = menor/A→Z primeiro
let lastRankRows=[];
// Ordenação via clique no header da tabela (th[data-key]) — clicar de novo no mesmo
// campo alterna a direção da ordenação.
function setRank(k){
  if(rankKey===k)rankDir=rankDir==='desc'?'asc':'desc';
  else{rankKey=k;rankDir=RANK_STR_FIELDS.has(k)?'asc':'desc';}
  rRanking();
}
function rRanking(){
  const q=(document.getElementById('rank-q').value||'').toLowerCase();
  let m=mf().filter(d=>!q||d.SG_EFPC.toLowerCase().includes(q)||(d.NM_RAZAO_SOCIAL||'').toLowerCase().includes(q));
  const isStr=RANK_STR_FIELDS.has(rankKey);
  if(isStr){
    const sign=rankDir==='asc'?1:-1;
    m.sort((a,b)=>sign*String(a[rankKey]||'').localeCompare(String(b[rankKey]||''),'pt-BR'));
  }else{
    const sign=rankDir==='desc'?1:-1;
    m.sort((a,b)=>sign*(Math.abs(b[rankKey])-Math.abs(a[rankKey])));
  }
  lastRankRows=m;
  document.querySelectorAll('#rank-thead-row th[data-key]').forEach(th=>{
    const active=th.dataset.key===rankKey;
    th.classList.toggle('th-sort-active',active);
    th.textContent=th.dataset.label+(active?(rankDir==='asc'?' ▲':' ▼'):'');
  });
  document.getElementById('rank-count').textContent=m.length+' de '+RAW.master.length+' EFPCs';
  document.getElementById('rank-body').innerHTML=m.map((d,i)=>`<tr>
  <td class="num" style="color:#B3B3B3">${i+1}</td>
  <td class="ef-open" style="font-weight:600" onclick="openFicha(${d.NU_MATRICULA_EFPC})">${d.SG_EFPC}</td>
  <td style="max-width:190px;overflow:hidden;text-overflow:ellipsis">${(d.NM_RAZAO_SOCIAL||'').substring(0,38)}</td>
  <td style="font-size:11.5px;color:var(--g600)">${d.CNPJ||'—'}</td>
  <td>${tierTag(d['Tier F3'])}</td>
  <td class="num">${d.TIER_QUANDO==='N/D'?'—':d.TIER_QUANDO}</td>
  <td style="color:${d['BPO(S/N)']==='S'?'var(--success)':'#B3B3B3'};font-weight:600">${d['BPO(S/N)']}</td>
  <td>${d.Prestador||'—'}</td><td style="color:var(--g600)">${d.Sistema||'—'}</td>
  <td class="num">${fmt.brl(d.PL_valor)}</td>
  <td class="num ${d.DESP_TOTAL>0?'pos':d.DESP_TOTAL<0?'neg':'neu'}">${fmt.brl(Math.abs(d.DESP_TOTAL))}</td>
  <td class="num ${d.DESP_PCT_PL>15?'neg':d.DESP_PCT_PL<5?'pos':'neu'}">${d.DESP_PCT_PL.toFixed(2)}%</td>
  <td class="num">${fmt.num(d.QT_ATIVOS)}</td>
  <td class="num ${d.RAZAO_MATURIDADE>1.5?'neg':''}">${d.RAZAO_MATURIDADE.toFixed(2)}</td>
  <td class="num">${fmt.num(d.QT_PLANOS||0)}<br><span style="font-size:9.5px;color:var(--g500)">${d.TIPOS_PLANOS||'—'}</span></td>
  <td>${contatoHtml(d)}</td></tr>`).join('');
}
function goE(mat){openFicha(mat);}
// ── DESPESAS ──
// Base de comparação com o Tier: 'f3' (Tier F3) ou 'tq' (Marina).
let despCmpBy='f3';
function setDespCmp(mode,el){
  despCmpBy=mode;
  document.querySelectorAll('#desp-cmp-tabs .tab').forEach(t=>t.classList.remove('active'));
  if(el)el.classList.add('active');
  rDespesas();
}
function despCmpMeta(){
  return despCmpBy==='tq'
    ?{field:'TIER_QUANDO',label:'Marina',get:d=>String(d.TIER_QUANDO||'N/D'),color:t=>({'1':'#1AAAB2','2':'#0788C9','3':'#FF9A47','N/D':'#B3B3B3'}[t]||'#B3B3B3')}
    :{field:'Tier F3',label:'Tier F3',get:d=>d['Tier F3']||'N/D',color:t=>TC[t]||'#1AAAB2'};
}
function rDespesas(){
  const selSet=MS.desp,selMat=selSet.size===1?[...selSet][0]:null;
  document.getElementById('desp-ficha-link').style.display=selMat?'inline':'none';
  const mats=selSet.size?new Set([...selSet].map(String)):mfMats();
  const rows=RAW.master.filter(d=>mats.has(String(d.NU_MATRICULA_EFPC)));
  const ef=selMat?RAW.master.find(d=>d.NU_MATRICULA_EFPC===selMat):null;
  const pl=rows.reduce((s,d)=>s+d.PL_valor,0);
  const desp=rows.reduce((s,d)=>s+d.DESP_TOTAL,0);
  const dp=pl>0?desp/pl*100:0;
  document.getElementById('desp-meta').innerHTML=ef?`
  <div class="em-it"><div class="l">Tier F3</div><div class="v" style="color:${TC[ef['Tier F3']||'N/D']}">${ef['Tier F3']||'N/D'}</div></div>
  <div class="em-it"><div class="l">Marina</div><div class="v">${ef.TIER_QUANDO==='N/D'?'—':ef.TIER_QUANDO}</div></div>
  <div class="em-it"><div class="l">BPO</div><div class="v">${ef['BPO(S/N)']}</div></div>
  <div class="em-it"><div class="l">Prestador</div><div class="v">${ef.Prestador||'—'}${ef.Sistema?' · '+ef.Sistema:''}</div></div>
  <div class="em-it"><div class="l">Fit Marina</div><div class="v">${ef.FIT_MARINA==='S'?'Sim':'Não'}</div></div>`:(selSet.size>1?`<div class="em-it"><div class="l">Seleção</div><div class="v">${selSet.size} EFPCs</div></div>`:'');
  document.getElementById('kpi-desp').innerHTML=`
  <div class="kpi teal"><div class="kpi-l">Ativos Sob Gestão</div><div class="kpi-v">R$ ${fmt.brl(pl)}</div></div>
  <div class="kpi ${Math.abs(dp)>15?'red':Math.abs(dp)<5?'green':'orange'}"><div class="kpi-l">Despesa Total</div><div class="kpi-v">R$ ${fmt.brl(Math.abs(desp))}</div><div class="kpi-s">${dp.toFixed(2)}% do PL</div></div>`;
  const despMeta=RAW.despesas_meta;
  document.getElementById('desp-ann-note').textContent=(despMeta&&despMeta.mes)
    ?`Valores anualizados a partir da competência ${String(despMeta.mes).padStart(2,'0')}/${despMeta.ano} do balancete consolidado (fator ${despMeta.fator.toFixed(2).replace('.',',')}×, equivalente a 12/${despMeta.mes}).`
    :'';
  let despRows=RAW.despesas_estrutura.filter(d=>mats.has(String(d.NU_MATRICULA_EFPC)));
  // Comparação com o Tier — só faz sentido com uma única EFPC selecionada (é a
  // referência "sua EFPC" contra as demais do mesmo tier). A base de comparação
  // (Tier F3 ou Marina) é escolhida pelo usuário em despCmpBy. O mesmo tierMedian
  // alimenta tanto as colunas "Mediana Tier"/"Vs. Mediana" da tabela (item a item, ex.:
  // Mão-de-Obra Temporária) quanto o gráfico e o insight de despesa total abaixo.
  const cmp=despCmpMeta();
  document.getElementById('desp-th-mediana').textContent=`Mediana ${cmp.label} (%AuM)`;
  const tierMedian=ef?tierMedianData(cmp.get(ef),cmp.field):null;
  renderDespGrupos('desp-grupos-tbody',despRows,pl,tierMedian,mats);
  const tierEmpty=document.getElementById('desp-tier-empty'),tierBody=document.getElementById('desp-tier-body');
  if(ef){
    tierEmpty.style.display='none';
    tierBody.style.display='flex';
    const tier=cmp.get(ef);
    document.getElementById('desp-tier-sub').textContent=`Desp%PL vs. EFPCs do ${cmp.label} ${tier}`;
    const peers=RAW.master.filter(d=>cmp.get(d)===tier&&d.PL_valor>0);
    const sorted=peers.slice().sort((a,b)=>b.DESP_PCT_PL-a.DESP_PCT_PL);
    const vSorted=sorted.map(d=>d.DESP_PCT_PL).slice().sort((a,b)=>a-b);
    const md=vSorted.length?(vSorted.length%2?vSorted[(vSorted.length-1)/2]:(vSorted[vSorted.length/2-1]+vSorted[vSorted.length/2])/2):0;
    const rank=sorted.findIndex(d=>d.NU_MATRICULA_EFPC===ef.NU_MATRICULA_EFPC)+1;
    const diff=ef.DESP_PCT_PL-md;
    const insight=document.getElementById('desp-tier-insight');
    insight.style.display='flex';
    insight.className='insight'+(diff>0?' warn':'');
    insight.innerHTML=`<div class="ico">${diff>0?'⚠':'✓'}</div><div><div class="tt">${ef.SG_EFPC} está ${Math.abs(diff).toFixed(2)} p.p. ${diff>=0?'acima':'abaixo'} da mediana do ${cmp.label} ${tier}</div><div class="tx">Posição ${rank}ª de ${sorted.length} EFPCs do ${cmp.label} ${tier} por Desp%PL · mediana do tier: ${md.toFixed(2)}%</div></div>`;
    // Recorta ao Top 20 do tier (por Desp%PL) para não poluir o gráfico em tiers grandes,
    // mas garante que a própria EFPC sempre apareça mesmo se estiver fora do recorte.
    let chartRows=sorted.slice(0,20);
    if(!chartRows.some(d=>d.NU_MATRICULA_EFPC===ef.NU_MATRICULA_EFPC))
      chartRows=chartRows.slice(0,19).concat([ef]).sort((a,b)=>b.DESP_PCT_PL-a.DESP_PCT_PL);
    mk('ch-desp-tier','bar',{labels:chartRows.map(d=>d.SG_EFPC),datasets:[{data:chartRows.map(d=>d.DESP_PCT_PL),backgroundColor:chartRows.map(d=>d.NU_MATRICULA_EFPC===ef.NU_MATRICULA_EFPC?'#FF725E':cmp.color(tier)),borderRadius:5}]},{indexAxis:'x',plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>` ${c.raw.toFixed(2)}%`}}},scales:{y:{ticks:{callback:v=>v+'%'}}}});
  }else{
    tierEmpty.style.display='block';
    tierBody.style.display='none';
  }
}
// ── INVESTIMENTOS (aloc. atual + série histórica, aba unificada) ──
function compLabel(c){const s=String(c);return s.slice(4,6)+'/'+s.slice(2,4);}
function rInvest(){
  const selSet=MS.inv,selMat=selSet.size===1?[...selSet][0]:null;
  document.getElementById('inv-ficha-link').style.display=selMat?'inline':'none';
  const m=mf();
  const mats=selSet.size?new Set([...selSet].map(String)):mfMats();
  const ef=selMat?RAW.master.find(d=>d.NU_MATRICULA_EFPC===selMat):null;
  document.getElementById('inv-meta').innerHTML=ef?`
  <div class="em-it"><div class="l">Tier F3</div><div class="v" style="color:${TC[ef['Tier F3']||'N/D']}">${ef['Tier F3']||'N/D'}</div></div>
  <div class="em-it"><div class="l">PL atual</div><div class="v">R$ ${fmt.brl(ef.PL_valor)}</div></div>`:(selSet.size>1?`<div class="em-it"><div class="l">Seleção</div><div class="v">${selSet.size} EFPCs</div></div>`:'');

  // Alocação atual (foto do balancete mais recente)
  const byClass={};INV_CLASSES.forEach(c=>byClass[c]=0);
  RAW.investimentos.forEach(d=>{if(mats.has(String(d.NU_MATRICULA_EFPC)))byClass[d.CLASSE]=(byClass[d.CLASSE]||0)+d.VL_SALDO_FINAL;});
  const cl=INV_CLASSES,dataVals=cl.map(c=>byClass[c]||0),tot=dataVals.reduce((s,v)=>s+v,0);
  mk('ch-ipie','doughnut',{labels:cl,datasets:[{data:dataVals.map(v=>v/1e9),backgroundColor:cl.map(c=>IC[c]),borderColor:'#fff',borderWidth:2}]},{...ns(),cutout:'62%',plugins:{tooltip:{callbacks:{label:c=>` ${c.label}: R$ ${c.raw.toFixed(1)}B`}}}},'ch-ipie-leg');
  mk('ch-ibar','bar',{labels:cl,datasets:[{data:dataVals.map(v=>tot>0?v/tot*100:0),backgroundColor:cl.map(c=>IC[c]),borderRadius:5}]},{indexAxis:'y',plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>` ${c.raw.toFixed(1)}%`}}},scales:{x:{ticks:{callback:v=>v+'%'}}}});
  const t20=m.slice().sort((a,b)=>b.PL_valor-a.PL_valor).slice(0,20);
  const aC=[...new Set(RAW.investimentos.map(d=>d.CLASSE))];
  mk('ch-istk','bar',{labels:t20.map(d=>d.SG_EFPC),datasets:aC.map(cls=>({label:cls,data:t20.map(ef2=>{const r=RAW.investimentos.find(d=>d.NU_MATRICULA_EFPC===ef2.NU_MATRICULA_EFPC&&d.CLASSE===cls);const t=RAW.investimentos.filter(d=>d.NU_MATRICULA_EFPC===ef2.NU_MATRICULA_EFPC).reduce((s,d)=>s+d.VL_SALDO_FINAL,0);return t>0?(r?.VL_SALDO_FINAL||0)/t*100:0;}),backgroundColor:IC[cls]}))},{plugins:{legend:{position:'bottom'}},scales:{x:{stacked:true},y:{stacked:true,max:100,ticks:{callback:v=>v+'%'}}}});

  // Série histórica — mesma seleção de EFPC (mats/ef) usada acima na alocação atual.
  if(!RAW.hist_competencias||!RAW.hist_competencias.length){
    document.getElementById('hist-insights').innerHTML='<div class="insight warn"><div class="ico">ℹ</div><div><div class="tt">Série histórica não carregada</div><div class="tx">Verifique se a pasta "Fundos Exclusivos" contém os arquivos mensais no formato YYYYMM.csv.</div></div></div>';
    document.getElementById('kpi-hist').innerHTML='';
    return;
  }
  const comps=RAW.hist_competencias;

  // série de PL (EFPC específica ou soma do grupo selecionado/mercado filtrado)
  let plSerie;
  if(selMat){
    plSerie=comps.map(c=>{const r=RAW.hist_pl.find(d=>d.COD_EFPC===selMat&&d.COMPETENCIA===c);return r?r.VL_TOTAL:null;});
  }else{
    plSerie=comps.map(c=>RAW.hist_pl.filter(d=>d.COMPETENCIA===c&&mats.has(String(d.COD_EFPC))).reduce((s,d)=>s+d.VL_TOTAL,0));
  }
  const ultimo=plSerie[plSerie.length-1]||0,primeiro=plSerie.find(v=>v!==null&&v>0)||0;
  const varTotal=primeiro>0?(ultimo-primeiro)/primeiro*100:0;

  // fundos exclusivos
  let feRows=RAW.hist_exclusivo.filter(d=>mats.has(String(d.COD_EFPC)));
  const pctExclMedio=feRows.length?feRows.reduce((s,d)=>s+d.PCT_EXCLUSIVO,0)/feRows.length:0;

  document.getElementById('kpi-hist').innerHTML=`
  <div class="kpi teal"><div class="kpi-l">Patrimônio Atual</div><div class="kpi-v">R$ ${fmt.brl(ultimo)}</div><div class="kpi-s">${comps.length} competências na série</div></div>
  <div class="kpi ${varTotal>=0?'green':'red'}"><div class="kpi-l">Variação no Período</div><div class="kpi-v">${varTotal>=0?'+':''}${varTotal.toFixed(1)}%</div><div class="kpi-s">${compLabel(comps[0])} → ${compLabel(comps[comps.length-1])}</div></div>
  <div class="kpi navy"><div class="kpi-l">% em Fundos Exclusivos</div><div class="kpi-v">${pctExclMedio.toFixed(1)}%</div><div class="kpi-s">última competência</div></div>`;
  document.getElementById('hist-insights').innerHTML='';

  // Evolução AuM
  mk('ch-hist-pl','line',{labels:comps.map(compLabel),datasets:[{label:'Patrimônio',data:plSerie,borderColor:'#1AAAB2',backgroundColor:'rgba(26,170,178,.12)',fill:true,tension:.3,pointRadius:4,pointBackgroundColor:'#1AAAB2'}]},{plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>` R$ ${fmt.brl(c.raw)}`}}},scales:{y:{ticks:{callback:v=>fmt.brl(v)}}}});
}
// ── PARTICIPANTES ──
let partModF=new Set(['CD','BD','CV']);
function tgModPart(el){
  const v=el.dataset.mod;
  if(partModF.has(v)){partModF.delete(v);el.classList.remove('on');el.classList.add('off');}
  else{partModF.add(v);el.classList.add('on');el.classList.remove('off');}
  rPart();
}
function rPart(){
  const selSet=MS.part,selMat=selSet.size===1?[...selSet][0]:null;
  document.getElementById('part-ficha-link').style.display=selMat?'inline':'none';
  const planoSel=MSPL.part;
  // Modo Plano: liga sempre que houver plano(s) marcado(s) explicitamente OU alguma
  // modalidade estiver desmarcada — nesses casos os totais de Ativos/Assistidos vêm da
  // base de planos (única com granularidade de modalidade), não do totalizador da EFPC.
  const planoMode=planoSel.size>0||partModF.size<3;
  let mats,rows,tA,tAs,planosPool=null;
  if(planoMode){
    let pool=RAW.planos.filter(p=>partModF.has(p.MODALIDADE));
    if(planoSel.size)pool=pool.filter(p=>planoSel.has(p.NU_CNPB));
    else if(selSet.size)pool=pool.filter(p=>selSet.has(p.NU_MATRICULA_EFPC));
    else{const gm=mfMats();pool=pool.filter(p=>gm.has(String(p.NU_MATRICULA_EFPC)));}
    planosPool=pool;
    mats=new Set(pool.map(p=>String(p.NU_MATRICULA_EFPC)));
    rows=RAW.master.filter(d=>mats.has(String(d.NU_MATRICULA_EFPC)));
    tA=pool.reduce((s,p)=>s+p.QT_ATIVOS,0);
    tAs=pool.reduce((s,p)=>s+p.QT_ASSISTIDOS,0);
  }else{
    mats=selSet.size?new Set([...selSet].map(String)):mfMats();
    rows=RAW.master.filter(d=>mats.has(String(d.NU_MATRICULA_EFPC)));
    tA=rows.reduce((s,d)=>s+d.QT_ATIVOS,0);
    tAs=rows.reduce((s,d)=>s+d.QT_ASSISTIDOS,0);
  }
  const r=tA>0?tAs/tA:0;
  document.getElementById('kpi-part').innerHTML=`
  <div class="kpi navy"><div class="kpi-l">Total de Participantes</div><div class="kpi-v">${fmt.num(tA+tAs)}</div></div>
  <div class="kpi teal"><div class="kpi-l">Ativos</div><div class="kpi-v">${fmt.num(tA)}</div></div>
  <div class="kpi green"><div class="kpi-l">Assistidos</div><div class="kpi-v">${fmt.num(tAs)}</div></div>
  <div class="kpi ${r>1.5?'red':r>1?'orange':'green'}"><div class="kpi-l">Maturidade</div><div class="kpi-v">${r.toFixed(2)}</div><div class="kpi-s">${r>1.5?'fundo maduro':r>1?'atenção':'saudável'}</div></div>
  ${planoMode?`<div class="kpi navy"><div class="kpi-l">Planos no Filtro</div><div class="kpi-v">${fmt.num(planosPool.length)}</div></div>`:''}`;
  const fx=['ATÉ 24 ANOS','ENTRE 25 E 34 ANOS','ENTRE 35 E 54 ANOS','ENTRE 55 E 64 ANOS','ENTRE 65 E 74 ANOS','ENTRE 75 E 84 ANOS','MAIOR QUE 85 ANOS'];
  const fd=Object.entries(RAW.faixa_etaria.reduce((a,d)=>{if(mats.has(String(d.NU_MATRICULA_EFPC)))a[d.NM_FAIXA_ETARIA]=(a[d.NM_FAIXA_ETARIA]||0)+d.QT_PESSOAS;return a;},{})).map(([k,v])=>({NM_FAIXA_ETARIA:k,QT_PESSOAS:v}));
  const fmap=Object.fromEntries(fd.map(d=>[d.NM_FAIXA_ETARIA,d.QT_PESSOAS]));
  mk('ch-pira','bar',{labels:fx.map(f=>f.replace('ENTRE ','').replace(' ANOS','')),datasets:[{data:fx.map(f=>fmap[f]||0),backgroundColor:'#1AAAB2',borderRadius:5}]},{indexAxis:'y',plugins:{legend:{display:false}}});
  // Ativos × Assistidos: agregado por padrão. Ao buscar/marcar uma única EFPC (selMat),
  // troca para um plano por barra (Ativos e Assistidos lado a lado), com os valores brutos
  // embutidos no próprio rótulo do eixo — e cada barra funciona como um seletor: clicar
  // (des)marca o plano em MSPL.part, exatamente como marcar pelo dropdown de busca.
  const tipoH=document.getElementById('tipo-h'),tipoS=document.getElementById('tipo-s'),tipoLeg=document.getElementById('ch-tipo-leg');
  const planosDoEfpc=selMat?RAW.planos.filter(p=>p.NU_MATRICULA_EFPC===selMat&&partModF.has(p.MODALIDADE)):[];
  if(selMat&&planosDoEfpc.length){
    const efSg=(RAW.master.find(d=>d.NU_MATRICULA_EFPC===selMat)||{}).SG_EFPC||'';
    tipoH.querySelector('.cc-t').textContent=`Planos de ${efSg} · Ativos × Assistidos`;
    tipoS.textContent='clique numa barra para (des)selecionar o plano';
    const nomePlano=p=>p.SG_PLANO||p.NM_PLANO||('CNPB '+p.NU_CNPB);
    const plist=planosDoEfpc.slice().sort((a,b)=>(b.QT_ATIVOS+b.QT_ASSISTIDOS)-(a.QT_ATIVOS+a.QT_ASSISTIDOS)).slice(0,25);
    mk('ch-tipo','bar',{
      labels:plist.map(p=>`${nomePlano(p)}  (${fmt.num(p.QT_ATIVOS)}/${fmt.num(p.QT_ASSISTIDOS)})`),
      datasets:[
        {label:'Ativos',data:plist.map(p=>p.QT_ATIVOS),backgroundColor:plist.map(p=>MSPL.part.has(p.NU_CNPB)?'#1AAAB2':'#BFE6E8'),borderRadius:4},
        {label:'Assistidos',data:plist.map(p=>p.QT_ASSISTIDOS),backgroundColor:plist.map(p=>MSPL.part.has(p.NU_CNPB)?'#153451':'#C3CBD3'),borderRadius:4},
      ]
    },{indexAxis:'x',plugins:{legend:{display:true},tooltip:{callbacks:{label:c=>` ${c.dataset.label}: ${fmt.num(c.raw)}`}}},
      scales:{y:{ticks:{callback:v=>fmt.num(v)}}},
      onClick:(evt,els)=>{if(!els.length)return;const p=plist[els[0].index];if(p)msplToggle('part',p.NU_CNPB,!MSPL.part.has(p.NU_CNPB));},
      onHover:(evt,els)=>{if(evt.native&&evt.native.target)evt.native.target.style.cursor=els.length?'pointer':'default';},
    },'ch-tipo-leg');
  }else{
    tipoH.querySelector('.cc-t').textContent='Ativos × Assistidos';
    tipoS.textContent=selMat?'esta EFPC não tem planos cadastrados no filtro de modalidade atual':'busque/marque uma única EFPC para ver por plano';
    if(tipoLeg)tipoLeg.innerHTML='';
    mk('ch-tipo','bar',{labels:[`Ativos  (${fmt.num(tA)})`,`Assistidos  (${fmt.num(tAs)})`],datasets:[{data:[tA,tAs],backgroundColor:['#1AAAB2','#153451'],borderRadius:5}]},{indexAxis:'x',plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>' '+fmt.num(c.raw)}}},scales:{y:{ticks:{callback:v=>fmt.num(v)}}}});
  }
  // Top 15 EFPCs com mais participantes: propositalmente alheio à seleção de fundos/planos
  // dentro de Participantes — reflete sempre o universo dos filtros globais (topo), para
  // servir de referência estável enquanto o usuário navega pelas seleções. Os valores
  // brutos (ativos/assistidos/total) vão embutidos no rótulo e no tooltip.
  const matuH=document.getElementById('matu-h');
  matuH.querySelector('.cc-t').textContent='Mais Participantes · Top 15';
  const m=mf();
  const t15=m.slice().sort((a,b)=>(b.QT_ATIVOS+b.QT_ASSISTIDOS)-(a.QT_ATIVOS+a.QT_ASSISTIDOS)).slice(0,15);
  mk('ch-matu','bar',{labels:t15.map(d=>`${d.SG_EFPC}  (${fmt.num(d.QT_ATIVOS+d.QT_ASSISTIDOS)})`),datasets:[{data:t15.map(d=>d.QT_ATIVOS+d.QT_ASSISTIDOS),backgroundColor:'#1AAAB2',borderRadius:5}]},{indexAxis:'y',plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>{const d=t15[c.dataIndex];return[` Total: ${fmt.num(d.QT_ATIVOS+d.QT_ASSISTIDOS)}`,` Ativos: ${fmt.num(d.QT_ATIVOS)}`,` Assistidos: ${fmt.num(d.QT_ASSISTIDOS)}`];}}}}});
  // Evolução histórica — em modo Plano usa a série por CNPB, senão a série por EFPC.
  const serieEvo=planoMode
    ?(()=>{const cnpbs=new Set(planosPool.map(p=>p.NU_CNPB));return RAW.epb_evolucao_planos.filter(d=>cnpbs.has(d.NU_CNPB));})()
    :RAW.epb_evolucao.filter(d=>mats.has(String(d.NU_MATRICULA_EFPC)));
  const comps=[...new Set(serieEvo.map(d=>d.DT_COMPETENCIA))].sort();
  const lbl=comps.map(c=>{const s=String(c);return s.slice(4,6)+'/'+s.slice(2,4);});
  const somaTipo=(c,tipo)=>serieEvo.filter(d=>d.DT_COMPETENCIA===c&&d.TIPO===tipo).reduce((s,d)=>s+d.QT_ATUAL,0);
  const ativosPorComp=comps.map(c=>somaTipo(c,'ATIVOS'));
  const assistidosPorComp=comps.map(c=>somaTipo(c,'ASSISTIDOS'));
  const totalPorComp=ativosPorComp.map((v,i)=>v+assistidosPorComp[i]);
  // Sem anualização financeira aqui: são contagens de pessoas num instante (foto de
  // cada competência do EPB), não saldo acumulado no ano como as despesas do
  // balancete — não há "taxa anual" a projetar. "Anualizado" nesse gráfico é só
  // reduzir a janela mensal cheia a 1 ponto por ano — a última competência
  // disponível em cada ano (normalmente dezembro; se o ano corrente ainda não
  // fechou, usa o mês mais recente que já tiver).
  const idxAnual=[];
  let anoAtual=null;
  comps.forEach((c,i)=>{
    const ano=Math.floor(c/100);
    if(ano!==anoAtual){idxAnual.push(i);anoAtual=ano;}
    else idxAnual[idxAnual.length-1]=i;
  });
  mk('ch-evo','line',{labels:idxAnual.map(i=>lbl[i]),datasets:[
    {label:'Ativos',data:idxAnual.map(i=>ativosPorComp[i]),borderColor:'#1AAAB2',backgroundColor:'rgba(26,170,178,.10)',fill:false,tension:0,pointRadius:4,pointBackgroundColor:'#1AAAB2'},
    {label:'Assistidos',data:idxAnual.map(i=>assistidosPorComp[i]),borderColor:'#153451',backgroundColor:'rgba(21,52,81,.10)',fill:false,tension:0,pointRadius:4,pointBackgroundColor:'#153451'},
    {label:'Total',data:idxAnual.map(i=>totalPorComp[i]),borderColor:'#FF9A47',backgroundColor:'rgba(255,154,71,.08)',fill:false,tension:0,pointRadius:4,pointBackgroundColor:'#FF9A47',borderDash:[5,4]},
  ]},{plugins:{legend:{display:true,position:'bottom'}}});
}
// ── DIRIGENTES ──
let dirSt='todos';
function setDirSt(s,el){dirSt=s;document.querySelectorAll('#page-dirigentes .tabs .tab').forEach(t=>t.classList.remove('active'));if(el)el.classList.add('active');rDirTable();}
function dirF(){
  const mats=mfMats();
  return RAW.dirigentes.filter(d=>mats.has(String(d.mat)));
}
function rDirTable(){
  const q=(document.getElementById('dir-q').value||'').toLowerCase();
  const tp=document.getElementById('dir-tipo').value;
  let dd=dirF();
  if(dirSt!=='todos')dd=dd.filter(d=>d.st===dirSt);
  if(tp)dd=dd.filter(d=>d.tp===tp);
  if(q)dd=dd.filter(d=>d.nm.toLowerCase().includes(q)||d.ef.toLowerCase().includes(q));
  dd=dd.slice().sort((a,b)=>(a.dias===99999?1e9:a.dias)-(b.dias===99999?1e9:b.dias));
  const stB={VIGENTE:'<span class="stbadge st-vigente">● Vigente</span>',VENCENDO:'<span class="stbadge st-vencendo">◷ Vencendo</span>',VENCIDO:'<span class="stbadge st-vencido">✕ Vencido</span>',SEM_DATA:'<span class="stbadge st-semdata">— Sem data</span>'};
  document.getElementById('dir-body').innerHTML=dd.slice(0,500).map(d=>`<tr>
  <td style="font-weight:500;color:var(--navy)">${d.nm}</td>
  <td class="ef-open" style="font-weight:600" onclick="openFicha(${d.mat})">${d.ef}</td>
  <td style="font-size:11.5px">${d.tp}</td>
  <td style="color:${d.pr==='S'?'var(--teal-600)':'#B3B3B3'};font-weight:600">${d.pr==='S'?'Sim':'—'}</td>
  <td style="color:${d.aetq==='S'?'var(--info)':'#B3B3B3'};font-weight:600">${d.aetq==='S'?'Sim':'—'}</td>
  <td style="color:${d.rem==='S'?'var(--success)':'#B3B3B3'}">${d.rem==='S'?'Sim':'—'}</td>
  <td>${fmt.dt(d.ini)}</td><td>${fmt.dt(d.fim)}</td>
  <td class="num ${d.dias<0?'neg':d.dias<=180&&d.dias!==99999?'':'pos'}" style="${d.dias<=180&&d.dias>=0?'color:#C26A14':''}">${d.dias===99999?'—':fmt.num(d.dias)}</td>
  <td>${stB[d.st]}</td></tr>`).join('')+(dd.length>500?`<tr><td colspan="10" style="text-align:center;color:#969696;font-style:italic">… ${fmt.num(dd.length-500)} registros adicionais — refine a busca</td></tr>`:'');
}
// ── PLANOS ──
// Mesmo esquema de ordenação por clique no header da Ranking de EFPCs (ver RANK_STR_FIELDS
// /setRank) — sem abas/"caixinhas" acima da tabela, só clicando no th mesmo.
const PLANO_STR_FIELDS=new Set(['NM_PLANO','CNPJ','SG_EFPC','MODALIDADE','TIPO_PATROCINIO','SITUACAO']);
let planoRankKey='PL_valor';
let planoRankDir='desc';
let lastPlanoRows=[];
let modF=new Set(['CD','BD','CV']);
function tgMod(el){
  const v=el.dataset.mod;
  if(modF.has(v)){modF.delete(v);el.classList.remove('on');el.classList.add('off');}
  else{modF.add(v);el.classList.add('on');el.classList.remove('off');}
  rPlanos();
}
function setPlanoRank(k){
  if(planoRankKey===k)planoRankDir=planoRankDir==='desc'?'asc':'desc';
  else{planoRankKey=k;planoRankDir=PLANO_STR_FIELDS.has(k)?'asc':'desc';}
  rPlanos();
}
function planosF(){
  const mats=mfMats();
  return RAW.planos.filter(d=>mats.has(String(d.NU_MATRICULA_EFPC))&&modF.has(d.MODALIDADE));
}
function modTag(m){const c=MC[m]||'#B3B3B3';return`<span class="tag" style="background:${c}1A;color:${c};border:1px solid ${c}55">${m||'N/D'}</span>`;}
const PC={'Patrocinado':'#0788C9','Instituído':'#FF9A47','N/D':'#B3B3B3'};
function patrocinioTag(p){const c=PC[p]||'#B3B3B3';return`<span class="tag" style="background:${c}1A;color:${c};border:1px solid ${c}55">${p||'N/D'}</span>`;}
function rPlanos(){
  const q=(document.getElementById('plano-q').value||'').toLowerCase();
  let d=planosF().filter(x=>!q||x.SG_EFPC.toLowerCase().includes(q)||(x.NM_PLANO||'').toLowerCase().includes(q)||(x.SG_PLANO||'').toLowerCase().includes(q));
  const totPL=d.reduce((s,x)=>s+x.PL_valor,0),totAt=d.reduce((s,x)=>s+x.QT_ATIVOS,0),totAs=d.reduce((s,x)=>s+x.QT_ASSISTIDOS,0);
  const efSet=new Set(d.map(x=>x.NU_MATRICULA_EFPC));
  document.getElementById('plano-count').textContent=efSet.size+' EFPCs · '+d.length+' planos';
  document.getElementById('kpi-planos').innerHTML=`
  <div class="kpi teal"><div class="kpi-l">Planos</div><div class="kpi-v">${fmt.num(d.length)}</div></div>
  <div class="kpi blue"><div class="kpi-l">Patrimônio dos Planos</div><div class="kpi-v">R$ ${fmt.brl(totPL)}</div></div>
  <div class="kpi green"><div class="kpi-l">Participantes Ativos</div><div class="kpi-v">${fmt.num(totAt)}</div></div>
  <div class="kpi orange"><div class="kpi-l">Assistidos</div><div class="kpi-v">${fmt.num(totAs)}</div></div>
  <div class="kpi navy"><div class="kpi-l">CD / BD / CV</div><div class="kpi-v" style="font-size:16px">${d.filter(x=>x.MODALIDADE==='CD').length} / ${d.filter(x=>x.MODALIDADE==='BD').length} / ${d.filter(x=>x.MODALIDADE==='CV').length}</div></div>
  <div class="kpi orange"><div class="kpi-l">Patrocinado / Instituído</div><div class="kpi-v" style="font-size:16px">${d.filter(x=>x.TIPO_PATROCINIO==='Patrocinado').length} / ${d.filter(x=>x.TIPO_PATROCINIO==='Instituído').length}</div></div>`;
  // ch-pmod: só as modalidades ativas no filtro entram no eixo — ao desmarcar uma
  // modalidade o gráfico passa a ter só as colunas restantes (não uma barra zerada).
  // {persist:true} mantém a mesma instância do Chart.js entre re-renders para que essa
  // redução de colunas seja animada (encolhe/desliza) em vez de um corte abrupto.
  const modsAtivos=['CD','BD','CV'].filter(m=>modF.has(m));
  mk('ch-pmod','bar',{labels:modsAtivos,datasets:[{data:modsAtivos.map(m=>d.filter(x=>x.MODALIDADE===m).reduce((s,x)=>s+x.PL_valor,0)/1e9),backgroundColor:modsAtivos.map(m=>MC[m]),borderRadius:6}]},{persist:true,plugins:{legend:{display:false}}});
  const mods=['CD','BD','CV'];
  mk('ch-pmodqt','doughnut',{labels:mods,datasets:[{data:mods.map(m=>d.filter(x=>x.MODALIDADE===m).length),backgroundColor:mods.map(m=>MC[m]),borderColor:'#fff',borderWidth:2}]},{...ns()},'ch-pmodqt-leg');
  if(PLANO_STR_FIELDS.has(planoRankKey)){
    const sign=planoRankDir==='asc'?1:-1;
    d=d.slice().sort((a,b)=>sign*String(a[planoRankKey]||'').localeCompare(String(b[planoRankKey]||''),'pt-BR'));
  }else{
    const sign=planoRankDir==='desc'?1:-1;
    d=d.slice().sort((a,b)=>sign*(Math.abs(b[planoRankKey])-Math.abs(a[planoRankKey])));
  }
  document.querySelectorAll('#plano-thead-row th[data-key]').forEach(th=>{
    const active=th.dataset.key===planoRankKey;
    th.classList.toggle('th-sort-active',active);
    th.textContent=th.dataset.label+(active?(planoRankDir==='asc'?' ▲':' ▼'):'');
  });
  lastPlanoRows=d;
  document.getElementById('plano-body').innerHTML=d.slice(0,500).map((x,i)=>`<tr>
  <td class="num" style="color:#B3B3B3">${i+1}</td>
  <td style="max-width:220px;overflow:hidden;text-overflow:ellipsis;font-weight:600">${x.NM_PLANO||x.SG_PLANO||'—'}</td>
  <td style="font-size:11.5px;color:var(--g600)">${x.CNPJ||'—'}</td>
  <td class="ef-open" style="font-weight:600" onclick="openFicha(${x.NU_MATRICULA_EFPC})">${x.SG_EFPC}</td>
  <td>${modTag(x.MODALIDADE)}</td>
  <td>${patrocinioTag(x.TIPO_PATROCINIO)}</td>
  <td style="color:var(--g600);font-size:11px">${x.SITUACAO||'—'}</td>
  <td class="num">${fmt.brl(x.PL_valor)}</td>
  <td class="num">${fmt.num(x.QT_ATIVOS)}</td>
  <td class="num">${fmt.num(x.QT_ASSISTIDOS)}</td>
  <td class="num">${fmt.num(x.QT_TOTAL)}</td></tr>`).join('')+(d.length>500?`<tr><td colspan="10" style="text-align:center;color:#969696;font-style:italic">… ${fmt.num(d.length-500)} registros adicionais — refine a busca</td></tr>`:'');
}
// ── EXPORTAÇÃO CSV/XLSX ──
const EXPORT_COLS={
  ranking:[['SG_EFPC','EFPC'],['NM_RAZAO_SOCIAL','Razão Social'],['CNPJ','CNPJ'],['Tier F3','Tier F3'],['TIER_QUANDO','Marina'],
    ['BPO(S/N)','BPO'],['Prestador','Prestador'],['Sistema','Sistema'],['PL_valor','PL'],['DESP_TOTAL','Despesas'],
    ['DESP_PCT_PL','Desp % PL'],['QT_ATIVOS','Ativos'],['QT_ASSISTIDOS','Assistidos'],['RAZAO_MATURIDADE','Maturidade'],
    ['QT_PLANOS','Qtd. Planos'],['TIPOS_PLANOS','Tipos de Planos'],['SITE','Site'],['EMAIL','E-mail'],['FONE','Telefone']],
  planos:[['SG_PLANO','Sigla Plano'],['NM_PLANO','Plano'],['CNPJ','CNPJ'],['SG_EFPC','EFPC'],['NM_RAZAO_SOCIAL','Razão Social EFPC'],
    ['MODALIDADE','Modalidade'],['TIPO_PATROCINIO','Patrocínio'],['SITUACAO','Situação'],['PL_valor','PL'],['QT_ATIVOS','Ativos'],
    ['QT_ASSISTIDOS','Assistidos'],['QT_TOTAL','Total Participantes']],
};
function exportTable(kind,format){
  const rows=kind==='ranking'?lastRankRows:lastPlanoRows;
  const cols=EXPORT_COLS[kind];
  if(!rows||!rows.length){alert('Nenhum registro para exportar.');return;}
  const data=rows.map(r=>{const o={};cols.forEach(([k,label])=>{o[label]=r[k]??'';});return o;});
  const stamp=new Date().toISOString().slice(0,10);
  const fname=`${kind==='ranking'?'ranking_efpcs':'planos_efpc'}_${stamp}`;
  if(format==='xlsx'){
    const ws=XLSX.utils.json_to_sheet(data);
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,ws,kind==='ranking'?'Ranking EFPCs':'Planos');
    XLSX.writeFile(wb,fname+'.xlsx');
  }else{
    const headers=cols.map(c=>c[1]);
    const esc=v=>{const s=v===null||v===undefined?'':String(v);return /[;"\n]/.test(s)?'"'+s.replace(/"/g,'""')+'"':s;};
    const lines=[headers.map(esc).join(';')].concat(data.map(row=>headers.map(h=>esc(row[h])).join(';')));
    const csv='﻿'+lines.join('\r\n');
    const blob=new Blob([csv],{type:'text/csv;charset=utf-8;'});
    const a=document.createElement('a');
    a.href=URL.createObjectURL(blob);a.download=fname+'.csv';document.body.appendChild(a);a.click();a.remove();URL.revokeObjectURL(a.href);
  }
}
// ── DESPESAS · EXPORT XLSX COM BREAKDOWN POR EFPC ──
// Uma linha por EFPC (respeitando os filtros gerais + a busca de EFPC(s) da aba Despesas),
// com colunas de PL/Despesa total e uma coluna para cada item de despesa (grupo · item),
// a partir de RAW.despesas_grupos_cfg — o mesmo detalhamento usado na tabela em tela.
function despItemCols(){
  const cols=[];
  RAW.despesas_grupos_cfg.forEach(g=>{
    g.itens.forEach(it=>cols.push({grupo:g.grupo,nome:it.nome}));
    cols.push({grupo:g.grupo,nome:g.residual});
  });
  return cols;
}
function exportDespesasXlsx(){
  const mats=MS.desp.size?new Set([...MS.desp].map(String)):mfMats();
  const efs=RAW.master.filter(d=>mats.has(String(d.NU_MATRICULA_EFPC))).sort((a,b)=>b.PL_valor-a.PL_valor);
  if(!efs.length){alert('Nenhuma EFPC no filtro atual para exportar.');return;}
  const cols=despItemCols();
  const rows=efs.map(ef=>{
    const byItem={};
    RAW.despesas_estrutura.filter(d=>d.NU_MATRICULA_EFPC===ef.NU_MATRICULA_EFPC)
      .forEach(d=>{byItem[d.GRUPO+'||'+d.ITEM]=(byItem[d.GRUPO+'||'+d.ITEM]||0)+d.VALOR;});
    const o={'EFPC':ef.SG_EFPC,'Razão Social':ef.NM_RAZAO_SOCIAL||'','Tier F3':ef['Tier F3']||'N/D',
      'Ativos Sob Gestão':ef.PL_valor,'Despesa Total':ef.DESP_TOTAL,'Despesa % PL':+(+ef.DESP_PCT_PL).toFixed(2)};
    cols.forEach(c=>{o[`${c.grupo} · ${c.nome}`]=byItem[c.grupo+'||'+c.nome]||0;});
    return o;
  });
  const ws=XLSX.utils.json_to_sheet(rows);
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Despesas por EFPC');
  const stamp=new Date().toISOString().slice(0,10);
  XLSX.writeFile(wb,`despesas_breakdown_efpc_${stamp}.xlsx`);
}
// ── INVESTIMENTOS · EXPORT XLSX COM BREAKDOWN POR EFPC ──
// Mesmo princípio do export de despesas: uma linha por EFPC (respeitando os filtros
// gerais + a busca de EFPC(s) da aba Investimentos), com o PL alocado em cada classe de
// investimento (INV_CLASSES) como coluna própria.
function exportInvestimentosXlsx(){
  const mats=MS.inv.size?new Set([...MS.inv].map(String)):mfMats();
  const efs=RAW.master.filter(d=>mats.has(String(d.NU_MATRICULA_EFPC))).sort((a,b)=>b.PL_valor-a.PL_valor);
  if(!efs.length){alert('Nenhuma EFPC no filtro atual para exportar.');return;}
  const rows=efs.map(ef=>{
    const byClass={};
    RAW.investimentos.filter(d=>d.NU_MATRICULA_EFPC===ef.NU_MATRICULA_EFPC)
      .forEach(d=>{byClass[d.CLASSE]=(byClass[d.CLASSE]||0)+d.VL_SALDO_FINAL;});
    const o={'EFPC':ef.SG_EFPC,'Razão Social':ef.NM_RAZAO_SOCIAL||'','Tier F3':ef['Tier F3']||'N/D',
      'Ativos Sob Gestão':ef.PL_valor};
    INV_CLASSES.forEach(c=>{o[c]=byClass[c]||0;});
    return o;
  });
  const ws=XLSX.utils.json_to_sheet(rows);
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Investimentos por EFPC');
  const stamp=new Date().toISOString().slice(0,10);
  XLSX.writeFile(wb,`investimentos_breakdown_efpc_${stamp}.xlsx`);
}
// ── DADOS MANUAIS — EDIÇÃO DIRETA DO EXCEL — desativado a pedido do usuário, mantido comentado para retomar depois
/*
const MAN_LABEL={tiers:'Tiers.xlsx',classificacao:'Classificação de Dados no Balancete.xlsx'};
let manCurrent='tiers';
const manState={tiers:{handle:null,cols:[],rows:[],sheetName:null,fileName:''},
  classificacao:{handle:null,cols:[],rows:[],sheetName:null,fileName:''}};
function manEscHtml(v){return String(v??'').replace(/[&<>]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]));}
function manEscAttr(v){return String(v??'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));}
function manSetTab(kind,el){
  manCurrent=kind;
  document.querySelectorAll('#man-tabs .tab').forEach(t=>t.classList.remove('active'));
  el.classList.add('active');
  document.getElementById('man-q').value='';
  document.getElementById('man-status').textContent=manState[kind].handle?('Aberto: '+manState[kind].fileName):'';
  manRender();
}
async function manAbrir(){
  if(!window.showOpenFilePicker){
    alert('Seu navegador não suporta esse recurso. Abra o dashboard no Google Chrome ou Microsoft Edge para editar os arquivos diretamente.');
    return;
  }
  try{
    const [handle]=await window.showOpenFilePicker({
      id:'man-'+manCurrent,
      excludeAcceptAllOption:false,
      types:[{description:'Excel',accept:{'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':['.xlsx']}}]
    });
    const file=await handle.getFile();
    const buf=await file.arrayBuffer();
    const wb=XLSX.read(buf,{type:'array',cellDates:true});
    const sheetName=wb.SheetNames[0];
    const ws=wb.Sheets[sheetName];
    const cols=(XLSX.utils.sheet_to_json(ws,{header:1,defval:''})[0]||[]).map(String);
    const rows=XLSX.utils.sheet_to_json(ws,{defval:''});
    manState[manCurrent]={handle,cols,rows,sheetName,fileName:file.name};
    document.getElementById('man-status').textContent='Aberto: '+file.name;
    manRender();
  }catch(err){
    if(err.name!=='AbortError'){console.error(err);alert('Não foi possível abrir o arquivo: '+err.message);}
  }
}
function manRender(){
  const st=manState[manCurrent];
  const wrap=document.getElementById('man-table-wrap');
  const addBtn=document.getElementById('man-addrow-btn');
  const saveBtn=document.getElementById('man-save-btn');
  const search=document.getElementById('man-q');
  if(!st.handle){
    wrap.innerHTML='<div style="padding:40px;text-align:center;color:var(--g500);font-size:13px">Nenhum arquivo carregado. Clique em "Selecionar arquivo…" e escolha '+MAN_LABEL[manCurrent]+'.</div>';
    addBtn.style.display='none';saveBtn.style.display='none';search.style.display='none';
    return;
  }
  addBtn.style.display='';saveBtn.style.display='';search.style.display='';
  const q=(search.value||'').toLowerCase().trim();
  const cols=st.cols;
  let html='<table><thead><tr>'+cols.map(c=>`<th>${manEscHtml(c)}</th>`).join('')+'<th></th></tr></thead><tbody>';
  st.rows.forEach((row,i)=>{
    if(q && !cols.some(c=>String(row[c]??'').toLowerCase().includes(q)))return;
    html+='<tr>'+cols.map(c=>`<td><input class="man-cell" data-row="${i}" data-col="${manEscAttr(c)}" value="${manEscAttr(row[c])}" oninput="manEdit(this)"></td>`).join('')+
      `<td><span class="man-del" onclick="manDelRow(${i})">excluir</span></td></tr>`;
  });
  html+='</tbody></table>';
  wrap.innerHTML=html;
}
function manEdit(el){
  const i=+el.dataset.row, col=el.dataset.col;
  let v=el.value;
  if(v.trim()!==''&&!isNaN(v)&&!/^0[0-9]/.test(v.trim()))v=Number(v);
  manState[manCurrent].rows[i][col]=v;
}
function manAddRow(){
  const st=manState[manCurrent];
  const obj={};
  st.cols.forEach(c=>obj[c]='');
  st.rows.push(obj);
  manRender();
  const wrap=document.getElementById('man-table-wrap');
  wrap.scrollTop=wrap.scrollHeight;
}
function manDelRow(i){
  if(!confirm('Excluir esta linha?'))return;
  manState[manCurrent].rows.splice(i,1);
  manRender();
}
async function manSalvar(){
  const st=manState[manCurrent];
  if(!st.handle){alert('Nenhum arquivo carregado.');return;}
  try{
    const ws=XLSX.utils.json_to_sheet(st.rows,{header:st.cols});
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,ws,st.sheetName||'Sheet1');
    const buf=XLSX.write(wb,{bookType:'xlsx',type:'array'});
    const writable=await st.handle.createWritable();
    await writable.write(buf);
    await writable.close();
    document.getElementById('man-status').textContent='✓ Salvo às '+new Date().toLocaleTimeString('pt-BR')+' — rode o gerador novamente para atualizar o dashboard.';
  }catch(err){
    console.error(err);
    alert('Erro ao salvar: '+err.message);
  }
}
function rManuais(){manRender();}
*/
// ── INIT ──
Chart.defaults.font.family='Inter';
document.getElementById('sb-sub').textContent='Quando · Previc '+compTxt((RAW.meta_atualizacao||{}).balancete_mes,(RAW.meta_atualizacao||{}).balancete_ano);
document.getElementById('pg-updated').textContent=pgUpdatedTxt('overview');
popSel();refresh();
// Fase de captura (3º arg `true`): precisa rodar ANTES do onclick da checkbox marcada
// dentro do dropdown, porque msToggle() reconstrói o innerHTML da .ac-list (para refletir
// o novo estado marcado) e isso desconecta o <input> clicado do DOM. Se este listener
// rodasse na fase de bubble (padrão), e.target já estaria desconectado quando chegasse
// aqui, e.target.closest('.ac-wrap') retornaria null, e o dropdown fecharia sozinho a
// cada seleção — exatamente o problema de "a aba minimiza ao marcar uma EFPC".
document.addEventListener('click',e=>{
  if(!e.target.closest('.ac-wrap'))
    document.querySelectorAll('.ac-list.open').forEach(l=>l.classList.remove('open'));
},true);
document.addEventListener('keydown',e=>{if(e.key==='Escape')closeFicha();});
</script>
</body>
</html>"""

# ─────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    rotear_dados_extracao(BASE)
    ARQUIVOS["dsi"] = _arquivo_mais_recente_dsi(os.path.join(BASE, "Dados Participantes", "Sexo e Idade dos participantes"))
    bases_tratadas = gerar_bases_tratadas(BASE)
    ARQUIVOS["consolidado"] = bases_tratadas["consolidado"]
    ARQUIVOS["pga"]         = bases_tratadas["pga"]
    ARQUIVOS["planos"]      = bases_tratadas["planos"]
    dfs  = carregar()
    data = processar(dfs)
    exportar_dados(data, OUTPUT_FILE)
    print(f"✓ Dados do dashboard atualizados: {OUTPUT_FILE}")
    print(f"  Tamanho: {os.path.getsize(OUTPUT_FILE)/1024:.0f} KB")
    # gerar_html() e HTML_TEMPLATE seguem no arquivo apenas como referência /
    # fallback manual; o fluxo de publicação normal não os chama mais, já que
    # o index.html publicado no Pages agora é estático e lê dashboard_data.json.