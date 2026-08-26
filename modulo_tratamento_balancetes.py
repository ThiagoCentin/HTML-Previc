"""
Modulo: Tratamento de Balancetes -- "Dados Extração" -> "Dados Tratados"
==========================================================================
Substitui o tratamento manual que era feito no Excel para os balancetes
brutos exportados da Previc (Consolidados, PGA e Planos). Para cada um dos
três, localiza a competência mais recente disponível na respectiva subpasta
de "Dados Extração", aplica o mesmo tratamento que era feito manualmente e
grava o resultado em "Dados Tratados", no formato já consumido pelo
dashboard (efpc_dashboard_gerador_v3.py).

O tratamento replica exatamente os passos manuais:
  1. Conversão para .xlsx.
  2. Remoção das linhas de cabeçalho do relatório (título/rodapé da Previc),
     mantendo só a linha de nomes de coluna e os dados.
  3. Separação da coluna única em colunas reais (o "Text to Columns" do
     Excel) -- aqui feita diretamente pelo parser de CSV, incluindo a
     conversão dos valores em formato BR ("1.234.567,89") para número.
  4. Cálculo do PL de cada EFPC via a fórmula
     =SUMPRODUCT(($B$2:$B$60000=B2)*($D$2:$D$60000=2030000000000)*$J$2:$J$60000),
     que soma o VL_SALDO_FINAL da conta de Patrimônio Social (2030000000000)
     por EFPC e propaga esse valor para todas as linhas da mesma EFPC.

O balancete de Planos, além disso, já traz nativamente do relatório as
colunas "Filtro Bloco" / "Filtro Hierarquia" / "Filtro Despesas" -- só que
o formato do arquivo mudou de uma competência para outra (ver
_tratar_planos), então a leitura é adaptada automaticamente ao layout de
cada arquivo em vez de presumir um único formato fixo.

Para o Consolidado (o único balancete de onde o dashboard lê essas 3
colunas), elas não vêm no relatório bruto -- são recalculadas aqui a partir
das fórmulas documentadas na aba "Sheet1" do arquivo tratado histórico:
  Filtro Bloco      = CHOOSE(LEFT(NUM_CONTA,1); Ativo/Passivo/Gestão Prev./
                       Gestão Admin./Fluxo Invest./Assistencial/Operações
                       Trans./Encerramento)
  Filtro Hierarquia = quantidade de dígitos não-zero em NUM_CONTA
  Filtro Despesas   = dentro de contas de Gestão, categoriza NM_CONTA por
                       palavra-chave (CONSELHO/AUDITORIA/TI/PESSOAL/
                       SERVIÇOS -> Outros)
"""

import glob
import os
import re
import shutil
from datetime import datetime

import numpy as np
import pandas as pd

CONTA_PL = 2030000000000

BLOCO_POR_DIGITO = {
    "1": "Ativo", "2": "Passivo", "3": "Gestão Prev.", "4": "Gestão Admin.",
    "5": "Fluxo Invest.", "6": "Assistencial", "7": "Operações Trans.", "8": "Encerramento",
}

MESES_PT = {
    "jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6,
    "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12,
}


def _brl_para_float(serie):
    """Converte string BR '1.234.567,89' -> float. Vazio/NaN -> 0.0."""
    s = serie.astype(str).str.strip()
    s = s.where(~s.isin(["", "nan", "None", "<NA>"]), "0")
    s = s.str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
    return pd.to_numeric(s, errors="coerce").fillna(0.0)


def _num_conta_para_int(serie):
    """
    Converte NUM_CONTA para inteiro, aceitando tanto o formato simples
    ('1000000000000') quanto o formato com separador de milhar visto num
    dos layouts do balancete de Planos ('1,000,000,000,000.00').
    """
    s = serie.astype(str).str.strip().str.replace(",", "", regex=False)
    return pd.to_numeric(s, errors="coerce").astype("Int64")


def _competencia_do_nome(nome_arquivo):
    """Extrai (ano, mes) do nome do arquivo bruto (padrão 'MM-YYYY' ou 'MM-YY')."""
    m = re.search(r"(\d{1,2})[-_](\d{2,4})(?!\d)", nome_arquivo)
    if not m:
        return None
    mes, ano = int(m.group(1)), int(m.group(2))
    if not (1 <= mes <= 12):
        return None
    if ano < 100:
        ano += 2000
    return (ano, mes)


def _todos_arquivos(pasta):
    """Retorna [((ano, mes), caminho), ...] de todo .csv com competência reconhecível
    na pasta, da competência mais antiga pra mais recente."""
    candidatos = []
    for caminho in glob.glob(os.path.join(pasta, "*.csv")):
        comp = _competencia_do_nome(os.path.basename(caminho))
        if comp is not None:
            candidatos.append((comp, caminho))
    candidatos.sort(key=lambda x: x[0])
    return candidatos


def _arquivo_mais_recente_tratado(pasta, prefixo):
    """Localiza, em `pasta` (Dados Tratados), o '{prefixo}*.xlsx' mais recente --
    mesmo critério (competência no nome, com fallback pra mtime) usado pelo gerador
    do dashboard pra escolher a base atual. Retorna None se não achar nada."""
    candidatos = glob.glob(os.path.join(pasta, f"{prefixo}*.xlsx"))
    if not candidatos:
        return None

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


def _mover_para_arquivo_morto(caminho, base_dir):
    """Move `caminho` pra 'arquivo_morto' (espelhando a subpasta de origem relativa a
    base_dir, carimbado com data/hora) -- nunca apaga um arquivo em silêncio, só tira
    ele do caminho de quem já foi processado."""
    rel = os.path.relpath(os.path.dirname(caminho), base_dir)
    pasta_morto = os.path.join(base_dir, "Arquivo_Morto", rel)
    os.makedirs(pasta_morto, exist_ok=True)
    stem, ext = os.path.splitext(os.path.basename(caminho))
    carimbo = datetime.now().strftime("%Y%m%d-%H%M%S")
    destino = os.path.join(pasta_morto, f"{stem}__{carimbo}{ext}")
    shutil.move(caminho, destino)
    return destino


def _filtro_bloco(num_conta):
    primeiro_digito = num_conta.astype(str).str[0]
    return primeiro_digito.map(BLOCO_POR_DIGITO).fillna("N/D")


def _filtro_hierarquia(num_conta):
    return num_conta.astype(str).str.replace("0", "", regex=False).str.len()


def _filtro_despesas(filtro_bloco, nm_conta):
    """Só classifica dentro de contas de Gestão (Prev./Admin.); fora disso, NaN."""
    nm_upper = nm_conta.astype(str).str.upper()
    eh_gestao = filtro_bloco.str.contains("Gestão", na=False)
    condicoes = [
        nm_upper.str.contains("CONSELHO", na=False),
        nm_upper.str.contains("AUDITORIA", na=False),
        nm_upper.str.contains("TI", na=False),
        nm_upper.str.contains("PESSOAL", na=False),
        nm_upper.str.contains("SERVIÇOS", na=False),
    ]
    categorias = ["Governança", "Auditoria", "Tecnologia", "Pessoal", "Serviços"]
    despesa = pd.Series(np.select(condicoes, categorias, default="Outros"),
                         index=nm_conta.index, dtype=object)
    despesa[~eh_gestao] = np.nan
    return despesa


def _tratar_padrao(caminho_csv, com_pl, com_filtros):
    """
    Trata um balancete no layout "padrão" (Consolidados / PGA):
    DATA_COMP,NU_MATRICULA_EFPC,SG_EFPC,NUM_CONTA,NM_CONTA,VL_SALDO_INICIAL,
    CS_NATUREZA_LANCAMENTO,VL_DEBITO,VL_CREDITO,VL_SALDO_FINAL,CS_NATUREZA_SALDO_FINAL
    -- sempre com 3 linhas de cabeçalho do relatório antes da linha de colunas.
    """
    df = pd.read_csv(caminho_csv, sep=",", skiprows=3, encoding="utf-8-sig", dtype=str)
    df["DATA_COMP"] = pd.to_datetime(df["DATA_COMP"], format="%m/%Y")
    df["NU_MATRICULA_EFPC"] = pd.to_numeric(df["NU_MATRICULA_EFPC"], errors="coerce").astype("Int64")
    df["NUM_CONTA"] = _num_conta_para_int(df["NUM_CONTA"])
    for col in ["VL_SALDO_INICIAL", "VL_DEBITO", "VL_CREDITO", "VL_SALDO_FINAL"]:
        df[col] = _brl_para_float(df[col])

    if com_filtros:
        df["Filtro Bloco"] = _filtro_bloco(df["NUM_CONTA"])
        df["Filtro Hierarquia"] = _filtro_hierarquia(df["NUM_CONTA"])
        df["Filtro Despesas"] = _filtro_despesas(df["Filtro Bloco"], df["NM_CONTA"])

    if com_pl:
        pl_por_efpc = (df[df["NUM_CONTA"] == CONTA_PL]
                       .groupby("NU_MATRICULA_EFPC")["VL_SALDO_FINAL"].sum())
        df["PL"] = df["NU_MATRICULA_EFPC"].map(pl_por_efpc).fillna(0.0)

    return df


def _detectar_layout_planos(caminho_csv):
    """
    O relatório de Planos já mudou de layout entre competências (ex.:
    10/2025 e 11/2025 vieram com 3 linhas de cabeçalho, separador "," e sem
    as colunas Filtro; 12/2025 veio sem cabeçalho extra, separador ";" e já
    com as colunas Filtro). Em vez de presumir um layout fixo, detecta pela
    própria primeira linha que de fato começa com "DATA_COMP".
    """
    with open(caminho_csv, encoding="utf-8-sig") as f:
        linhas = [f.readline() for _ in range(6)]
    n_pular = 0
    for linha in linhas:
        if linha.startswith("DATA_COMP"):
            break
        n_pular += 1
    else:
        raise ValueError(f"Não encontrei a linha de cabeçalho (DATA_COMP) em: {caminho_csv}")
    sep = ";" if ";" in linhas[n_pular] else ","
    return n_pular, sep


def _data_comp_planos(serie):
    """Aceita tanto 'MM/YYYY' quanto a abreviação em português 'mmm/aa' (ex: 'dez/25')."""
    s = serie.astype(str).str.strip()
    if s.str.match(r"^\d{1,2}/\d{4}$").all():
        return pd.to_datetime(s, format="%m/%Y")
    partes = s.str.lower().str.split("/", expand=True)
    mes = partes[0].map(MESES_PT)
    ano = pd.to_numeric(partes[1], errors="coerce")
    ano = ano.where(ano >= 100, ano + 2000)
    return pd.to_datetime({"year": ano, "month": mes, "day": 1})


def _tratar_planos(caminho_csv):
    """Trata o balancete de Planos, adaptando-se ao layout do arquivo (ver _detectar_layout_planos)."""
    n_pular, sep = _detectar_layout_planos(caminho_csv)
    df = pd.read_csv(caminho_csv, sep=sep, skiprows=n_pular, encoding="utf-8-sig", dtype=str)

    df["DATA_COMP"] = _data_comp_planos(df["DATA_COMP"])
    df["NU_CNPB"] = pd.to_numeric(df["NU_CNPB"], errors="coerce").astype("Int64")
    df["NUM_CONTA"] = _num_conta_para_int(df["NUM_CONTA"])
    for col in ["VL_SALDO_INICIAL", "VL_DEBITO", "VL_CREDITO", "VL_SALDO_FINAL"]:
        df[col] = _brl_para_float(df[col])
    # Essas 3 colunas (quando vêm no relatório) não são usadas na análise por
    # plano -- o dashboard só olha "Filtro Despesas"/"Filtro Hierarquia" do
    # Consolidado -- então ficam de fora do arquivo tratado, como já era.
    df = df.drop(columns=[c for c in ("Filtro Bloco", "Filtro Hierarquia", "Filtro Despesas") if c in df.columns])

    return df


def gerar_bases_tratadas(base_dir):
    """
    Trata TODOS os balancetes brutos (Consolidado/PGA/Planos) disponíveis em "Dados
    Extração" -- não só a competência mais recente -- e grava o .xlsx correspondente
    de cada um em "Dados Tratados". Depois de tratado com sucesso, o bruto sai de
    "Dados Extração" e vai pra "arquivo_morto" (nunca apagado, só arquivado) -- então
    uma vez tratado ele não volta a ser reprocessado nas próximas execuções, e "Dados
    Extração" só acumula o que ainda está pendente. Se dois brutos caírem na mesma
    competência (reenvio/correção), o .xlsx tratado antigo dessa competência também
    é arquivado antes do novo entrar no lugar.

    Retorna {"consolidado": caminho, "pga": caminho, "planos": caminho} só com o
    tratado da competência MAIS RECENTE de cada um -- é só isso que o gerador do
    dashboard consome como "base atual"; o histórico multi-competência é montado à
    parte, lendo tudo que sobrar em "Dados Tratados".

    Cada um dos três é checado de forma independente: se não houver nenhum .csv
    novo em "Dados Extração" pra um deles, essa etapa é pulada e o caminho
    retornado é o do .xlsx mais recente já existente em "Dados Tratados" -- só dá
    erro se não houver bruto novo NEM tratado existente pra usar.
    """
    extracao = os.path.join(base_dir, "Dados Extração")
    tratados = os.path.join(base_dir, "Dados Tratados")

    specs = [
        ("consolidado", os.path.join(extracao, "Balancetes Consolidados"),
         os.path.join(tratados, "Balancetes Consolidado Tratados"), "BALANCETES CONSOLIDADOS",
         lambda p: _tratar_padrao(p, com_pl=True, com_filtros=True)),
        ("pga", os.path.join(extracao, "Balancetes PGA"),
         os.path.join(tratados, "Balancetes PGAs Tratados"), "BALANCETES PGA",
         lambda p: _tratar_padrao(p, com_pl=False, com_filtros=False)),
        ("planos", os.path.join(extracao, "Balancetes de Planos"),
         os.path.join(tratados, "Balancetes Planos Tratados"), "BALANCETES PLANOS",
         _tratar_planos),
    ]

    print("Tratando balancetes (Dados Extração -> Dados Tratados)...")
    resultado = {}
    for chave, pasta_origem, pasta_destino, prefixo, tratar in specs:
        candidatos = _todos_arquivos(pasta_origem)
        os.makedirs(pasta_destino, exist_ok=True)

        if not candidatos:
            existente = _arquivo_mais_recente_tratado(pasta_destino, prefixo)
            if existente is None:
                raise FileNotFoundError(
                    f"Nenhum .csv com competência reconhecível em '{pasta_origem}' "
                    f"e nenhum tratado existente em '{pasta_destino}'."
                )
            print(f"  → {prefixo}: nenhum dado novo em Dados Extração, mantendo tratado existente "
                  f"({os.path.basename(existente)})")
            resultado[chave] = existente
            continue

        caminho_saida_mais_recente = None
        for (ano, mes), caminho_bruto in candidatos:
            sufixo = f"{mes:02d}-{ano % 100:02d}"
            print(f"  → {prefixo}: competência {mes:02d}/{ano} ({os.path.basename(caminho_bruto)})")
            try:
                df = tratar(caminho_bruto)
            except Exception as e:
                print(f"    ⚠ falhou ao tratar, deixado em Dados Extração pra checar: {e}")
                continue

            caminho_saida = os.path.join(pasta_destino, f"{prefixo}-{sufixo}.xlsx")
            if os.path.exists(caminho_saida):
                arquivado = _mover_para_arquivo_morto(caminho_saida, base_dir)
                print(f"    ↪ tratado anterior dessa competência arquivado em: {arquivado}")
            df.to_excel(caminho_saida, sheet_name=f"{prefixo} - {sufixo}"[:31], index=False)
            print(f"    ✓ {len(df):,} linhas gravadas em: {caminho_saida}")

            bruto_arquivado = _mover_para_arquivo_morto(caminho_bruto, base_dir)
            print(f"    ↪ bruto original arquivado em: {bruto_arquivado}")
            caminho_saida_mais_recente = caminho_saida

        if caminho_saida_mais_recente is None:
            raise RuntimeError(f"Nenhum balancete de '{prefixo}' foi tratado com sucesso em: {pasta_origem}")
        resultado[chave] = caminho_saida_mais_recente

    print()
    return resultado
