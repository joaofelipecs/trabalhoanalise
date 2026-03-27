import pandas as pd
import xlwings as xw

# =============================================================================
# CONFIGURAÇÕES — edite aqui se mudar nomes de abas ou colunas
# =============================================================================

ABA_MARS = "MARS"          # Aba com os dados principais
ABA_FTS  = "FTS"           # Aba com a tabela de tradução (colunas C, BR, I)
ABA_AUX  = "AUX"           # Aba auxiliar com mapa de Swap (colunas H e I)

COLUNA_SWAP = "AN"         # Coluna em MARS usada como chave para Swap

# Linha onde começa o cabeçalho em cada aba (1 = primeira linha)
LINHA_CABECALHO_MARS = 1
LINHA_CABECALHO_FTS  = 1

# =============================================================================
# TIPOS DE PRODUTO — edite aqui se surgir novo nome de produto
# =============================================================================

PRODUTOS_OPCAO = frozenset([
    "Opcao",
    "Opção de ações",
    "Opcao de açoes",
    "Option",
])

PRODUTOS_SWAP = frozenset([
    "Swap",
    "Swap - Generic",
])

# Códigos que começam com esses prefixos recebem tratamento especial via FTS
PREFIXOS_ESPECIAIS = ("OFEQ", "EQC", "EQV", "INDV", "INDC")


# =============================================================================
# FUNÇÕES AUXILIARES
# =============================================================================

def limpar_texto(valor) -> str:
    """Remove espaços, &nbsp; e converte para string. Retorna '' se vazio."""
    try:
        if pd.isna(valor):
            return ""
    except (TypeError, ValueError):
        pass
    return str(valor).replace("\xa0", " ").strip()


def tem_valor(v) -> bool:
    """Retorna True se v não for None nem string vazia."""
    return v is not None and limpar_texto(v) != ""


def tratar_trade_id(valor) -> str:
    """Extrai a parte antes do primeiro '-' no Trade ID."""
    valor = limpar_texto(valor)
    if "-" in valor:
        return valor.split("-")[0].strip()
    return valor


# =============================================================================
# LEITURA DOS DADOS DO EXCEL
# =============================================================================

def ler_dataframe_da_aba(aba) -> pd.DataFrame:
    """Lê a aba inteira como DataFrame (cabeçalho na primeira linha)."""
    return (
        aba
        .range("A1")
        .options(pd.DataFrame, expand="table", header=1)
        .value
    )


def ler_mapa_aux(sht_aux) -> dict:
    """
    Lê as colunas H e I da aba AUX e devolve um dicionário {H: I}.
    Chamado UMA vez antes do loop — não lê o Excel linha por linha.
    """
    ultima_linha = sht_aux.range("H1").end("down").row
    dados = sht_aux.range(f"H1:I{ultima_linha}").value

    if not dados:
        return {}

    # xlwings retorna lista simples quando há só 1 linha; normaliza para lista de listas
    if isinstance(dados[0], (str, int, float)) or dados[0] is None:
        dados = [dados]

    mapa = {}
    for linha in dados:
        if linha is None:
            continue
        chave  = limpar_texto(linha[0]) if len(linha) > 0 else ""
        valor  = linha[1]               if len(linha) > 1 else None
        if chave:
            mapa[chave] = valor

    return mapa


def criar_mapas_fts(df_fts: pd.DataFrame):
    """
    Cria dois dicionários a partir da aba FTS:
      - fts_c_para_i  : coluna C  → coluna I
      - fts_br_para_i : coluna BR → coluna I
    """
    def montar_mapa(coluna: str) -> dict:
        col_limpa = coluna + "_limpo"
        return (
            df_fts
            .dropna(subset=[coluna])
            .assign(**{
                col_limpa: lambda x: (
                    x[coluna].astype(str)
                    .str.replace("\xa0", " ", regex=False)
                    .str.strip()
                )
            })
            .drop_duplicates(subset=[col_limpa])
            .set_index(col_limpa)["I"]
            .to_dict()
        )

    fts_c_para_i  = montar_mapa("C")
    fts_br_para_i = montar_mapa("BR")

    return fts_c_para_i, fts_br_para_i


# =============================================================================
# LÓGICA DE NEGÓCIO: montar código-base e traduzir Asset Name
# =============================================================================

def montar_codigo_base(linha) -> str:
    """
    Monta o código Bloomberg/padrão da linha conforme o tipo de produto.
    Este valor é usado como fallback quando nenhuma tradução é encontrada.
    """
    produto = limpar_texto(linha["Product"])
    asset   = limpar_texto(linha["Asset"])

    if produto == "Ações":
        return asset + " BZ EQUITY"

    if produto == "Equity":
        return asset + " EQUITY"

    if produto == "Option":
        return asset + " Equity"

    if produto in PRODUTOS_OPCAO:   # demais variantes de opção brasileira
        return asset

    return asset


def codigo_comeca_com_prefixo_especial(codigo: str) -> bool:
    return limpar_texto(codigo).startswith(PREFIXOS_ESPECIAIS)


def traduzir_asset_name(
    linha,
    fts_c_para_i:  dict,
    fts_br_para_i: dict,
    mapa_aux:      dict,
) -> str:
    """
    Decide o Asset Name de uma linha do MARS seguindo as prioridades:

    Opções
      1. Rating Description → busca em FTS (coluna C)
      2. Trade ID (antes do '-') → busca em FTS (coluna BR)
      3. Asset → busca em FTS (coluna C)
      4. Fallback: código-base

    Swap
      1. Coluna AN → busca no mapa da aba AUX (H→I)
      2. Fallback: código-base

    Outros com prefixo especial (OFEQ, EQC…)
      1. Código-base → busca em FTS (coluna C)
      2. Fallback: código-base

    Todos os demais
      → código-base diretamente
    """
    produto      = limpar_texto(linha["Product"])
    codigo_base  = montar_codigo_base(linha)

    # --- Opções ---------------------------------------------------------------
    if produto in PRODUTOS_OPCAO:

        # Prioridade 1: Rating Description → FTS C
        chave = limpar_texto(linha["Rating Description"])
        if chave in fts_c_para_i:
            return fts_c_para_i[chave]

        # Prioridade 2: Trade ID → FTS BR
        chave = tratar_trade_id(linha["Trade ID"])
        if chave in fts_br_para_i:
            return fts_br_para_i[chave]

        # Prioridade 3: Asset → FTS C
        chave = limpar_texto(linha["Asset"])
        if chave in fts_c_para_i:
            return fts_c_para_i[chave]

        return codigo_base  # fallback

    # --- Swap -----------------------------------------------------------------
    if produto in PRODUTOS_SWAP:
        chave     = limpar_texto(linha.get(COLUNA_SWAP, ""))
        resultado = mapa_aux.get(chave)
        if tem_valor(resultado):
            return resultado
        return codigo_base  # fallback

    # --- Prefixos especiais ---------------------------------------------------
    if codigo_comeca_com_prefixo_especial(codigo_base):
        chave = limpar_texto(codigo_base)
        if chave in fts_c_para_i:
            return fts_c_para_i[chave]
        return codigo_base  # fallback

    # --- Todos os demais ------------------------------------------------------
    return codigo_base


# =============================================================================
# PONTO DE ENTRADA — chamado pelo VBA via RunPython
# =============================================================================

def atualizar_asset_name():
    """
    Função principal chamada pelo VBA.
    Lê os dados do Excel, calcula o Asset Name e grava de volta na aba MARS.
    """
    wb = xw.Book.caller()

    # 1. Abrir as abas
    sht_mars = wb.sheets[ABA_MARS]
    sht_fts  = wb.sheets[ABA_FTS]
    sht_aux  = wb.sheets[ABA_AUX]

    # 2. Ler os dados
    df_mars  = ler_dataframe_da_aba(sht_mars)
    df_fts   = ler_dataframe_da_aba(sht_fts)
    mapa_aux = ler_mapa_aux(sht_aux)

    # 3. Montar os dicionários de tradução da FTS
    fts_c_para_i, fts_br_para_i = criar_mapas_fts(df_fts)

    # 4. Calcular o Asset Name linha a linha
    df_mars["Asset Name"] = df_mars.apply(
        lambda linha: traduzir_asset_name(
            linha          = linha,
            fts_c_para_i   = fts_c_para_i,
            fts_br_para_i  = fts_br_para_i,
            mapa_aux       = mapa_aux,
        ),
        axis=1,
    )

    # 5. Gravar a coluna de volta no Excel (a partir da linha 2, abaixo do cabeçalho)
    col_idx = df_mars.columns.get_loc("Asset Name") + 1
    sht_mars.range((2, col_idx)).options(transpose=True).value = df_mars["Asset Name"].tolist()

    print("✓ Asset Name atualizado com sucesso.")
