import pandas as pd
import xlwings as xw

# Single source of truth for option product types
PRODUTOS_OPCAO = frozenset(["Opcao", "Opção de ações", "Opcao de açoes", "Option"])

PRODUTOS_SWAP = frozenset(["Swap", "Swap - Generic"])

PREFIXOS_ESPECIAIS = ("OFEQ", "EQC", "EQV", "INDV", "INDC")


def limpar_texto(valor) -> str:
    try:
        if pd.isna(valor):
            return ""
    except (TypeError, ValueError):
        pass
    return str(valor).replace("\xa0", " ").strip()


def _tem_valor(v) -> bool:
    return v is not None and limpar_texto(v) != ""


def tratar_trade_id(valor) -> str:
    valor = limpar_texto(valor)
    if "-" in valor:
        return valor.split("-")[0].strip()
    return valor


def montar_codigo_base(linha) -> str:
    produto = limpar_texto(linha["Product"])
    asset = limpar_texto(linha["Asset"])

    if produto == "Ações":
        return asset + " BZ EQUITY"
    elif produto == "Equity":
        return asset + " EQUITY"
    elif produto in PRODUTOS_OPCAO:
        return asset if produto != "Option" else asset + " Equity"

    return asset


def comeca_com_prefixo_especial(codigo: str) -> bool:
    return limpar_texto(codigo).startswith(PREFIXOS_ESPECIAIS)


def criar_mapas_fts(df_fts: pd.DataFrame):
    fts_c_para_i = (
        df_fts.dropna(subset=["C"])
        .assign(C_limpo=lambda x: x["C"].astype(str).str.replace("\xa0", " ", regex=False).str.strip())
        .drop_duplicates(subset=["C_limpo"])
        .set_index("C_limpo")["I"]
        .to_dict()
    )

    fts_br_para_i = (
        df_fts.dropna(subset=["BR"])
        .assign(BR_limpo=lambda x: x["BR"].astype(str).str.replace("\xa0", " ", regex=False).str.strip())
        .drop_duplicates(subset=["BR_limpo"])
        .set_index("BR_limpo")["I"]
        .to_dict()
    )

    return fts_c_para_i, fts_br_para_i


def carregar_mapa_aux(sht_aux) -> dict:
    """Reads H:I range from AUX sheet once and returns a {H_value: I_value} dict."""
    ultima_linha = sht_aux.range("H" + str(sht_aux.cells.last_cell.row)).end("up").row
    dados_hi = sht_aux.range(f"H1:I{ultima_linha}").value

    if not dados_hi:
        return {}

    # xlwings returns a flat list for a single row; wrap it
    if isinstance(dados_hi[0], (str, int, float)) or dados_hi[0] is None:
        dados_hi = [dados_hi]

    mapa = {}
    for linha in dados_hi:
        if linha is None:
            continue
        valor_h = linha[0] if len(linha) > 0 else None
        valor_i = linha[1] if len(linha) > 1 else None
        chave = limpar_texto(valor_h)
        if chave:
            mapa[chave] = valor_i

    return mapa


def traduzir_asset_name(
    linha,
    fts_c_para_i: dict,
    fts_br_para_i: dict,
    mapa_aux: dict,
    nome_coluna_swap: str = "AN",
) -> str:
    produto = limpar_texto(linha["Product"])
    codigo_base = montar_codigo_base(linha)

    # Options: try FTS lookups in priority order
    if produto in PRODUTOS_OPCAO:
        chave_rating = limpar_texto(linha["Rating Description"])
        if chave_rating in fts_c_para_i:
            return fts_c_para_i[chave_rating]

        chave_trade_id = tratar_trade_id(linha["Trade ID"])
        if chave_trade_id in fts_br_para_i:
            return fts_br_para_i[chave_trade_id]

        chave_asset = limpar_texto(linha["Asset"])
        if chave_asset in fts_c_para_i:
            return fts_c_para_i[chave_asset]

        return codigo_base

    # Swap: look up in pre-built AUX dict
    if produto in PRODUTOS_SWAP:
        chave_swap = limpar_texto(linha.get(nome_coluna_swap, ""))
        resultado = mapa_aux.get(chave_swap)
        if _tem_valor(resultado):
            return resultado
        return codigo_base

    # Special prefixes: try FTS C lookup
    if comeca_com_prefixo_especial(codigo_base):
        chave_base = limpar_texto(codigo_base)
        if chave_base in fts_c_para_i:
            return fts_c_para_i[chave_base]
        return codigo_base

    return codigo_base


def atualizar_asset_name():
    global df_mars, df_fts

    wb = xw.Book.caller()
    sht_mars = wb.sheets["MARS"]
    sht_aux = wb.sheets["AUX"]

    fts_c_para_i, fts_br_para_i = criar_mapas_fts(df_fts)
    mapa_aux = carregar_mapa_aux(sht_aux)

    df_mars["Asset Name"] = df_mars.apply(
        lambda linha: traduzir_asset_name(
            linha=linha,
            fts_c_para_i=fts_c_para_i,
            fts_br_para_i=fts_br_para_i,
            mapa_aux=mapa_aux,
            nome_coluna_swap="AN",
        ),
        axis=1,
    )

    col_asset_name = df_mars.columns.get_loc("Asset Name") + 1
    sht_mars.range((2, col_asset_name)).options(transpose=True).value = df_mars["Asset Name"].tolist()

    print("Asset Name atualizado com sucesso.")
