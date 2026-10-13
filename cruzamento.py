

import json
import unicodedata
from pathlib import Path

import pandas as pd


# ---------------------------------------------------------------------------
# Configuração
# ---------------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
SAIDA_DIR = BASE_DIR / "saida"

ARQ_MATERIAIS = BASE_DIR / "materiais .xlsx"
ARQ_CONFERENCIA = BASE_DIR / "conferência de produtos.xlsx"
ARQ_REGISTROS = BASE_DIR / "registros.xlsx"

CHAVE = "Material"      # código do produto, presente nas três fontes
SHEET = 0               # todas as fontes usam a primeira aba ('Exportação SAPUI5')


# ---------------------------------------------------------------------------
# 1. EXTRACT
# ---------------------------------------------------------------------------
def ler_excel(caminho: Path) -> pd.DataFrame:
    if not caminho.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {caminho}")
    print(f"  Lendo: {caminho.name}")
    return pd.read_excel(caminho, sheet_name=SHEET, engine="openpyxl")


# ---------------------------------------------------------------------------
# 2. TRANSFORM
# ---------------------------------------------------------------------------
def preparar_materiais(df: pd.DataFrame) -> pd.DataFrame:
    """Cadastro mestre: define o grão da tabela final (1 linha por material)."""
    df = df.copy()
    df[CHAVE] = df[CHAVE].astype("int64")
    # Garante unicidade da chave
    duplicados = df[CHAVE].duplicated().sum()
    if duplicados:
        print(f"  [aviso] {duplicados} material(is) duplicado(s) no cadastro; "
              f"mantendo o primeiro.")
        df = df.drop_duplicates(subset=CHAVE, keep="first")
    return df


def agregar_conferencia(df: pd.DataFrame) -> pd.DataFrame:
    """
    A conferência tem várias linhas por material (uma por unidade de medida).
    Agregamos ao nível de material para não multiplicar as linhas no merge.
    """
    df = df.copy()
    df[CHAVE] = df[CHAVE].astype("int64")

    agg = df.groupby(CHAVE).agg(
        qtd_unidades_medida=("UM alternativa", "nunique"),
        unidades_medida=("UM alternativa",
                         lambda s: ", ".join(sorted(s.dropna().astype(str).unique()))),
        peso_liquido_max=("Peso líquido", "max"),
        peso_bruto_max=("Peso bruto", "max"),
        volume_max=("Volume", "max"),
        tem_ean=("Código EAN/UPC", lambda s: s.notna().any()),
    ).reset_index()

    # Atributos que são constantes por material (pega o primeiro valor).
    # 'UMB' aqui vira a unidade de venda base vinda da conferência.
    const_cols = ["Hierarq.produtos", "Denominação", "Denominação.1",
                  "Denominação.2", "UMB"]
    const_cols = [c for c in const_cols if c in df.columns]
    if const_cols:
        atributos = df.groupby(CHAVE)[const_cols].first().reset_index()
        # evita colidir com a 'UMB' do cadastro de materiais no merge
        atributos = atributos.rename(columns={"UMB": "umb_conferencia"})
        agg = agg.merge(atributos, on=CHAVE, how="left")

    return agg


def preparar_registros(df: pd.DataFrame) -> pd.DataFrame:
    """Lista de preços: 1 linha por material (preço sugerido / mín / máx)."""
    df = df.copy()
    df[CHAVE] = df[CHAVE].astype("int64")

    colunas = {
        "Família": "familia",
        "Marca": "marca",
        "Lista de Preço": "lista_preco",
        "Desc. Lista de Preço": "desc_lista_preco",
        "Preço Sugerido": "preco_sugerido",
        "Mínimo": "preco_minimo",
        "Máximo": "preco_maximo",
        "Moeda": "moeda",
        "Válido desde": "preco_valido_desde",
    }
    colunas = {k: v for k, v in colunas.items() if k in df.columns}
    df = df[[CHAVE, *colunas.keys()]].rename(columns=colunas)

    # Se houver mais de um preço por material, fica o mais recente
    if "preco_valido_desde" in df.columns:
        df = (df.sort_values("preco_valido_desde")
                .drop_duplicates(subset=CHAVE, keep="last"))
    else:
        df = df.drop_duplicates(subset=CHAVE, keep="last")
    return df


# ---------------------------------------------------------------------------
# 3. LOAD / MERGE
# ---------------------------------------------------------------------------
def cruzar(materiais, conferencia, registros) -> pd.DataFrame:
    resultado = (
        materiais
        .merge(conferencia, on=CHAVE, how="left")
        .merge(registros, on=CHAVE, how="left")
    )

    # Flags de qualidade: o material existe em cada fonte?
    resultado["tem_conferencia"] = resultado["qtd_unidades_medida"].notna()
    resultado["tem_preco"] = resultado["preco_sugerido"].notna()
    return resultado


# ---------------------------------------------------------------------------
# 3b. ENRIQUECIMENTO - preço de ATACADO (BRL) para itens sem precificação
# ---------------------------------------------------------------------------
# ATENÇÃO: preços de ATACADO em BRL (por CAIXA fechada), NÃO comparáveis
# diretamente com a lista de exportação EXP-MINOIL (USD). Pesquisados na web
# em out/2026. Uso indicativo apenas, não oficial.
#
# Qualidade dos dados:
#   - 12037417 (filtro): preço de atacado REAL por caixa (fonte Panapilhas,
#     proxy nº 103/30 em caixa 40x30un, mesmo formato da base).
#   - 12034109 e 20908885: SEM cotação de atacado publicada -> ESTIMADO a
#     partir do preço de varejo unitário:
#         preco_atacado = preco_varejo_un * unid_por_caixa * (1 - desconto)
#     com desconto de atacado = 20%. unid_por_caixa vem da descrição do produto.
DESCONTO_ATACADO = 0.20  # desconto típico de atacado sobre o varejo unitário

# Taxa de câmbio para converter o preço de atacado BRL -> USD.
# Referência comercial/BCB de out/2026 (~R$5,00/USD). Ajuste conforme a data.
TAXA_CAMBIO_BRL_USD = 5.00

# Material: (preco_atacado_brl, metodo, fonte)
PRECOS_ATACADO = {
    12037417: (152.05, "cotado",
               "Panapilhas atacado - Filtro 3C 103/30 cx 40x30un (R$152,05)"),
    # Estimados: varejo_un * unid/caixa * (1 - 0,20)
    12034109: (round(26.14 * 24 * (1 - DESCONTO_ATACADO), 2), "estimado",
               "varejo R$26,14 x 24 potes/cx x 0,8 (Empório Valleju; Rappi)"),
    20908885: (round(11.30 * 12 * (1 - DESCONTO_ATACADO), 2), "estimado",
               "varejo R$11,30 x 12 un/cx x 0,8 (proxy 102 Mercado Livre; Amazon)"),
}


def enriquecer_preco_atacado(df: pd.DataFrame) -> pd.DataFrame:
    """
    Adiciona colunas de preço de ATACADO (BRL, por caixa) para os materiais
    sem precificação de exportação. NÃO sobrescreve o preço original.
    Distingue valores cotados de estimados via coluna 'preco_atacado_metodo'.
    """
    df = df.copy()
    df["preco_atacado_brl"] = df[CHAVE].map(
        lambda m: PRECOS_ATACADO[m][0] if m in PRECOS_ATACADO else None)
    df["preco_atacado_moeda"] = df[CHAVE].map(
        lambda m: "BRL" if m in PRECOS_ATACADO else None)
    df["preco_atacado_metodo"] = df[CHAVE].map(
        lambda m: PRECOS_ATACADO[m][1] if m in PRECOS_ATACADO else None)
    df["preco_atacado_fonte"] = df[CHAVE].map(
        lambda m: PRECOS_ATACADO[m][2] if m in PRECOS_ATACADO else None)
    df["preco_atacado_estimado"] = df[CHAVE].map(
        lambda m: PRECOS_ATACADO[m][1] == "estimado" if m in PRECOS_ATACADO else False)

    # Conversão BRL -> USD (preço de atacado em dólar)
    df["preco_atacado_usd"] = (df["preco_atacado_brl"] / TAXA_CAMBIO_BRL_USD).round(2)
    df["taxa_cambio_brl_usd"] = df[CHAVE].map(
        lambda m: TAXA_CAMBIO_BRL_USD if m in PRECOS_ATACADO else None)

    n = int(df["preco_atacado_brl"].notna().sum())
    n_est = int(df["preco_atacado_estimado"].sum())
    print(f"  preço de atacado (BRL/caixa) incluído para {n} material(is) "
          f"sem precificação ({n_est} estimado[s], {n - n_est} cotado[s])")
    print(f"  conversão USD aplicada à taxa de R${TAXA_CAMBIO_BRL_USD:.2f}/USD")
    return df


# ---------------------------------------------------------------------------
# 3c. PREÇO DE VENDA (plataforma de pedidos, moeda = USD)
# ---------------------------------------------------------------------------
# Regra de negócio definida:
#   - Moeda da plataforma: USD (todos os produtos).
#   - Preço de venda = preco_sugerido (lista de exportação, já em USD) quando
#     existir; senão o preço de atacado estimado convertido para USD.
#   - Todos os produtos ficam disponíveis para compra.
MOEDA_PLATAFORMA = "USD"


def definir_preco_venda(df: pd.DataFrame) -> pd.DataFrame:
    """Consolida um único preço de venda em USD por produto e marca a origem."""
    df = df.copy()

    # preço de venda: lista (preco_sugerido) tem prioridade; fallback = atacado USD
    df["preco_venda_usd"] = df["preco_sugerido"].where(
        df["preco_sugerido"].notna(), df["preco_atacado_usd"])
    df["preco_venda_moeda"] = MOEDA_PLATAFORMA

    # origem do preço: 'lista' (oficial) ou 'estimado' (atacado convertido)
    df["preco_origem"] = df["preco_sugerido"].notna().map(
        {True: "lista", False: "estimado"})
    # onde não houver nenhum dos dois, não é vendável
    df.loc[df["preco_venda_usd"].isna(), "preco_origem"] = "indefinido"

    df["disponivel_para_venda"] = df["preco_venda_usd"].notna()

    n_lista = int((df["preco_origem"] == "lista").sum())
    n_est = int((df["preco_origem"] == "estimado").sum())
    n_indef = int((df["preco_origem"] == "indefinido").sum())
    print(f"  preço de venda (USD): {n_lista} de lista, {n_est} estimado, "
          f"{n_indef} indefinido | disponíveis: "
          f"{int(df['disponivel_para_venda'].sum())}/{len(df)}")
    return df


def completar_catalogo(df: pd.DataFrame) -> pd.DataFrame:
    """
    Preenche campos de vitrine que ficam vazios para itens sem lista de preço:
      - unidade_venda: UMB do cadastro -> UMB da conferência -> 'UN'
      - familia: 'familia' da lista -> 'Denominação' da conferência
      - marca:   'marca' da lista -> inferida do nome (3C / Iguaçu)
    """
    df = df.copy()

    # unidade de venda (fallback encadeado)
    umb_cad = df["UMB"] if "UMB" in df.columns else pd.Series([None] * len(df))
    umb_conf = df["umb_conferencia"] if "umb_conferencia" in df.columns \
        else pd.Series([None] * len(df))
    df["unidade_venda"] = (
        umb_cad.where(umb_cad.notna() & (umb_cad.astype(str) != "nan"))
               .fillna(umb_conf)
               .fillna("UN")
    )

    # familia: usa a Denominação da conferência quando a família da lista falta
    if "Denominação" in df.columns:
        df["familia_catalogo"] = df["familia"].where(
            df["familia"].notna(), df["Denominação"])
    else:
        df["familia_catalogo"] = df["familia"]

    # marca: infere do nome do produto quando a marca da lista falta
    def _infere_marca(row):
        if pd.notna(row.get("marca")):
            return row["marca"]
        nome = str(row.get("Descrição", "")).upper()
        if "IGUA" in nome:
            return "IGUAÇU"
        if "3C" in nome:
            return "3 CORAÇÕES"
        return None

    df["marca_catalogo"] = df.apply(_infere_marca, axis=1)

    n_fam = int(df["familia_catalogo"].notna().sum())
    n_marca = int(df["marca_catalogo"].notna().sum())
    n_un = int(df["unidade_venda"].notna().sum())
    print(f"  catálogo completado: unidade_venda {n_un}/{len(df)}, "
          f"familia {n_fam}/{len(df)}, marca {n_marca}/{len(df)}")
    return df


# ---------------------------------------------------------------------------
# 4. EXPORT
# ---------------------------------------------------------------------------
def _slug(nome: str) -> str:
    """Normaliza nome de coluna: minúsculo, sem acento, snake_case."""
    nome = str(nome).strip().lower()
    nome = unicodedata.normalize("NFKD", nome).encode("ascii", "ignore").decode()
    out = []
    for ch in nome:
        out.append(ch if ch.isalnum() else "_")
    slug = "".join(out)
    while "__" in slug:
        slug = slug.replace("__", "_")
    return slug.strip("_")


def exportar_json_catalogo(df: pd.DataFrame, nome: str) -> None:
    """
    catalogo.json -> o que o CLIENTE vê na plataforma.
    Só produtos disponíveis, só campos de vitrine, sem dado interno.
    """
    SAIDA_DIR.mkdir(parents=True, exist_ok=True)
    caminho = SAIDA_DIR / f"{nome}.json"

    vis = df[df["disponivel_para_venda"]].copy()
    itens = []
    for _, r in vis.iterrows():
        fam = r.get("familia_catalogo")
        marca = r.get("marca_catalogo")
        un = r.get("unidade_venda")
        itens.append({
            "material": int(r["Material"]),
            "nome": str(r["Descrição"]),
            "familia": None if pd.isna(fam) else str(fam),
            "marca": None if pd.isna(marca) else str(marca),
            "unidade_venda": None if pd.isna(un) else str(un),
            "preco_venda": round(float(r["preco_venda_usd"]), 2),
            "moeda": str(r["preco_venda_moeda"]),
            "disponivel": True,
        })

    doc = {
        "schema": "catalogo",
        "versao": 1,
        "moeda": MOEDA_PLATAFORMA,
        "gerado_em": pd.Timestamp.now().isoformat(timespec="seconds"),
        "total_itens": len(itens),
        "itens": itens,
    }
    caminho.write_text(json.dumps(doc, ensure_ascii=False, indent=2),
                       encoding="utf-8")
    print(f"  catálogo (cliente): {len(itens)} itens -> {caminho}")


def exportar_json_interno(df: pd.DataFrame, nome: str) -> None:
    """
    produtos_internos.json -> o que o VENDEDOR/backoffice vê.
    Todos os campos, com rastreabilidade de preço e flags de qualidade.
    Nomes de coluna normalizados (snake_case, sem acento) = contrato estável.
    """
    SAIDA_DIR.mkdir(parents=True, exist_ok=True)
    caminho = SAIDA_DIR / f"{nome}.json"

    norm = df.rename(columns={c: _slug(c) for c in df.columns})
    # pandas/numpy -> tipos nativos para JSON; NaN -> None
    registros = json.loads(
        norm.to_json(orient="records", force_ascii=False, date_format="iso"))

    doc = {
        "schema": "produtos_internos",
        "versao": 1,
        "moeda_plataforma": MOEDA_PLATAFORMA,
        "taxa_cambio_brl_usd": TAXA_CAMBIO_BRL_USD,
        "gerado_em": pd.Timestamp.now().isoformat(timespec="seconds"),
        "total_itens": len(registros),
        "itens": registros,
    }
    caminho.write_text(json.dumps(doc, ensure_ascii=False, indent=2),
                       encoding="utf-8")
    print(f"  base interna (vendedor): {len(registros)} itens -> {caminho}")


def _escrever_xlsx(df_ou_func, xlsx: Path) -> bool:
    """Escreve o Excel tratando o caso de o arquivo estar aberto (lock do Windows)."""
    try:
        df_ou_func()
        return True
    except PermissionError:
        print(f"  [ERRO] '{xlsx.name}' está aberto (provavelmente no Excel). "
              f"Feche-o e rode de novo. Pulando este arquivo.")
        return False


def exportar(df: pd.DataFrame, nome: str) -> None:
    SAIDA_DIR.mkdir(parents=True, exist_ok=True)
    xlsx = SAIDA_DIR / f"{nome}.xlsx"
    csv = SAIDA_DIR / f"{nome}.csv"
    df.to_csv(csv, index=False, sep=";", encoding="utf-8-sig")
    _escrever_xlsx(lambda: df.to_excel(xlsx, index=False, engine="openpyxl"), xlsx)
    print(f"\nResultado ({len(df)} linhas, {len(df.columns)} colunas) salvo em:")
    print(f"  {xlsx}")
    print(f"  {csv}")


def gerar_relatorio(df: pd.DataFrame, nome: str) -> None:
    """
    Gera um Excel analítico com várias abas:
      - Resultado      : tabela completa do cruzamento
      - Itens_sem_preco: materiais conferidos mas sem lista de preço
      - Resumo_familia : contagem e preço médio/mín/máx por família
    """
    SAIDA_DIR.mkdir(parents=True, exist_ok=True)
    xlsx = SAIDA_DIR / f"{nome}.xlsx"

    cols_sem_preco = ["Material", "Descrição", "Denominação", "Centro"]
    cols_atacado = ["preco_atacado_brl", "preco_atacado_usd", "taxa_cambio_brl_usd",
                    "preco_atacado_metodo", "preco_atacado_fonte"]
    cols_sem_preco += [c for c in cols_atacado if c in df.columns]
    sem_preco = df.loc[~df["tem_preco"], cols_sem_preco]

    resumo = (
        df.assign(familia=df["familia"].fillna("(sem preço/família)"))
          .groupby("familia")
          .agg(qtd_materiais=("Material", "count"),
               com_preco=("tem_preco", "sum"),
               preco_medio=("preco_sugerido", "mean"),
               preco_min=("preco_sugerido", "min"),
               preco_max=("preco_sugerido", "max"))
          .round(2)
          .reset_index()
          .sort_values("qtd_materiais", ascending=False)
    )

    def _escrever():
        with pd.ExcelWriter(xlsx, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Resultado", index=False)
            sem_preco.to_excel(writer, sheet_name="Itens_sem_preco", index=False)
            resumo.to_excel(writer, sheet_name="Resumo_familia", index=False)

    if _escrever_xlsx(_escrever, xlsx):
        print(f"\nRelatório analítico salvo em:\n  {xlsx}")
        print(f"  abas: Resultado ({len(df)}), "
              f"Itens_sem_preco ({len(sem_preco)}), "
              f"Resumo_familia ({len(resumo)})")


# ---------------------------------------------------------------------------
# Orquestração
# ---------------------------------------------------------------------------
def main() -> None:
    print("1. EXTRACT")
    df_materiais = ler_excel(ARQ_MATERIAIS)
    df_conferencia = ler_excel(ARQ_CONFERENCIA)
    df_registros = ler_excel(ARQ_REGISTROS)

    print("\n2. TRANSFORM")
    materiais = preparar_materiais(df_materiais)
    conferencia = agregar_conferencia(df_conferencia)
    registros = preparar_registros(df_registros)
    print(f"  materiais:   {materiais.shape}")
    print(f"  conferencia: {conferencia.shape} (agregada por material)")
    print(f"  registros:   {registros.shape}")

    print("\n3. MERGE (base = cadastro de materiais, chave = 'Material')")
    resultado = cruzar(materiais, conferencia, registros)

    # Resumo de cobertura
    total = len(resultado)
    com_conf = int(resultado["tem_conferencia"].sum())
    com_preco = int(resultado["tem_preco"].sum())
    print(f"  {total} materiais | com conferência: {com_conf} | com preço: {com_preco}")

    print("\n3b. ENRIQUECIMENTO (preço de atacado p/ itens sem precificação)")
    resultado = enriquecer_preco_atacado(resultado)

    print("\n3c. PREÇO DE VENDA (plataforma, moeda USD)")
    resultado = definir_preco_venda(resultado)
    resultado = completar_catalogo(resultado)

    print("\n4. EXPORT")
    exportar(resultado, "cruzamento_resultado")
    gerar_relatorio(resultado, "cruzamento_relatorio")

    print("\n5. SAÍDA PARA A PLATAFORMA (JSON)")
    exportar_json_catalogo(resultado, "catalogo")
    exportar_json_interno(resultado, "produtos_internos")


if __name__ == "__main__":
    main()


    .
