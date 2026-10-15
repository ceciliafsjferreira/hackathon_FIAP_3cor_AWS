import json
import unicodedata
from pathlib import Path

import pandas as pd


BASE_DIR = Path(__file__).resolve().parent
SAIDA_DIR = BASE_DIR / "saida"
ARQ_ORDEM = BASE_DIR / "ORDEN BASIS TESTE.xlsx"
ARQ_CONFERENCIA = BASE_DIR / "conferência de produtos.xlsx"

CHAVE = "Material"
MOEDA_PLATAFORMA = "USD"


HEADER = {"TAB": 1, "AUX": 3, "PROD": 1, "DADOS CLIENTES": 1, "ORDEM": 15}


def _slug(nome: str) -> str:
    nome = str(nome).strip().lower()
    nome = unicodedata.normalize("NFKD", nome).encode("ascii", "ignore").decode()
    slug = "".join(c if c.isalnum() else "_" for c in nome)
    while "__" in slug:
        slug = slug.replace("__", "_")
    return slug.strip("_")


def _material_int(serie: pd.Series) -> pd.Series:
    return pd.to_numeric(serie, errors="coerce").astype("Int64")


def ler_aba(aba: str) -> pd.DataFrame:
    if not ARQ_ORDEM.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {ARQ_ORDEM}")
    print(f"  Lendo aba: {aba}")
    return pd.read_excel(ARQ_ORDEM, sheet_name=aba, engine="openpyxl",
                         header=HEADER.get(aba, 0))


def preparar_tab(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df = df[df["Material"].notna()]
    df[CHAVE] = _material_int(df[CHAVE])
    df = df[df[CHAVE].notna()]

    out = pd.DataFrame({
        CHAVE: df[CHAVE],
        "cliente": df.get("CLIENTE"),
        "org": df.get("Org"),
        "lista_preco": df.get("Lista de Preço"),
        "descricao": df.get("Descrição do Material Português"),
        "unidade_med": df.get("Un Med"),
        "preco_sugerido": pd.to_numeric(df.get("Preço Sugerido"), errors="coerce"),
        "moeda": df.get("Moeda Tabela Preço"),
        "umb": df.get("UMB"),
        "paletizacao_umb": pd.to_numeric(df.get("PALETIZAÇAO UMB"), errors="coerce"),
        "paletizacao_cx": pd.to_numeric(df.get("PALETIZAÇÃO CX"), errors="coerce"),
    })
    return out.drop_duplicates(subset=CHAVE, keep="first").reset_index(drop=True)


def preparar_descricoes(df_aux: pd.DataFrame) -> pd.DataFrame:
    df = df_aux.copy()
    df = df[df["Material"].notna() & df["LANGUAGE"].notna()]
    df[CHAVE] = _material_int(df[CHAVE])
    df = df[df[CHAVE].notna()]

    mapa_idioma = {"PORTUGUÊS": "pt", "ESPAÑOL": "es", "ENGLISH": "en"}
    df["idioma"] = df["LANGUAGE"].map(mapa_idioma)
    df = df[df["idioma"].notna()]

    desc = (df.pivot_table(index=CHAVE, columns="idioma", values="DESC",
                           aggfunc="first")
              .rename(columns={"pt": "descricao_pt", "es": "descricao_es",
                               "en": "descricao_en"})
              .reset_index())

    fam_pt = df[df["idioma"] == "pt"].groupby(CHAVE)["FAMÍLIA TRADUÇÃO"].first()
    fam_en = df[df["idioma"] == "en"].groupby(CHAVE)["FAMÍLIA TRADUÇÃO"].first()

    pares = pd.concat([fam_pt.rename("pt"), fam_en.rename("en")], axis=1).dropna()
    mapa_pt_en = (pares.drop_duplicates("pt").set_index("pt")["en"].to_dict())

    todos = fam_pt.index.union(fam_en.index)
    familia = {}
    for m in todos:
        en = fam_en.get(m)
        if pd.notna(en):
            familia[m] = en
        else:
            pt = fam_pt.get(m)
            familia[m] = mapa_pt_en.get(pt, pt)

    fam = (pd.Series(familia, name="familia").rename_axis(CHAVE).reset_index())
    return desc.merge(fam, on=CHAVE, how="left")


def agregar_prod(df_prod: pd.DataFrame) -> pd.DataFrame:
    df = df_prod.copy()
    df = df[df["Material"].notna()]
    df[CHAVE] = _material_int(df[CHAVE])
    df = df[df[CHAVE].notna()]

    agg = df.groupby(CHAVE).agg(
        qtd_unidades_medida=("Unidade de medida alternativa", "nunique"),
        unidades_medida=("Unidade de medida alternativa",
                         lambda s: ", ".join(sorted(s.dropna().astype(str).unique()))),
        volume_max=("VolUNe", "max"),
        tem_ean=("Código EAN/UPC", lambda s: s.notna().any()),
    ).reset_index()

    const = (df.groupby(CHAVE)[["LINHA", "NCM"]].first().reset_index()
               .rename(columns={"LINHA": "linha", "NCM": "ncm"}))
    return agg.merge(const, on=CHAVE, how="left")


def pesos_conferencia(caminho: Path) -> pd.DataFrame:
    if not caminho.exists():
        print(f"  [aviso] '{caminho.name}' não encontrado; pesos ficarão nulos.")
        return pd.DataFrame({CHAVE: pd.Series(dtype="Int64"),
                             "peso_bruto_max": [], "peso_liquido_max": []})
    df = pd.read_excel(caminho, engine="openpyxl")
    df = df[df["Material"].notna()]
    df[CHAVE] = _material_int(df[CHAVE])
    df = df[df[CHAVE].notna()]
    return (df.groupby(CHAVE)
              .agg(peso_bruto_max=("Peso bruto", "max"),
                   peso_liquido_max=("Peso líquido", "max"))
              .reset_index())


def preparar_pedido(df_ordem: pd.DataFrame) -> pd.DataFrame:
    df = df_ordem.copy()
    col_cod = "CÓDIGOS" if "CÓDIGOS" in df.columns else df.columns[5]
    df = df[pd.to_numeric(df[col_cod], errors="coerce").notna()]
    df[CHAVE] = _material_int(df[col_cod])

    def _num(col):
        return pd.to_numeric(df.get(col), errors="coerce")

    out = pd.DataFrame({
        CHAVE: df[CHAVE],
        "caixas_pedidas": _num("CAIXAS").fillna(0),
        "pallets_pedidos": _num("QUANTIDADE DE PALLETS").fillna(0),
        "valor_pedido": _num("VALOR").fillna(0.0),
        "peso_bruto_pedido": _num("PESO BRUTO.1").fillna(0.0),
        "peso_liq_pedido": _num("PESO LIQ.1").fillna(0.0),
        "cubagem_pedido": _num("CUBAGEM TOTAL").fillna(0.0),
        "pedido_validado": df.get("PEDIDO VALIDADO ?").astype(str).str.strip().str.upper(),
    })
    out["foi_pedido"] = out["caixas_pedidas"] > 0
    return out.drop_duplicates(subset=CHAVE, keep="first").reset_index(drop=True)


def cabecalho_pedido(df_ordem_raw: pd.DataFrame) -> dict:
    def _busca(rotulo):
        for _, row in df_ordem_raw.iterrows():
            vals = [str(v).strip() for v in row.tolist() if pd.notna(v)]
            for i, v in enumerate(vals):
                if v.upper().startswith(rotulo.upper()) and i + 1 < len(vals):
                    return vals[i + 1]
        return None

    return {
        "referencia_cliente": _busca("REFERÊNCIA CLIENTE:"),
        "filial": _busca("FILIAL"),
        "embarque": _busca("EMBARQUE:"),
        "incoterms": _busca("INCOTERMS:"),
        "prazo_pagamento": _busca("PRAZO DE PAGAMENTO:"),
    }


def extrair_cliente(df_cli: pd.DataFrame) -> dict:
    df = df_cli.dropna(how="all")
    if df.empty:
        return {}
    r = df.iloc[0]
    return {
        "cliente": str(r.get("RESUMO", "")).strip() or None,
        "codigo": None if pd.isna(r.get("Cliente")) else str(r.get("Cliente")),
        "moeda": None if pd.isna(r.get("Moeda")) else str(r.get("Moeda")),
        "incoterm": None if pd.isna(r.get("Incoterm")) else str(r.get("Incoterm")),
        "cond_pagamento": None if pd.isna(r.get("Cond. Pagto"))
                          else str(r.get("Cond. Pagto")),
    }


def cruzar(tab, descricoes, prod, pesos, pedido) -> pd.DataFrame:
    res = (tab
           .merge(descricoes, on=CHAVE, how="left")
           .merge(prod, on=CHAVE, how="left")
           .merge(pesos, on=CHAVE, how="left")
           .merge(pedido, on=CHAVE, how="left"))

    res["tem_preco"] = res["preco_sugerido"].notna()
    res["tem_logistica"] = res["qtd_unidades_medida"].notna()
    res["disponivel_para_venda"] = res["tem_preco"]
    res["nome"] = res["descricao_pt"].where(
        res["descricao_pt"].notna(), res["descricao"])

    for c, default in [("caixas_pedidas", 0), ("pallets_pedidos", 0),
                       ("valor_pedido", 0.0), ("foi_pedido", False)]:
        if c in res.columns:
            res[c] = res[c].fillna(default)
    if "pedido_validado" in res.columns:
        res["pedido_validado"] = res["pedido_validado"].fillna("NÃO")
    return res


def _escrever_xlsx(func, xlsx: Path) -> bool:
    try:
        func()
        return True
    except PermissionError:
        print(f"  [ERRO] '{xlsx.name}' está aberto (Excel). Feche e rode de novo.")
        return False


def exportar_tabela(df: pd.DataFrame, nome: str) -> None:
    SAIDA_DIR.mkdir(parents=True, exist_ok=True)
    norm = df.rename(columns={c: _slug(c) for c in df.columns})
    csv = SAIDA_DIR / f"{nome}.csv"
    xlsx = SAIDA_DIR / f"{nome}.xlsx"
    norm.to_csv(csv, index=False, sep=";", encoding="utf-8-sig")
    _escrever_xlsx(lambda: norm.to_excel(xlsx, index=False, engine="openpyxl"), xlsx)
    print(f"  tabela: {len(norm)} linhas, {len(norm.columns)} colunas -> {csv.name}, {xlsx.name}")


def gerar_relatorio(df: pd.DataFrame, nome: str) -> None:
    SAIDA_DIR.mkdir(parents=True, exist_ok=True)
    xlsx = SAIDA_DIR / f"{nome}.xlsx"

    completo = df.rename(columns={c: _slug(c) for c in df.columns})
    resumo = (df.groupby(df["familia"].fillna("(sem família)"))
                .agg(qtd_materiais=(CHAVE, "count"),
                     preco_medio=("preco_sugerido", "mean"),
                     preco_min=("preco_sugerido", "min"),
                     preco_max=("preco_sugerido", "max"))
                .round(2).reset_index().rename(columns={"familia": "familia"})
                .sort_values("qtd_materiais", ascending=False))

    def _w():
        with pd.ExcelWriter(xlsx, engine="openpyxl") as wr:
            completo.to_excel(wr, sheet_name="Resultado", index=False)
            resumo.to_excel(wr, sheet_name="Resumo_familia", index=False)

    if _escrever_xlsx(_w, xlsx):
        print(f"  relatório -> {xlsx.name} (abas: Resultado, Resumo_familia)")


def exportar_json_catalogo(df: pd.DataFrame, meta_cliente: dict, nome: str) -> None:
    SAIDA_DIR.mkdir(parents=True, exist_ok=True)
    caminho = SAIDA_DIR / f"{nome}.json"
    vis = df[df["disponivel_para_venda"]].copy()

    def _s(v):
        return None if pd.isna(v) else str(v)

    itens = []
    for _, r in vis.iterrows():
        itens.append({
            "material": int(r[CHAVE]),
            "nome": _s(r["nome"]),
            "descricao_pt": _s(r.get("descricao_pt")),
            "descricao_es": _s(r.get("descricao_es")),
            "descricao_en": _s(r.get("descricao_en")),
            "familia": _s(r.get("familia")),
            "unidade_venda": _s(r.get("unidade_med")) or _s(r.get("umb")),
            "paletizacao_caixas": None if pd.isna(r.get("paletizacao_cx"))
                                  else float(r["paletizacao_cx"]),
            "peso_bruto": None if pd.isna(r.get("peso_bruto_max"))
                          else float(r["peso_bruto_max"]),
            "peso_liquido": None if pd.isna(r.get("peso_liquido_max"))
                            else float(r["peso_liquido_max"]),
            "preco_venda": round(float(r["preco_sugerido"]), 4),
            "moeda": MOEDA_PLATAFORMA,
            "caixas_pedidas": float(r.get("caixas_pedidas", 0) or 0),
            "valor_pedido": round(float(r.get("valor_pedido", 0) or 0), 2),
            "disponivel": True,
        })

    doc = {
        "schema": "catalogo", "versao": 2, "moeda": MOEDA_PLATAFORMA,
        "cliente": meta_cliente,
        "gerado_em": pd.Timestamp.now().isoformat(timespec="seconds"),
        "total_itens": len(itens), "itens": itens,
    }
    caminho.write_text(json.dumps(doc, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"  catálogo (cliente): {len(itens)} itens -> {caminho.name}")


def exportar_json_pedido(df: pd.DataFrame, meta_cliente: dict,
                         cabecalho: dict, nome: str) -> None:
    SAIDA_DIR.mkdir(parents=True, exist_ok=True)
    caminho = SAIDA_DIR / f"{nome}.json"

    pedidos = df[df.get("foi_pedido", False) == True].copy()

    def _s(v):
        return None if pd.isna(v) else str(v)

    itens = []
    for _, r in pedidos.iterrows():
        itens.append({
            "material": int(r[CHAVE]),
            "nome": _s(r["nome"]),
            "familia": _s(r.get("familia")),
            "preco_unitario": round(float(r["preco_sugerido"]), 4),
            "caixas": float(r["caixas_pedidas"]),
            "pallets": None if pd.isna(r.get("pallets_pedidos"))
                       else float(r["pallets_pedidos"]),
            "valor_total": round(float(r.get("valor_pedido", 0) or 0), 2),
            "peso_bruto": None if pd.isna(r.get("peso_bruto_pedido"))
                          else float(r["peso_bruto_pedido"]),
            "peso_liquido": None if pd.isna(r.get("peso_liq_pedido"))
                            else float(r["peso_liq_pedido"]),
            "validado": _s(r.get("pedido_validado")) == "SIM",
            "moeda": MOEDA_PLATAFORMA,
        })

    totais = {
        "qtd_itens": len(itens),
        "valor_total": round(sum(i["valor_total"] for i in itens), 2),
        "caixas_total": sum(i["caixas"] for i in itens),
        "peso_liquido_total": round(
            sum((i["peso_liquido"] or 0) for i in itens), 2),
        "peso_bruto_total": round(
            sum((i["peso_bruto"] or 0) for i in itens), 2),
    }

    doc = {
        "schema": "pedido", "versao": 1, "moeda": MOEDA_PLATAFORMA,
        "cliente": meta_cliente, "cabecalho": cabecalho,
        "gerado_em": pd.Timestamp.now().isoformat(timespec="seconds"),
        "totais": totais, "itens": itens,
    }
    caminho.write_text(json.dumps(doc, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"  pedido (vendedor): {len(itens)} itens pedidos, "
          f"valor total {totais['valor_total']} {MOEDA_PLATAFORMA} -> {caminho.name}")


def exportar_json_interno(df: pd.DataFrame, meta_cliente: dict, nome: str) -> None:
    SAIDA_DIR.mkdir(parents=True, exist_ok=True)
    caminho = SAIDA_DIR / f"{nome}.json"
    norm = df.rename(columns={c: _slug(c) for c in df.columns})
    registros = json.loads(norm.to_json(orient="records", force_ascii=False,
                                        date_format="iso"))
    doc = {
        "schema": "produtos_internos", "versao": 2,
        "moeda_plataforma": MOEDA_PLATAFORMA, "cliente": meta_cliente,
        "gerado_em": pd.Timestamp.now().isoformat(timespec="seconds"),
        "total_itens": len(registros), "itens": registros,
    }
    caminho.write_text(json.dumps(doc, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"  base interna (vendedor): {len(registros)} itens -> {caminho.name}")


def main() -> None:
    print("1. EXTRACT (ORDEN BASIS TESTE.xlsx + conferência de produtos.xlsx)")
    tab_raw = ler_aba("TAB")
    aux_raw = ler_aba("AUX")
    prod_raw = ler_aba("PROD")
    cli_raw = ler_aba("DADOS CLIENTES")
    ordem_itens_raw = ler_aba("ORDEM")
    ordem_cab_raw = pd.read_excel(ARQ_ORDEM, sheet_name="ORDEM",
                                  engine="openpyxl", header=None, nrows=14)

    print("\n2. TRANSFORM")
    tab = preparar_tab(tab_raw)
    descricoes = preparar_descricoes(aux_raw)
    prod = agregar_prod(prod_raw)
    pesos = pesos_conferencia(ARQ_CONFERENCIA)
    pedido = preparar_pedido(ordem_itens_raw)
    meta_cliente = extrair_cliente(cli_raw)
    cabecalho = cabecalho_pedido(ordem_cab_raw)
    print(f"  precos (TAB): {tab.shape} | descrições (AUX): {descricoes.shape} | "
          f"logística (PROD): {prod.shape}")
    print(f"  pesos (conferência): {pesos.shape} | pedido (ORDEM): {pedido.shape}")
    print(f"  cliente: {meta_cliente.get('cliente')} ({meta_cliente.get('moeda')}) "
          f"| ref: {cabecalho.get('referencia_cliente')}")

    print("\n3. MERGE (base = TAB, chave = 'Material')")
    resultado = cruzar(tab, descricoes, prod, pesos, pedido)
    n_pedidos = int(resultado.get("foi_pedido", pd.Series(dtype=bool)).sum())
    print(f"  {len(resultado)} materiais | com preço: {int(resultado['tem_preco'].sum())} "
          f"| disponíveis: {int(resultado['disponivel_para_venda'].sum())} "
          f"| itens pedidos: {n_pedidos}")

    print("\n4. EXPORT (análise)")
    exportar_tabela(resultado, "cruzamento_resultado")
    gerar_relatorio(resultado, "cruzamento_relatorio")

    print("\n5. SAÍDA PARA A PLATAFORMA (JSON)")
    exportar_json_catalogo(resultado, meta_cliente, "catalogo")
    exportar_json_interno(resultado, meta_cliente, "produtos_internos")
    exportar_json_pedido(resultado, meta_cliente, cabecalho, "pedido")


if __name__ == "__main__":
    main()
