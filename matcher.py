"""Núcleo de conciliação entre a carteira de compras e a carteira de pedidos.

A única chave de ligação entre as duas planilhas é a descrição do tecido.
A palavra "TECIDO" é considerada ruído e removida antes de qualquer comparação.
"""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass

import pandas as pd
from rapidfuzz import fuzz, process

WIDTH_RE = re.compile(r"\b(\d,\d{2})\b")
NUMBER_RE = re.compile(r"\d+")


@dataclass
class Config:
    """Parâmetros ajustáveis pela interface."""

    remover_tecido: bool = True
    remover_new: bool = True
    normalizar_codigos: bool = True
    limiar_fuzzy: int = 60
    tolerancia_qtd: float = 0.0


def sem_acentos(texto: str) -> str:
    """Remove acentos (ÁQUILA -> AQUILA, VALÊNCIA -> VALENCIA)."""
    return "".join(
        char
        for char in unicodedata.normalize("NFKD", texto)
        if not unicodedata.combining(char)
    )


def normalizar(valor, cfg: Config) -> str:
    """Normaliza uma descrição de tecido para comparação.

    Remove a palavra TECIDO (conforme exigido), sufixos /F e /P, a letra M
    isolada, acentos, pontuação e espaços repetidos.
    """
    if pd.isna(valor):
        return ""
    texto = sem_acentos(str(valor)).upper()
    if cfg.remover_tecido:
        texto = re.sub(r"\bTECIDO\b", " ", texto)
    if cfg.remover_new:
        texto = re.sub(r"\bNEW\b", " ", texto)
    texto = re.sub(r"(?<=[A-Z0-9])/(?:F|P)\b", " ", texto)
    texto = re.sub(r"\bM\b", " ", texto)
    texto = re.sub(r"[^A-Z0-9,./\- ]+", " ", texto)
    return re.sub(r"\s+", " ", texto).strip()


def campos_tecido(descricao_normalizada: str, cfg: Config) -> dict[str, str]:
    """Extrai largura, artigo e cor de uma descrição já normalizada.

    Nas duas planilhas o artigo é o primeiro código numérico após a largura e
    a cor é o código mais significativo entre os demais (a planilha de compras
    ainda inclui o nome da cor e, às vezes, um subcódigo como em '1000-ZL 02').
    """
    match = WIDTH_RE.search(descricao_normalizada)
    largura = match.group(1) if match else ""
    resto = (
        descricao_normalizada.replace(match.group(0), " ")
        if match
        else descricao_normalizada
    )
    numeros = NUMBER_RE.findall(resto)
    if cfg.normalizar_codigos:
        numeros = [str(int(n)) for n in numeros]
    artigo = numeros[0] if numeros else ""
    demais = numeros[1:]
    cor = max(demais, key=lambda n: int(n)) if demais else ""
    return {"largura": largura, "artigo": artigo, "cor": cor, "numeros": numeros}


def chave_tecido(descricao_normalizada: str, cfg: Config):
    """Chave canônica: largura + códigos numéricos normalizados.

    Em uma planilha o código da cor aparece como '0001' (4 dígitos) e na outra
    como '000001' (6 dígitos). Remover os zeros à esquerda faz os dois lados
    convergirem sem depender do nome do tecido, que muda de uma empresa para
    outra (ex.: MADRAS x MADRAS NEW, AQUILA x QUILA).
    """
    campos = campos_tecido(descricao_normalizada, cfg)
    numeros = tuple(sorted(campos["numeros"]))
    if not campos["largura"] and not numeros:
        return None
    return (campos["largura"], numeros)


def detectar_colunas(df: pd.DataFrame, papel: str) -> dict[str, str | None]:
    """Localiza colunas por nome, tolerando variações de acentuação/caixa."""
    colunas = {str(c): c for c in df.columns}
    norm = {sem_acentos(k).lower().replace(" ", ""): c for k, c in colunas.items()}

    def buscar(*termos: str):
        for termo in termos:
            alvo = sem_acentos(termo).lower().replace(" ", "")
            for nome, col in norm.items():
                if alvo in nome:
                    return col
        return None

    if papel == "compras":
        return {
            "tecido": buscar("tecido", "descricao", "produto"),
            "qtd1": buscar("comprado", "compra"),
            "qtd2": buscar("saldo", "disponivel", "estoque"),
            "codigo": buscar("codprod", "codigo", "produto"),
            "data": buscar("prev", "entrega"),
        }
    return {
        "tecido": buscar("tecido", "descricao", "produto"),
        "qtd1": buscar("vendido", "venda"),
        "qtd2": buscar("empenhado", "empenho", "reservado"),
        "codigo": buscar("pedido", "ordem"),
        "data": buscar("dtpedido", "data", "pedido"),
        "cliente": buscar("cliente", "razão", "razao"),
    }


def _para_numero(serie: pd.Series) -> pd.Series:
    return pd.to_numeric(serie, errors="coerce").fillna(0.0)


def _descrever(chave, mapa_desc: dict) -> str:
    if chave is None:
        return "(sem descrição)"
    descricoes = mapa_desc.get(chave, [])
    if not descricoes:
        largura, numeros = chave
        return f"{' '.join(numeros)} {largura}".strip()
    return descricoes[0]


def _soma_e_lista(
    grupo: pd.DataFrame, qtd1: str, qtd2: str | None, codigo: str | None
) -> tuple[float, float, str, int]:
    q1 = float(_para_numero(grupo[qtd1]).sum()) if qtd1 in grupo else 0.0
    q2 = float(_para_numero(grupo[qtd2]).sum()) if qtd2 and qtd2 in grupo else 0.0
    codigos = ""
    if codigo and codigo in grupo:
        unicos = []
        for valor in grupo[codigo].dropna().tolist():
            texto = str(valor).strip()
            if texto and texto not in unicos:
                unicos.append(texto)
        codigos = ", ".join(unicos[:8]) + (" ..." if len(unicos) > 8 else "")
    return q1, q2, codigos, len(grupo)


def conciliar(
    df_compras: pd.DataFrame,
    df_pedidos: pd.DataFrame,
    cols_compras: dict,
    cols_pedidos: dict,
    cfg: Config,
) -> dict[str, pd.DataFrame | dict]:
    """Executa a conciliação completa e devolve tabelas prontas para exibição."""
    cc, cp = cols_compras, cols_pedidos
    tc, tp = cc["tecido"], cp["tecido"]

    compras = df_compras.copy()
    pedidos = df_pedidos.copy()
    compras["_desc"] = compras[tc].map(lambda v: normalizar(v, cfg))
    pedidos["_desc"] = pedidos[tp].map(lambda v: normalizar(v, cfg))
    compras["_chave"] = compras["_desc"].map(lambda v: chave_tecido(v, cfg))
    pedidos["_chave"] = pedidos["_desc"].map(lambda v: chave_tecido(v, cfg))

    # Ignora linhas sem descrição de tecido.
    compras_validas = compras[compras["_chave"].notna()].copy()
    pedidos_validos = pedidos[pedidos["_chave"].notna()].copy()

    desc_compras = {}
    desc_pedidos = {}
    for chave, grupo in compras_validas.groupby("_chave"):
        desc_compras[chave] = grupo["_desc"].dropna().unique().tolist()
    for chave, grupo in pedidos_validos.groupby("_chave"):
        desc_pedidos[chave] = grupo["_desc"].dropna().unique().tolist()

    chaves_compras = list(desc_compras)
    chaves_pedidos = list(desc_pedidos)
    set_compras = set(chaves_compras)
    set_pedidos = set(chaves_pedidos)

    exatas = sorted(set_compras & set_pedidos)

    campos_compras = {
        chave: campos_tecido(_descrever(chave, desc_compras), cfg) for chave in chaves_compras
    }
    campos_pedidos = {
        chave: campos_tecido(_descrever(chave, desc_pedidos), cfg) for chave in chaves_pedidos
    }

    # Etapa fuzzy: somente pedidos que não casaram exatamente. Exigimos o mesmo
    # artigo e a mesma cor para não transformar divergências reais em "match".
    pendentes_pedidos = [c for c in chaves_pedidos if c not in set_compras]
    candidatos = [c for c in chaves_compras if c not in set_pedidos]

    aproximadas: dict[tuple, tuple] = {}
    pendentes_sem_match = list(pendentes_pedidos)
    for chave_pedido in pendentes_pedidos:
        origem = _descrever(chave_pedido, desc_pedidos)
        cp_campos = campos_pedidos[chave_pedido]
        elegiveis = [
            c
            for c in candidatos
            if campos_compras[c]["artigo"] == cp_campos["artigo"]
            and campos_compras[c]["cor"] == cp_campos["cor"]
        ]
        if not elegiveis:
            continue
        elegiveis_desc = [_descrever(c, desc_compras) for c in elegiveis]
        resultado = process.extractOne(
            origem, elegiveis_desc, scorer=fuzz.WRatio, score_cutoff=cfg.limiar_fuzzy
        )
        if resultado:
            _, score, indice = resultado
            chave_compra = elegiveis[indice]
            aproximadas[(chave_pedido, chave_compra)] = (origem, float(score))
            pendentes_sem_match.remove(chave_pedido)

    usadas_fuzzy = {par[1] for par in aproximadas}
    sobras_compras = [c for c in chaves_compras if c not in set_pedidos and c not in usadas_fuzzy]

    registros: list[dict] = []

    def adicionar(
        tipo: str,
        chave_pedido,
        chave_compra,
        Similaridade,
        g_pedidos: pd.DataFrame | None,
        g_compras: pd.DataFrame | None,
    ) -> None:
        vendido, empenhado, cod_pedido, n_ped = (0.0, 0.0, "", 0)
        comprado, saldo, cod_compra, n_compra = (0.0, 0.0, "", 0)
        cliente = ""
        if g_pedidos is not None and len(g_pedidos):
            vendido, empenhado, cod_pedido, n_ped = _soma_e_lista(
                g_pedidos, cp["qtd1"], cp.get("qtd2"), cp.get("codigo")
            )
            if cp.get("cliente") and cp["cliente"] in g_pedidos:
                cliente = ", ".join(
                    sorted({str(v) for v in g_pedidos[cp["cliente"]].dropna()})[:3]
                )
        if g_compras is not None and len(g_compras):
            comprado, saldo, cod_compra, n_compra = _soma_e_lista(
                g_compras, cc["qtd1"], cc.get("qtd2"), cc.get("codigo")
            )

        demanda = vendido - empenhado
        folga = saldo - demanda
        if g_pedidos is None:
            status = "Sem pedido vinculado"
        elif g_compras is None:
            status = "Sem compra vinculada"
        elif demanda > saldo + cfg.tolerancia_qtd:
            status = "Saldo insuficiente"
        elif demanda < 0:
            status = "Empenhado > vendido"
        elif vendido > comprado + cfg.tolerancia_qtd:
            status = "Vendido > comprado"
        else:
            status = "OK"

        descricao_p = _descrever(chave_pedido, desc_pedidos) if chave_pedido is not None else ""
        descricao_c = _descrever(chave_compra, desc_compras) if chave_compra is not None else ""
        origem_campos = (
            campos_pedidos.get(chave_pedido)
            or campos_compras.get(chave_compra)
            or {"artigo": "", "cor": ""}
        )

        registros.append(
            {
                "Tipo": tipo,
                "Artigo": origem_campos["artigo"],
                "Cor": origem_campos["cor"],
                "Largura": origem_campos.get("largura", ""),
                "Tecido (pedido)": descricao_p,
                "Tecido (compra)": descricao_c,
                "Chave": str(chave_pedido or chave_compra),
                "Similaridade %": round(Similaridade, 1) if Similaridade else None,
                "Qtd. Vendida": vendido,
                "Empenhado": empenhado,
                "Demanda líquida": demanda,
                "Qtd. Comprada": comprado,
                "Saldo": saldo,
                "Folga (saldo - demanda)": folga,
                "Status quantidade": status,
                "Pedidos": cod_pedido,
                "Cliente": cliente,
                "Compras": cod_compra,
                "Linhas pedido": n_ped,
                "Linhas compra": n_compra,
                "É conflito": tipo.startswith("Conflito")
                or status in {"Saldo insuficiente", "Vendido > comprado"},
            }
        )

    for chave in exatas:
        adicionar(
            "Correspondência exata",
            chave,
            chave,
            100.0,
            pedidos_validos[pedidos_validos["_chave"] == chave],
            compras_validas[compras_validas["_chave"] == chave],
        )

    for (chave_p, chave_c), (_, score) in aproximadas.items():
        adicionar(
            "Conflito: divergência de descrição",
            chave_p,
            chave_c,
            score,
            pedidos_validos[pedidos_validos["_chave"] == chave_p],
            compras_validas[compras_validas["_chave"] == chave_c],
        )

    for chave in pendentes_sem_match:
        adicionar(
            "Conflito: pedido sem compra",
            chave,
            None,
            None,
            pedidos_validos[pedidos_validos["_chave"] == chave],
            None,
        )

    for chave in sobras_compras:
        adicionar(
            "Conflito: compra sem pedido",
            None,
            chave,
            None,
            None,
            compras_validas[compras_validas["_chave"] == chave],
        )

    tabela = pd.DataFrame(registros)
    if tabela.empty:
        tabela = pd.DataFrame(
            columns=[
                "Tipo", "Artigo", "Cor", "Largura",
                "Tecido (pedido)", "Tecido (compra)", "Chave",
                "Similaridade %", "Qtd. Vendida", "Empenhado", "Demanda líquida",
                "Qtd. Comprada", "Saldo", "Folga (saldo - demanda)",
                "Status quantidade", "Pedidos", "Cliente", "Compras",
                "Linhas pedido", "Linhas compra", "É conflito",
            ]
        )

    ordem = {
        "Conflito: pedido sem compra": 0,
        "Conflito: compra sem pedido": 1,
        "Conflito: divergência de descrição": 2,
        "Correspondência exata": 3,
    }
    tabela["_ordem"] = tabela["Tipo"].map(ordem).fillna(9)
    tabela = (
        tabela.sort_values(["_ordem", "Tecido (pedido)", "Tecido (compra)"])
        .drop(columns="_ordem")
        .reset_index(drop=True)
    )

    conflitos = tabela[tabela["É conflito"]].copy().reset_index(drop=True)
    correspondencias = tabela[~tabela["É conflito"]].copy().reset_index(drop=True)

    resumo = (
        tabela.groupby(["Tipo", "Status quantidade"])
        .size()
        .reset_index(name="Itens")
        .sort_values("Itens", ascending=False)
        .reset_index(drop=True)
    )

    metricas = {
        "linhas_compras": int(len(compras)),
        "linhas_pedidos": int(len(pedidos)),
        "tecidos_compras": int(len(chaves_compras)),
        "tecidos_pedidos": int(len(chaves_pedidos)),
        "exatas": int(len(exatas)),
        "aproximadas": int(len(aproximadas)),
        "sem_compra": int(len(pendentes_sem_match)),
        "sem_pedido": int(len(sobras_compras)),
        "qtd_insuficiente": int((tabela["Status quantidade"] == "Saldo insuficiente").sum()),
        "total_conflitos": int(len(conflitos)),
        "ignoradas_compras": int(len(compras) - len(compras_validas)),
        "ignoradas_pedidos": int(len(pedidos) - len(pedidos_validos)),
    }

    return {
        "tabela": tabela,
        "conflitos": conflitos,
        "correspondencias": correspondencias,
        "resumo": resumo,
        "metricas": metricas,
        "normalizacao": {
            "compras": compras[["_desc", "_chave"]].copy(),
            "pedidos": pedidos[["_desc", "_chave"]].copy(),
        },
    }
