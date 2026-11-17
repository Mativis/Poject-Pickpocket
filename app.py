"""Interface Streamlit para resolução de conflito entre carteiras.

Execução:
    streamlit run app.py
"""

from __future__ import annotations

import time
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st

from matcher import Config, conciliar, detectar_colunas, normalizar, campos_tecido

BASE = Path(__file__).resolve().parent
PADRAO_COMPRAS = BASE / "Carteira Nazareth 17112026.xls"
PADRAO_PEDIDOS = BASE / "CARTEIRA NAZARETH.xls"

st.set_page_config(
    page_title="Conciliação de Carteiras",
    page_icon="🧵",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.title("🧵 Resolução de conflito de carteira de pedidos")
st.caption(
    "A única ligação entre as duas empresas é a **descrição do tecido**. "
    "A palavra **TECIDO** é descartada antes de qualquer comparação."
)


def carregar(upload, caminho_padrao: Path, rotulo: str) -> pd.DataFrame | None:
    if upload is not None:
        try:
            return pd.read_excel(upload)
        except Exception as erro:  # pragma: no cover - feedback de UI
            st.sidebar.error(f"Não foi possível ler {rotulo}: {erro}")
            return None
    if caminho_padrao.exists():
        return pd.read_excel(caminho_padrao)
    return None


def escolher_coluna(label: str, detectada, df: pd.DataFrame, chave: str):
    opcoes = [None, *list(df.columns)]
    indice = opcoes.index(detectada) if detectada in opcoes else 0
    return st.sidebar.selectbox(label, opcoes, index=indice, key=chave)


def para_csv(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig")


# ---------------------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------------------
st.sidebar.header("1. Carteiras")
up_compras = st.sidebar.file_uploader(
    "Carteira de COMPRAS (.xls/.xlsx)", type=["xls", "xlsx"], key="up_compras"
)
up_pedidos = st.sidebar.file_uploader(
    "Carteira de PEDIDOS (.xls/.xlsx)", type=["xls", "xlsx"], key="up_pedidos"
)

df_compras = carregar(up_compras, PADRAO_COMPRAS, "compras")
df_pedidos = carregar(up_pedidos, PADRAO_PEDIDOS, "pedidos")

if df_compras is None or df_pedidos is None:
    st.warning(
        "Envie as duas planilhas na barra lateral ou deixe os arquivos de exemplo "
        "na mesma pasta deste programa."
    )
    st.stop()

st.sidebar.caption(
    f"Compras: {len(df_compras)} linhas · Pedidos: {len(df_pedidos)} linhas"
)

st.sidebar.header("2. Colunas")
det_compras = detectar_colunas(df_compras, "compras")
det_pedidos = detectar_colunas(df_pedidos, "pedidos")

cols_compras = {
    "tecido": escolher_coluna("Descrição do tecido · compras", det_compras["tecido"], df_compras, "c_tecido"),
    "qtd1": escolher_coluna("Quantidade comprada", det_compras["qtd1"], df_compras, "c_qtd1"),
    "qtd2": escolher_coluna("Saldo disponível", det_compras["qtd2"], df_compras, "c_qtd2"),
    "codigo": escolher_coluna("Código do produto", det_compras["codigo"], df_compras, "c_codigo"),
    "data": det_compras["data"],
}
cols_pedidos = {
    "tecido": escolher_coluna("Descrição do tecido · pedidos", det_pedidos["tecido"], df_pedidos, "p_tecido"),
    "qtd1": escolher_coluna("Quantidade vendida", det_pedidos["qtd1"], df_pedidos, "p_qtd1"),
    "qtd2": escolher_coluna("Empenhado", det_pedidos["qtd2"], df_pedidos, "p_qtd2"),
    "codigo": escolher_coluna("Número do pedido", det_pedidos["codigo"], df_pedidos, "p_codigo"),
    "data": det_pedidos["data"],
    "cliente": det_pedidos["cliente"],
}

if not cols_compras["tecido"] or not cols_pedidos["tecido"]:
    st.error("Selecione a coluna de descrição do tecido nas duas carteiras.")
    st.stop()

st.sidebar.header("3. Regras de casamento")
remover_tecido = st.sidebar.checkbox("Ignorar a palavra TECIDO", value=True)
remover_new = st.sidebar.checkbox("Ignorar a palavra NEW", value=True)
normalizar_codigos = st.sidebar.checkbox("Normalizar zeros à esquerda dos códigos", value=True)
limiar = st.sidebar.slider(
    "Sensibilidade para divergência de descrição (%)", 40, 100, 60, 5
)
tolerancia = st.sidebar.number_input("Tolerância de quantidade", min_value=0.0, value=0.0, step=1.0)

cfg = Config(
    remover_tecido=remover_tecido,
    remover_new=remover_new,
    normalizar_codigos=normalizar_codigos,
    limiar_fuzzy=limiar,
    tolerancia_qtd=tolerancia,
)

# ---------------------------------------------------------------------------
# Processamento
# ---------------------------------------------------------------------------
inicio = time.perf_counter()
with st.spinner("Conciliando carteiras..."):
    resultado = conciliar(df_compras, df_pedidos, cols_compras, cols_pedidos, cfg)
elapsed = time.perf_counter() - inicio

m = resultado["metricas"]
conflitos = resultado["conflitos"]
correspondencias = resultado["correspondencias"]

st.success(f"Conciliação concluída em **{elapsed:.2f} s** — {m['total_conflitos']} itens de conflito encontrados.")

k1, k2, k3, k4, k5, k6 = st.columns(6)
k1.metric("Tecidos em pedidos", m["tecidos_pedidos"])
k2.metric("Casamentos exatos", m["exatas"])
k3.metric("Pedidos sem compra", m["sem_compra"], delta=None)
k4.metric("Compras sem pedido", m["sem_pedido"])
k5.metric("Divergências de descrição", m["aproximadas"])
k6.metric("Saldo insuficiente", m["qtd_insuficiente"])

st.caption(
    f"Linhas processadas: {m['linhas_pedidos']} pedidos e {m['linhas_compras']} compras · "
    f"linhas sem descrição ignoradas: {m['ignoradas_pedidos']} pedidos / {m['ignoradas_compras']} compras."
)

# ---------------------------------------------------------------------------
# Exportação
# ---------------------------------------------------------------------------
buffer = BytesIO()
with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
    conflitos.to_excel(writer, sheet_name="Conflitos", index=False)
    correspondencias.to_excel(writer, sheet_name="Correspondencias", index=False)
    resultado["resumo"].to_excel(writer, sheet_name="Resumo", index=False)

exportar, csv_conf, csv_corr = st.columns([2, 1, 1])
with exportar:
    st.download_button(
        "⬇️ Relatório completo (Excel, 3 abas)",
        buffer.getvalue(),
        "relatorio_conflito_carteiras.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        width="stretch",
    )
with csv_conf:
    st.download_button(
        "⬇️ Conflitos (CSV)",
        para_csv(conflitos),
        "conflitos.csv",
        "text/csv",
        width="stretch",
    )
with csv_corr:
    st.download_button(
        "⬇️ Correspondências (CSV)",
        para_csv(correspondencias),
        "correspondencias.csv",
        "text/csv",
        width="stretch",
    )

st.divider()

# ---------------------------------------------------------------------------
# Abas
# ---------------------------------------------------------------------------
aba_conflitos, aba_corr, aba_resumo, ajuda = st.tabs(
    ["⚠️ Conflitos", "✅ Correspondências", "📊 Resumo", "ℹ️ Como funciona"]
)

NUMERICAS = [
    "Qtd. Vendida", "Empenhado", "Demanda líquida", "Qtd. Comprada",
    "Saldo", "Folga (saldo - demanda)", "Similaridade %",
]
ORDENAR = ["Tipo", "Artigo", "Cor", "Tecido (pedido)", "Tecido (compra)"]


def mostrar(df: pd.DataFrame, chave: str) -> None:
    if df.empty:
        st.info("Nenhum registro nesta categoria com os filtros atuais.")
        return
    colunas = [c for c in ORDENAR if c in df.columns] + [
        c for c in df.columns if c not in ORDENAR
    ]
    st.dataframe(
        df[colunas],
        width="stretch",
        hide_index=True,
        height=min(520, 90 + len(df) * 32),
        column_config={
            col: st.column_config.NumberColumn(col, format="%.1f")
            for col in NUMERICAS
            if col in df.columns
        },
        key=chave,
    )


with aba_conflitos:
    f1, f2 = st.columns(2)
    with f1:
        tipos = st.multiselect(
            "Tipo de conflito",
            options=sorted(conflitos["Tipo"].unique()) if not conflitos.empty else [],
            default=sorted(conflitos["Tipo"].unique()) if not conflitos.empty else [],
            key="filtro_tipo",
        )
    with f2:
        status = st.multiselect(
            "Status de quantidade",
            options=sorted(conflitos["Status quantidade"].unique()) if not conflitos.empty else [],
            default=sorted(conflitos["Status quantidade"].unique()) if not conflitos.empty else [],
            key="filtro_status",
        )
    filtrados = conflitos[
        conflitos["Tipo"].isin(tipos) & conflitos["Status quantidade"].isin(status)
    ]
    st.caption(f"{len(filtrados)} de {len(conflitos)} conflitos exibidos")
    mostrar(filtrados, "tabela_conflitos")

with aba_corr:
    st.caption(
        "Tecidos presentes nas duas carteiras com quantidades conciliadas. "
        "Itens com divergência de descrição aparecem na aba de conflitos."
    )
    mostrar(correspondencias, "tabela_correspondencias")

with aba_resumo:
    st.subheader("Itens por tipo e status")
    resumo = resultado["resumo"]
    st.dataframe(resumo, width="stretch", hide_index=True)

    por_tipo = (
        resumo.groupby("Tipo", as_index=False)["Itens"]
        .sum()
        .sort_values("Itens", ascending=False)
    )
    if not por_tipo.empty:
        st.bar_chart(por_tipo.set_index("Tipo")["Itens"])

    st.subheader("Maiores faltas de saldo")
    falta = conflitos[conflitos["Status quantidade"] == "Saldo insuficiente"].copy()
    if not falta.empty:
        falta = falta.sort_values("Folga (saldo - demanda)")
        mostrar(falta, "tabela_faltas")
    else:
        st.info("Nenhuma insuficiência de saldo nos tecidos casados.")

with ajuda:
    st.markdown(
        """
        ### Critério de ligação

        As duas planilhas não compartilham códigos de cliente, pedido ou produto.
        O único campo comum é a **descrição do tecido**, e ela vem em formatos
        diferentes:

        | Carteira de compras | Carteira de pedidos |
        |---|---|
        | `9558 MADRAS 2,50 9330 INDIGO` | `TECIDO MADRAS NEW 2,50 M 9558/P 009330` |

        Conforme solicitado, a palavra **TECIDO** é removida. Depois o sistema:

        1. retira acentos, sufixos `/F` e `/P`, a letra `M` isolada e pontuação;
        2. zera os zeros à esquerda dos códigos (`0001` → `1`, `000001` → `1`);
        3. monta a chave com **largura + códigos numéricos** — o nome do tecido
           não interfere (ele muda de uma empresa para outra);
        4. para o que sobrou, procura uma correspondência aproximada exigindo
           **o mesmo artigo e a mesma cor**, sinalizando divergências reais de
           descrição (ex.: largura `2,80` vs `2,90`);
        5. compara quantidades: `Demanda líquida = Vendido − Empenhado` contra o
           `Saldo` comprado.

        ### Tipos de conflito

        - **Pedido sem compra** — há demanda e nenhum tecido comprado equivalente;
        - **Compra sem pedido** — há tecido comprado sem demanda vinculada;
        - **Divergência de descrição** — artigo e cor iguais, mas a descrição
          difere (largura, subcódigo ou redação);
        - **Saldo insuficiente** — a demanda líquida ultrapassa o saldo disponível.
        """
    )

    with st.expander("Ver exemplos de normalização aplicados"):
        cfg_demo = Config(
            remover_tecido=remover_tecido,
            remover_new=remover_new,
            normalizar_codigos=normalizar_codigos,
        )
        exemplos = []
        for nome, df, col in [
            ("Compras", df_compras, cols_compras["tecido"]),
            ("Pedidos", df_pedidos, cols_pedidos["tecido"]),
        ]:
            for valor in df[col].dropna().astype(str).head(4):
                limpo = normalizar(valor, cfg_demo)
                campos = campos_tecido(limpo, cfg_demo)
                exemplos.append(
                    {
                        "Carteira": nome,
                        "Original": valor,
                        "Normalizado": limpo,
                        "Artigo": campos["artigo"],
                        "Cor": campos["cor"],
                    }
                )
        st.dataframe(pd.DataFrame(exemplos), width="stretch", hide_index=True)
