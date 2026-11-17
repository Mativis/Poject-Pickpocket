# Resolução de conflito de carteira de pedidos

Sistema Streamlit que cruza duas carteiras da Nazareth usando **apenas a descrição do tecido** como chave de ligação. A palavra `TECIDO` é ignorada na comparação, conforme solicitado.

## Executar

```powershell
pip install -r requirements.txt
streamlit run app.py
```

A aplicação abre em `http://localhost:8501`. Os dois arquivos `.xls` de exemplo já são carregados automaticamente quando estão na mesma pasta. Também é possível enviar outras planilhas pela barra lateral.

## O que o sistema faz

| Etapa | Ação |
|---|---|
| Normalização | Remove `TECIDO`, `NEW`, sufixos `/F` e `/P`, acentos e pontuação; zera zeros à esquerda dos códigos (`0001` = `000001` = `1`). |
| Chave exata | Largura + códigos numéricos. O nome do tecido é irrelevante (MADRAS vs MADRAS NEW, AQUILA vs QUILA). |
| Casamento aproximado | Quando a chave exata falha, exige o **mesmo artigo e a mesma cor** antes de aceitar uma similaridade textual. Evita casar cores diferentes por engano. |
| Quantidade | `Demanda líquida = Vendido − Empenhado`, comparada com o `Saldo` da compra. |

## Saídas

- **Conflitos**: pedido sem compra, compra sem pedido, divergência de descrição e saldo insuficiente;
- **Correspondências**: tecidos casados com as quantidades de ambos os lados;
- **Resumo**: contagem por tipo/status e maiores faltas de saldo;
- **Exportação**: Excel com 3 abas e CSVs separados (separador `;`, decode `,` para Excel brasileiro).

## Estrutura

- `app.py` — interface Streamlit;
- `matcher.py` — motor de normalização, casamento e conciliação;
- `requirements.txt` — dependências.

## Resultado nos arquivos de exemplo

Processamento em cerca de 1 segundo: 72 casamentos exatos, 2 divergências de descrição, 6 pedidos sem compra, 49 compras sem pedido e 7 tecidos com saldo insuficiente.
