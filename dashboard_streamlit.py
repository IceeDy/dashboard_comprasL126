import streamlit as st
import pandas as pd
import plotly.express as px

# Upload do arquivo
st.sidebar.markdown("## 📁 Selecione a planilha de dados")
uploaded_file = st.sidebar.file_uploader(
    "Escolha um arquivo .csv ou .xls/.xlsx",
    type=["csv", "xls", "xlsx"]
)

def to_float(valor):
    if pd.isna(valor):
        return 0.0
    valor = str(valor).replace("R$", "").replace(" ", "").strip()
    if "," in valor:
        valor = valor.replace(".", "").replace(",", ".")
    try:
        return float(valor)
    except Exception:
        return 0.0

def format_brl(value):
    try:
        return f"R$ {float(value):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return f"R$ {value}"

if uploaded_file is not None:
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file, sep=";", dtype=str)
    else:
        df = pd.read_excel(uploaded_file, dtype=str)
    df.columns = df.columns.str.strip()
    if "Categoria" in df.columns:
        df['Categoria'] = df['Categoria'].astype(str).str.strip()

    # Garantir colunas esperadas existem
    expected_cols = [
        "Categoria","Código","Insumo","Necessidade Prof.","Necessidade Aluno","Medida","Estoque",
        "Necessidade Compra","saldo pós compra","Menor Preço","custo","Custo Estoque","total Previsto",
        "Situação","Orçamento 1","Orçamento 2","Orçamento 3","Melhor Preço","Redução Menor Preço",
        "Redução %","Redução R$ unt","Redução R$ total","Qtd Negociada","Valor Total Compra",
        "Valor Total Necessidade","Valor Previsto","Valor Total Histórico","Overstock","Qtd Armazenada",
        "Local","Posição","Fornecedor","Nota Fiscal","Faturado?","Recompra?","Data Compra","Data Entrega","compras"
    ]
    for col in expected_cols:
        if col not in df.columns:
            df[col] = ""

    # Normaliza valores vazios
    # Preenche vazios e infere tipos para evitar aviso de downcasting futuro
    df = df.fillna("").infer_objects(copy=False)
    
    # ---- Novo: detecta coluna de status (N ou Situação) e cria versão normalizada para uso em todo o script
    status_col = "N" if "N" in df.columns else "Situação"
    # garante string e trim
    df[status_col] = df[status_col].astype(str).str.strip()

    # garante que colunas de orçamento existam
    orc_cols = ["Orçamento 1", "Orçamento 2", "Orçamento 3"]
    for c in orc_cols:
        if c not in df.columns:
            df[c] = ""

    # marca como orçado se qualquer célula em Orçamento 1/2/3 tiver valor não vazio
    df["is_orcado"] = df[orc_cols].fillna("").astype(str).apply(lambda row: any(str(v).strip() != "" for v in row), axis=1)

    # se status estiver vazio mas houver orçamento, considerar como "Em Orçamento"
    empty_status_mask = df[status_col] == ""
    df.loc[empty_status_mask & df["is_orcado"], status_col] = "Em Orçamento"

    # versão normalizada para comparações e filtragens
    df["status_norm"] = df[status_col].str.lower()

    # REMOVER totalmente linhas sem status (agora respeita orçamentos)
    df = df[df["status_norm"] != ""].copy()

else:
    st.warning("Por favor, selecione uma planilha para continuar.")
    st.stop()

st.title("📊 Dashboard de Compras - Linha 1 2026")

with st.sidebar:
    st.header("🔍 Filtros")
    search_term = st.text_input("Buscar Insumo")

    if search_term:
        similares = df[df["Insumo"].str.contains(search_term, case=False, na=False)]["Insumo"].unique()
        st.markdown("**Insumos encontrados:**")
        for insumo in similares:
            st.write(insumo)

    selected_category = st.multiselect(
        "Categoria",
        options=df["Categoria"].replace("", "Sem Categoria").unique(),
        default=df["Categoria"].replace("", "Sem Categoria").unique()
    )

    faturamento_options = df['Faturado?'].replace("", "Sem Info").fillna("Sem Info").unique()
    selected_faturamento = st.multiselect(
        "Status Faturamento",
        options=faturamento_options,
        default=faturamento_options
    )

# Converte colunas numéricas essenciais (faz antes de calcular 'qtd')
numeric_cols = [
    "Necessidade Prof.","Necessidade Aluno","Necessidade Compra","Estoque","saldo pós compra",
    "Menor Preço","custo","Custo Estoque","total Previsto","Redução Menor Preço","Redução %",
    "Redução R$ unt","Redução R$ total","Qtd Negociada","Valor Total Compra","Valor Total Necessidade",
    "Valor Previsto","Valor Total Histórico","Overstock","Qtd Armazenada","Melhor Preço","compras"
]
for col in numeric_cols:
    if col in df.columns:
        df.loc[:, col] = df[col].apply(to_float)

# Cria coluna unificada 'qtd' usada por todo o script:
# prioridade: "Necessidade Compra" > "compras" > soma("Necessidade Prof.","Necessidade Aluno") > 0
def compute_qtd_row(row):
    if row.get("Necessidade Compra", 0) and to_float(row.get("Necessidade Compra", 0)) > 0:
        return to_float(row.get("Necessidade Compra", 0))
    if row.get("compras", 0) and to_float(row.get("compras", 0)) > 0:
        return to_float(row.get("compras", 0))
    prof = to_float(row.get("Necessidade Prof.", 0))
    aluno = to_float(row.get("Necessidade Aluno", 0))
    if prof + aluno > 0:
        return prof + aluno
    return 0.0

df["qtd"] = df.apply(compute_qtd_row, axis=1)

# Aplica filtros
filtered_df = df[df["Categoria"].isin(selected_category)].copy()
# evita fillna em Series de dtype object — já substitui valores vazios e garante string
filtered_df.loc[:, "Faturado?"] = filtered_df["Faturado?"].replace("", "Sem Info").astype(str)
filtered_df = filtered_df[filtered_df['Faturado?'].isin(selected_faturamento)].copy()

# também manter a coluna normalizada no filtered_df
filtered_df[status_col] = filtered_df[status_col].astype(str)
filtered_df["status_norm"] = filtered_df[status_col].str.strip().str.lower()

if search_term:
    filtered_df = filtered_df[filtered_df["Insumo"].str.contains(search_term, case=False, na=False)].copy()

# Assegura que as colunas numéricas também existam no filtered_df e cria 'qtd' nele
for col in numeric_cols:
    if col in filtered_df.columns:
        filtered_df.loc[:, col] = filtered_df[col].apply(to_float)
filtered_df["qtd"] = filtered_df.apply(compute_qtd_row, axis=1)

# Recalcula campos financeiros com base em 'qtd' e dados novos
filtered_df.loc[:, "Valor Previsto"] = filtered_df["custo"] * filtered_df["qtd"]
filtered_df.loc[:, "Valor Total Compra"] = filtered_df["Melhor Preço"] * filtered_df["qtd"]
filtered_df.loc[:, "Valor Total Negociado"] = filtered_df["Melhor Preço"] * filtered_df["Qtd Negociada"].fillna(0)
filtered_df.loc[:, "Valor Total Necessidade"] = filtered_df["Menor Preço"] * filtered_df["qtd"]
filtered_df.loc[:, "Valor Total Histórico"] = filtered_df["custo"] * filtered_df["Estoque"].fillna(0)
filtered_df.loc[:, "Overstock"] = filtered_df["Valor Total Negociado"].fillna(0) - filtered_df["Valor Total Compra"].fillna(0)

# Garante tipos float pós-cálculo
for col in ["Valor Previsto","Valor Total Compra","Valor Total Negociado","Valor Total Necessidade","Valor Total Histórico","Overstock"]:
    if col in filtered_df.columns:
        filtered_df.loc[:, col] = filtered_df[col].apply(to_float)

total_negociado = filtered_df["Valor Total Negociado"].sum()
total_compra = filtered_df["Valor Total Compra"].sum()
total_previsto = filtered_df["Valor Previsto"].sum()
total_necessidade = filtered_df["Valor Total Necessidade"].sum()
total_historico = filtered_df["Valor Total Histórico"].sum()
total_economia = total_previsto - total_compra
total_overstock_calculado = (filtered_df["Valor Total Negociado"] - filtered_df["Valor Total Compra"]).sum()
percentual_overstock = (total_overstock_calculado / total_compra) * 100 if total_compra > 0 else 0

# Resumo Executivo
st.markdown("### 📌 Resumo Executivo")
col1, col2 = st.columns(2)
col1.metric("🧾 Total de Itens", len(filtered_df))
col2.metric("💰 Total Compra", format_brl(total_compra))
col3, col4 = st.columns(2)
col3.metric("📊 Total Previsto", format_brl(total_previsto))
col3.metric("📌 Total Necessidade (Menor Preço)", format_brl(total_necessidade))
col4.metric("🎯 Economia Total", format_brl(total_economia))
col4.metric("💰 Total Negociado", format_brl(total_negociado))
col5, col6 = st.columns(2)
col5.metric("📦 Overstock (Negociado - Compra)", format_brl(total_overstock_calculado), f"{percentual_overstock:.1f}%")
col6.metric("🕰️ Valor Histórico (estoque)", format_brl(total_historico))

# Indicadores de Necessidade Prof. e Necessidade Aluno (métricas + gráfico)
# Observação: existem várias medidas (UN, M, PNL etc). Em vez de somar sem contexto,
# mostramos: total bruto por tipo, contagem de itens com necessidade e decomposição por Medida.
need_prof_total = filtered_df["Necessidade Prof."].sum()
need_aluno_total = filtered_df["Necessidade Aluno"].sum()

items_prof_count = int(filtered_df[filtered_df["Necessidade Prof."] > 0].shape[0])
items_aluno_count = int(filtered_df[filtered_df["Necessidade Aluno"] > 0].shape[0])

c1, c2, c3 = st.columns([1,1,1])
c1.metric("👨‍🏫 Necessidade Prof. (total bruto)", f"{need_prof_total:.0f}")
c2.metric("🧑‍🎓 Necessidade Aluno (total bruto)", f"{need_aluno_total:.0f}")
c3.metric("📦 Itens com necessidade (Prof. / Aluno)", f"{items_prof_count} / {items_aluno_count}")

# decomposição por Medida — permite ver onde as somas podem ser comparáveis
need_by_medida = (
    filtered_df.groupby(filtered_df["Medida"].fillna("SEM MEDIDA"))
    .agg({"Necessidade Prof.": "sum", "Necessidade Aluno": "sum"})
    .reset_index()
)
st.markdown("#### Necessidades por Medida (verificar unidades)")
st.dataframe(need_by_medida.sort_values(by=["Medida"]), use_container_width=True, hide_index=True)

# gráfico agrupado por Medida (ajuda a identificar medidas dominantes)
melt_need = need_by_medida.melt(id_vars="Medida", value_vars=["Necessidade Prof.", "Necessidade Aluno"], var_name="Tipo", value_name="Quantidade")
fig_need = px.bar(
    melt_need,
    x="Medida",
    y="Quantidade",
    color="Tipo",
    title="Necessidade por Medida (Prof. vs Aluno)",
    barmode="group"
)
st.plotly_chart(fig_need, use_container_width=True, key="need_by_medida")

# Resumo por Status de Faturamento
st.markdown("### 💳 Resumo por Status de Faturamento")
if 'Faturado?' in filtered_df.columns:
    faturamento_df = filtered_df.groupby('Faturado?').agg({
        'Insumo': 'count',
        'Valor Total Compra': 'sum',
        'Nota Fiscal': lambda x: x.nunique()
    }).rename(columns={
        'Insumo': 'Qtd Itens',
        'Valor Total Compra': 'Total Compra',
        'Nota Fiscal': 'Qtd Notas Fiscais'
    }).reset_index()
    faturamento_df['Total Compra Formatado'] = faturamento_df['Total Compra'].apply(format_brl)
    a1, a2 = st.columns(2)
    with a1:
        st.dataframe(
            faturamento_df[['Faturado?', 'Qtd Itens', 'Qtd Notas Fiscais', 'Total Compra Formatado']],
            use_container_width=True, hide_index=True
        )
    with a2:
        fig_faturamento = px.pie(
            faturamento_df, values='Total Compra', names='Faturado?',
            title='Distribuição por Status de Faturamento', hole=0.4
        )
        st.plotly_chart(fig_faturamento, use_container_width=True, key="fat_chart")

# Status Geral por Situação (usa a coluna detectada dinamicamente)
if status_col in filtered_df.columns:
    st.markdown("### 🧮 Status Geral do Processo de Compras")
    # garante string e versão normalizada (sem fillna em Series)
    filtered_df[status_col] = filtered_df[status_col].astype(str).str.strip()
    filtered_df["status_norm"] = filtered_df[status_col].str.lower()
    # excluir quaisquer linhas sem status (defensivo)
    filtered_df = filtered_df[filtered_df["status_norm"] != ""].copy()

    # --- agora o groupby não criará linha vazia ---
    status_df = (
        filtered_df.groupby(status_col)
        .agg({
            "Insumo": "count",
            "Valor Previsto": "sum",
            "Valor Total Compra": "sum",
            "Valor Total Negociado": "sum"
        })
        .rename(columns={"Insumo": "Qtd. Itens"})
        .reset_index()
    )
    status_df["Valor Previsto"] = status_df["Valor Previsto"].apply(format_brl)
    status_df["Valor Total Compra"] = status_df["Valor Total Compra"].apply(format_brl)
    status_df["Valor Total Negociado"] = status_df["Valor Total Negociado"].apply(format_brl)

    # --- NOVO: indicadores específicos para 'Aguardando' (orçamentos aprovados aguardando entrega) ---
    aguardando_mask = filtered_df["status_norm"] == "aguardando"
    aguardando_count = int(filtered_df.loc[aguardando_mask].shape[0])
    aguardando_valor = filtered_df.loc[aguardando_mask, "Valor Total Compra"].sum() if "Valor Total Compra" in filtered_df.columns else 0.0

    # base para percentual = itens com status em (em orçamento, aguardando, entregue)
    base_mask = filtered_df["status_norm"].isin(["em orçamento", "aguardando", "entregue"])
    base_total = int(filtered_df.loc[base_mask].shape[0])
    aguardando_pct = (aguardando_count / base_total * 100) if base_total > 0 else 0.0

    i1, i2, i3 = st.columns([1,2,1])
    i1.metric("⏳ Itens Aguardando", f"{aguardando_count}")
    i2.metric("💰 Valor Total Aguardando", format_brl(aguardando_valor))
    i3.metric("📊 % Aguardando (base)", f"{aguardando_pct:.1f}%")

    st.dataframe(status_df, use_container_width=True, hide_index=True)

    # Entrega dos itens: base e critérios adaptados aos status: "Em Orçamento", "Aguardando", "Entregue"
    # usar filtered_df para que a métrica responda aos filtros (categoria, faturamento, busca, etc.)
    df_status_base = filtered_df[filtered_df["status_norm"].isin(["em orçamento", "aguardando", "entregue"])].copy()
    df_status_base.loc[:, "Qtd Armazenada"] = df_status_base["Qtd Armazenada"].apply(to_float)

    # Entregue quando status == 'entregue' OU quando há Qtd Armazenada > 0
    mask_entregue = (df_status_base["status_norm"] == "entregue") | (df_status_base["Qtd Armazenada"] > 0)
    entregues_ok = df_status_base[mask_entregue].shape[0]
    total_ok = df_status_base.shape[0]
    percentual_entregue = (entregues_ok / total_ok) * 100 if total_ok > 0 else 0

    s1, s2 = st.columns(2)
    s1.metric("📦 Itens (base: Em Orçamento / Aguardando / Entregue)", f"{total_ok}")
    s2.metric("✅ % Entregue (base acima)", f"{percentual_entregue:.1f}%")

# Materiais Aguardando / Em Orçamento (usa status_norm)
st.markdown("### 🧾 Materiais Aguardando / Em Orçamento")
pendente_norm = ["em orçamento", "aguardando"]
df_status_pendentes = df[df["status_norm"].isin(pendente_norm)].copy()
if df_status_pendentes.empty:
    st.info("Nenhum material com status 'Aguardando' ou 'Em Orçamento' encontrado.")
else:
    df_status_pendentes.loc[:, "Qtd Negociada"] = df_status_pendentes["Qtd Negociada"].apply(to_float)
    df_status_pendentes.loc[:, "Melhor Preço"] = df_status_pendentes["Melhor Preço"].apply(to_float)
    df_status_pendentes.loc[:, "Valor Estimado"] = df_status_pendentes["Qtd Negociada"] * df_status_pendentes["Melhor Preço"]
    df_status_pendentes.loc[:, "Valor Estimado"] = df_status_pendentes["Valor Estimado"].apply(format_brl)
    st.dataframe(
        df_status_pendentes[[
            "Categoria", "Insumo", status_col, "Qtd Negociada", "Medida", "Melhor Preço", "Valor Estimado", "Nota Fiscal", "Faturado?"
        ]],
        use_container_width=True,
        hide_index=True
    )

# Análise de Notas Fiscais por Categoria
st.markdown("### 📝 Análise de Notas Fiscais por Categoria")
if 'Nota Fiscal' in filtered_df.columns:
    nf_df = filtered_df[filtered_df['Nota Fiscal'] != '']
    if not nf_df.empty:
        nf_analysis = nf_df.groupby(['Categoria', 'Nota Fiscal']).agg({
            'Insumo': 'count',
            'Valor Total Compra': 'sum',
            'Faturado?': 'first'
        }).reset_index()
        nf_analysis['Valor Total Compra Formatado'] = nf_analysis['Valor Total Compra'].apply(format_brl)
        st.dataframe(
            nf_analysis.sort_values('Valor Total Compra', ascending=False)[
                ['Categoria', 'Nota Fiscal', 'Insumo', 'Valor Total Compra Formatado', 'Faturado?']
            ],
            use_container_width=True,
            hide_index=True
        )
        fig_nf = px.sunburst(
            nf_analysis,
            path=['Categoria', 'Nota Fiscal'],
            values='Valor Total Compra',
            color='Valor Total Compra',
            title='Distribuição de Notas Fiscais por Categoria e Valor',
            color_continuous_scale='Blues'
        )
        st.plotly_chart(fig_nf, use_container_width=True, key="nf_sunburst")
    else:
        st.info("Nenhuma nota fiscal válida encontrada para análise.")

# Abas de conteúdo
tab1, tab2, tab3 = st.tabs(["📊 Visualizações", "📋 Tabela de Itens", "📈 Análise Avançada"])

with tab1:
    st.subheader("💼 Distribuição de Custos por Categoria")
    pie_data = filtered_df.groupby("Categoria")["Valor Total Compra"].sum().reset_index()
    fig_pie = px.pie(pie_data, values="Valor Total Compra", names="Categoria",
                     title="Distribuição do Investimento por Categoria", hole=0.4)
    st.plotly_chart(fig_pie, use_container_width=True, key="grafico_pizza")

    st.subheader("📦 Overstock por Categoria (Base: Negociado - Compra)")
    overstock_df = filtered_df.groupby("Categoria")[["Valor Total Compra", "Valor Total Negociado"]].sum().reset_index()
    overstock_df["Overstock"] = overstock_df["Valor Total Negociado"] - overstock_df["Valor Total Compra"]
    overstock_df["Valor Total Compra"] = overstock_df["Valor Total Compra"].fillna(0).astype(float)
    overstock_df["Overstock"] = overstock_df["Overstock"].fillna(0).astype(float)
    overstock_df["% Overstock"] = (overstock_df["Overstock"] / overstock_df["Valor Total Compra"].replace({0: pd.NA})) * 100
    overstock_df["% Overstock"] = overstock_df["% Overstock"].fillna(0)
    st.markdown("#### 🔎 Filtrar por categorias com prejuízo")
    show_negative_only = st.checkbox("Mostrar apenas categorias com Overstock negativo")
    if show_negative_only:
        overstock_df = overstock_df[overstock_df["Overstock"] < 0]
    fig_over = px.bar(
        overstock_df,
        x="Categoria",
        y="Overstock",
        color=overstock_df["Overstock"].apply(lambda x: "Acima" if x > 0 else "Abaixo"),
        text=overstock_df["Overstock"].apply(lambda x: format_brl(x)),
        title="Excesso de Compra por Categoria (R$)",
        color_discrete_map={"Acima": "green", "Abaixo": "red"}
    )
    fig_over.update_traces(textposition="outside")
    fig_over.update_layout(yaxis_tickformat=",")
    fig_over.update_yaxes(title_text="Overstock (R$)")
    st.plotly_chart(fig_over, use_container_width=True, key="grafico_overstock")

    st.subheader("💳 Status de Faturamento por Categoria")
    if 'Faturado?' in filtered_df.columns:
        fat_cat_df = filtered_df.groupby(['Categoria', 'Faturado?']).size().reset_index(name='Qtd')
        fig_fat_cat = px.bar(
            fat_cat_df,
            x='Categoria',
            y='Qtd',
            color='Faturado?',
            title='Quantidade de Itens por Status de Faturamento e Categoria',
            barmode='group'
        )
        st.plotly_chart(fig_fat_cat, use_container_width=True, key="fat_cat")

with tab2:
    st.subheader("📋 Lista Detalhada de Insumos")
    if not filtered_df.empty:
        filtered_df.loc[:, "Overstock"] = filtered_df["Valor Total Negociado"] - filtered_df["Valor Total Compra"]
        editable_cols = [
            "Categoria", "Código", "Insumo", "Necessidade Prof.", "Necessidade Aluno", "qtd", "Necessidade Compra",
            "Medida", "Estoque", "saldo pós compra", "Menor Preço", "custo", "Custo Estoque", "total Previsto",
            "Situação", "Orçamento 1", "Orçamento 2", "Orçamento 3", "Melhor Preço",
            "Redução Menor Preço", "Redução %", "Redução R$ unt", "Redução R$ total",
            "Qtd Negociada", "Valor Total Compra", "Valor Total Necessidade", "Valor Previsto",
            "Valor Total Histórico", "Overstock", "Qtd Armazenada", "Local", "Posição", "Fornecedor",
            "Nota Fiscal", "Faturado?", "Recompra?", "Data Compra", "Data Entrega"
        ]
        available_cols = [c for c in editable_cols if c in filtered_df.columns]
        rename_map = {
            "custo": "Valor Unitário Antigo",
            "Melhor Preço": "Valor Unitário Atual",
            "qtd": "qtd"
        }
        display_df = filtered_df[available_cols].rename(columns={k: v for k, v in rename_map.items() if k in available_cols})
        edited_df = st.data_editor(display_df, use_container_width=True, num_rows="dynamic", key="insumos_editor")

        def compute_qty_series(mask):
            # prioridade: "Necessidade Compra" > "compras" > soma(Necessidade Prof., Necessidade Aluno) > qtd fallback
            if "Necessidade Compra" in df.columns and df.loc[mask, "Necessidade Compra"].sum() > 0:
                return df.loc[mask, "Necessidade Compra"].fillna(0)
            if "compras" in df.columns and df.loc[mask, "compras"].sum() > 0:
                return df.loc[mask, "compras"].fillna(0)
            if "Necessidade Prof." in df.columns and "Necessidade Aluno" in df.columns:
                return df.loc[mask, "Necessidade Prof."].fillna(0) + df.loc[mask, "Necessidade Aluno"].fillna(0)
            return df.loc[mask, "qtd"].fillna(0)

        if st.button("💾 Salvar Alterações"):
            for idx, row in edited_df.iterrows():
                mask = pd.Series(True, index=df.index)
                for key in ["Categoria", "Código", "Insumo", "Medida", "Necessidade Compra", "Necessidade Prof.", "Necessidade Aluno", "qtd"]:
                    if key in df.columns and key in row.index:
                        mask &= df[key].astype(str) == str(row[key])

                if "Valor Unitário Antigo" in row.index:
                    df.loc[mask, "custo"] = row["Valor Unitário Antigo"]
                elif "custo" in row.index:
                    df.loc[mask, "custo"] = row["custo"]

                if "Valor Unitário Atual" in row.index:
                    df.loc[mask, "Melhor Preço"] = row["Valor Unitário Atual"]
                elif "Melhor Preço" in row.index:
                    df.loc[mask, "Melhor Preço"] = row["Melhor Preço"]

                if "Nota Fiscal" in row.index and "Nota Fiscal" in df.columns:
                    df.loc[mask, "Nota Fiscal"] = row["Nota Fiscal"]
                if "Faturado?" in row.index and "Faturado?" in df.columns:
                    df.loc[mask, "Faturado?"] = row["Faturado?"]

                try:
                    qty_series = compute_qty_series(mask)
                    unit_atual = to_float(row.get("Valor Unitário Atual", row.get("Melhor Preço", 0)))
                    unit_antigo = to_float(row.get("Valor Unitário Antigo", row.get("custo", 0)))
                    df.loc[mask, "Valor Total Compra"] = unit_atual * qty_series
                    df.loc[mask, "Valor Previsto"] = unit_antigo * qty_series
                    if "Qtd Negociada" in df.columns:
                        df.loc[mask, "Valor Total Negociado"] = unit_atual * df.loc[mask, "Qtd Negociada"].fillna(0)
                    # atualiza coluna unificada qtd após edição de necessidade
                    df.loc[mask, "qtd"] = df.loc[mask].apply(compute_qtd_row, axis=1)
                except Exception:
                    pass

            df.to_csv("dashboard_data.csv", sep=";", index=False)
            st.success("Alterações salvas com sucesso!")
    else:
        st.warning("Nenhum dado encontrado com os filtros aplicados.")

with tab3:
    st.subheader("📈 Ranking de Categorias - Análise Avançada")
    if not filtered_df.empty:
        rank_df = filtered_df.groupby ("Categoria").agg({
            "Valor Total Compra": "sum",
            "Valor Total Negociado": "sum",
            "Valor Previsto": "sum",
            "Nota Fiscal": "nunique"
        }).reset_index()
        rank_df["Economia"] = rank_df["Valor Previsto"] - rank_df["Valor Total Negociado"]
        rank_df["Overstock"] = rank_df["Valor Total Negociado"] - rank_df["Valor Total Compra"]
        rank_df = rank_df.rename(columns={"Nota Fiscal": "Qtd Notas Fiscais"})
        styled_rank_df = rank_df.copy()
        for col in ["Valor Total Compra", "Valor Total Negociado", "Valor Previsto", "Economia", "Overstock"]:
            styled_rank_df[col] = styled_rank_df[col].apply(format_brl)
        st.markdown("### 🏆 Maiores Economias")
        st.dataframe(
            styled_rank_df.sort_values(by="Economia", ascending=False)
            .style.map(lambda x: "color: green" if "-" not in str(x) else "color: red", subset=["Economia"]),
            use_container_width=True,
            hide_index=True
        )
        st.markdown("### 💸 Categorias com Maior Investimento")
        st.dataframe(styled_rank_df.sort_values(by="Valor Total Compra", ascending=False), use_container_width=True, hide_index=True)

        st.markdown("### 💰 Status de Faturamento por Categoria")
        if 'Faturado?' in filtered_df.columns:
            fat_cat_analysis = filtered_df.groupby(['Categoria','Faturado?']).agg({
                'Valor Total Compra':'sum','Insumo':'count'
            }).reset_index()
            fig_fat_cat_value = px.bar(fat_cat_analysis, x='Categoria', y='Valor Total Compra', color='Faturado?', title='Valor Total por Status de Faturamento e Categoria', barmode='stack')
            st.plotly_chart(fig_fat_cat_value, use_container_width=True, key="fat_cat_value")
    else:
        st.warning("Sem dados disponíveis para análise avançada.")

# Status de entrega detalhado (usa 'qtd' para exibição)
st.markdown("### 📦 Status de Entrega dos Itens (OK)")
df.loc[:, "Qtd Armazenada"] = df["Qtd Armazenada"].apply(to_float)
df.loc[:, "Situação"] = df["Situação"].astype(str)

# Na seção de exibição detalhada (Entregues / Aguardando) usamos status_norm para decidir
df_ok = df[df["status_norm"].isin(["entregue", "aguardando", "em orçamento"])].copy()
df_ok.loc[:, "Status Entrega"] = df_ok.apply(
    lambda r: "Entregue" if (r["status_norm"] == "entregue" or to_float(r.get("Qtd Armazenada", 0)) > 0) else "Aguardando Entrega",
    axis=1
)

# Entregues: todos que foram marcados como "Entregue"
entregues = df_ok[df_ok["Status Entrega"] == "Entregue"].copy()

# Aguardando: somente os que têm status_norm == "aguardando" (orçamentos aprovados aguardando entrega)
aguardando = df_ok[(df_ok["Status Entrega"] == "Aguardando Entrega") & (df_ok["status_norm"] == "aguardando")].copy()

# Exibe colunas (usa 'qtd' unificada)
col_entregue, col_aguardando = st.columns(2)
with col_entregue:
    st.markdown("#### ✅ Entregues")
    qty_options = ["qtd", "Necessidade Compra", "Necessidade Prof.", "Necessidade Aluno", "compras"]
    qty_col = next((c for c in qty_options if c in entregues.columns), None)
    cols_entrega = ["Categoria", "Código", "Insumo"]
    if qty_col:
        cols_entrega.append(qty_col)
    cols_entrega += ["Medida", "Qtd Armazenada", "Status Entrega", status_col, "Nota Fiscal", "Faturado?"]
    cols_entrega = [c for c in cols_entrega if c in entregues.columns]
    st.dataframe(entregues[cols_entrega], use_container_width=True, hide_index=True)

with col_aguardando:
    st.markdown("#### ⏳ Aguardando Entrega")
    qty_col2 = next((c for c in qty_options if c in aguardando.columns), None)
    cols_agu = ["Categoria", "Código", "Insumo"]
    if qty_col2:
        cols_agu.append(qty_col2)
    cols_agu += ["Medida", "Qtd Armazenada", "Status Entrega", status_col, "Nota Fiscal", "Faturado?"]
    cols_agu = [c for c in cols_agu if c in aguardando.columns]
    st.dataframe(aguardando[cols_agu], use_container_width=True, hide_index=True)

st.markdown("""
---
Desenvolvido por Leon – 2026 | Interface aprimorada com foco em clareza e desempenho visual.
""")

# --- Relatório resumido (tela + download CSV / Excel) ---
summary = {
    "Total Itens": len(filtered_df),
    "Total Compra (R$)": total_compra,
    "Total Negociado (R$)": total_negociado,
    "Total Previsto (R$)": total_previsto,
    "Total Necessidade (R$)": total_necessidade,
    "Valor Histórico (R$)": total_historico,
    "Economia Total (R$)": total_economia,
    "Total Overstock (R$)": total_overstock_calculado,
    "Necessidade Prof. (unidades)": need_prof_total,
    "Necessidade Aluno (unidades)": need_aluno_total,
    "Itens Aguardando (count)": aguardando_count,
    "Valor Aguardando (R$)": aguardando_valor,
    "% Aguardando (base)": aguardando_pct,
    "Itens Base Entrega": base_total,
    "Itens Entregues": entregues_ok,
    "% Entregue (base)": percentual_entregue
}

# DataFrame legível
summary_df = pd.DataFrame.from_dict(summary, orient="index", columns=["Valor"])
# mostra na tela
st.markdown("### 🧾 Relatório Resumido")
st.dataframe(summary_df, use_container_width=True, hide_index=False)

# prepara downloads (CSV)
csv_bytes = summary_df.to_csv(sep=";", encoding="utf-8-sig").encode("utf-8-sig")
st.download_button(
    label="⬇️ Baixar Relatório Resumido (CSV)",
    data=csv_bytes,
    file_name="relatorio_resumido.csv",
    mime="text/csv"
)

# prepara downloads (Excel) com folha 'Resumo' + opcional 'Detalhado' com filtered_df
from io import BytesIO
output = BytesIO()

# tenta escolher um engine Excel disponível: prefer xlsxwriter, fallback para openpyxl
import importlib.util
engine = None
if importlib.util.find_spec("xlsxwriter") is not None:
    engine = "xlsxwriter"
elif importlib.util.find_spec("openpyxl") is not None:
    engine = "openpyxl"
else:
    engine = None

if engine:
    with pd.ExcelWriter(output, engine=engine) as writer:
        # salva resumo
        summary_df.to_excel(writer, sheet_name="Resumo")
        # salva aba detalhada (opcional: limitar colunas)
        try:
            filtered_df.to_excel(writer, sheet_name="Detalhado", index=False)
        except Exception:
            pass
    excel_data = output.getvalue()
    st.download_button(
        label="⬇️ Baixar Relatório Resumido (Excel)",
        data=excel_data,
        file_name="relatorio_resumido.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.warning("Não foi possível gerar XLSX — instale 'XlsxWriter' ou 'openpyxl' (pip install XlsxWriter openpyxl) para habilitar o download em Excel.")

# opcional: botão para salvar snapshot no servidor (somente se desejar)
if st.button("💾 Salvar relatório no servidor"):
    summary_df.to_csv("relatorio_resumido_snapshot.csv", sep=";", encoding="utf-8-sig")
    filtered_df.to_csv("relatorio_detalhado_snapshot.csv", sep=";", index=False, encoding="utf-8-sig")
    st.success("Snapshots salvos no diretório do app.")