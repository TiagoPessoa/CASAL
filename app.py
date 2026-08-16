import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
import os

# Configuração da página
st.set_page_config(
    page_title="Gestão Financeira — Casal",
    page_icon="💑",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilização CSS
st.markdown("""
<style>
    .metric-card {
        background-color: #1e293b;
        border: 1px solid #334155;
        border-radius: 12px;
        padding: 16px;
        text-align: center;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
    }
    .metric-title {
        font-size: 0.85rem;
        color: #94a3b8;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        margin-bottom: 6px;
    }
    .metric-value {
        font-size: 1.5rem;
        font-weight: 700;
    }
    .val-green { color: #10b981; }
    .val-red { color: #ef4444; }
    .val-blue { color: #3b82f6; }
    .val-amber { color: #f59e0b; }
</style>
""", unsafe_allow_html=True)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FILE_PATH = os.path.join(BASE_DIR, 'CASAL.xlsx')

months = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 
          'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']

@st.cache_data
def load_data(file_path):
    if not os.path.exists(file_path):
        st.error(f"Arquivo '{file_path}' não foi encontrado.")
        st.stop()
        
    xls = pd.ExcelFile(file_path)
    data_by_year = {}
    
    for sheet in xls.sheet_names:
        if sheet in ['2026', '2027']:
            df = pd.read_excel(file_path, sheet_name=sheet)
            cat_col = df.columns[0]
            
            receitas = []
            despesas = []
            in_despesas = False
            
            for idx, row in df.iterrows():
                item_name = str(row[cat_col]).strip() if pd.notna(row[cat_col]) else ''
                if not item_name or item_name == 'nan':
                    continue
                
                if item_name.lower() == 'despesas':
                    in_despesas = True
                    continue
                
                if any(x in item_name.lower() for x in ['total', 'saldo', 'reserva', 'acumulado']):
                    continue
                
                vals = {}
                for m in months:
                    col_m = m if m in df.columns else (m + ' ' if (m + ' ') in df.columns else None)
                    if col_m and pd.notna(row[col_m]):
                        try:
                            vals[m] = float(row[col_m])
                        except:
                            vals[m] = 0.0
                    else:
                        vals[m] = 0.0
                
                entry = {"Item": item_name, **vals}
                if in_despesas:
                    despesas.append(entry)
                else:
                    receitas.append(entry)
                        
            data_by_year[sheet] = {
                "receitas": pd.DataFrame(receitas),
                "despesas": pd.DataFrame(despesas)
            }
            
    return data_by_year

try:
    data = load_data(FILE_PATH)
except Exception as e:
    st.error(f"Erro ao carregar os dados: {e}")
    st.stop()

# Menu lateral
st.sidebar.title("💑 Finanças do Casal")
ano_sel = st.sidebar.selectbox("📅 Selecione o Ano", list(data.keys()), index=0)

data_year = data[ano_sel]
df_rec = data_year["receitas"]
df_desp = data_year["despesas"]

# Mês de foco
mes_atual_idx = datetime.now().month - 1
mes_foco = st.sidebar.selectbox("📌 Mês de Foco Inicial", months, index=mes_atual_idx)
idx_foco = months.index(mes_foco)
mes_foco_prox = months[(idx_foco + 1) % 12]

rec_totais = df_rec[months].sum() if not df_rec.empty else pd.Series([0]*12, index=months)
desp_totais = df_desp[months].sum() if not df_desp.empty else pd.Series([0]*12, index=months)
saldo_mensal = rec_totais - desp_totais

tot_rec = rec_totais.sum()
tot_desp = desp_totais.sum()
saldo_anual = tot_rec - tot_desp

st.title(f"📊 Painel de Controle — {ano_sel}")

# Métricas do Ano
col1, col2, col3, col4 = st.columns(4)

with col1:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-title">🟢 Receita Anual</div>
        <div class="metric-value val-green">R$ {tot_rec:,.2f}</div>
    </div>
    """.replace(",", "X").replace(".", ",").replace("X", "."), unsafe_allow_html=True)

with col2:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-title">🔴 Despesa Anual</div>
        <div class="metric-value val-red">R$ {tot_desp:,.2f}</div>
    </div>
    """.replace(",", "X").replace(".", ",").replace("X", "."), unsafe_allow_html=True)

with col3:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-title">🔵 Saldo Anual Líquido</div>
        <div class="metric-value val-blue">R$ {saldo_anual:,.2f}</div>
    </div>
    """.replace(",", "X").replace(".", ",").replace("X", "."), unsafe_allow_html=True)

with col4:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-title">📌 Foco Atual</div>
        <div class="metric-value val-amber">{mes_foco} / {mes_foco_prox}</div>
    </div>
    """, unsafe_allow_html=True)

st.write("")
st.divider()

# Abas do aplicativo
tab_foco, tab_add, tab_editar, tab_geral, tab_completa = st.tabs([
    "📌 Mês Atual & Próximo", 
    "➕ Novo Lançamento",
    "✏️ Editar Planilha",
    "📈 Visão Geral", 
    "📑 Tabela Completa 12M"
])

# ----------------- ABA 1: FOCO NO MÊS ATUAL E PRÓXIMO -----------------
with tab_foco:
    st.subheader(f"🔍 Comparativo Prático: {mes_foco} vs. {mes_foco_prox}")
    
    c_m1, c_m2 = st.columns(2)
    
    rec_m1 = rec_totais[mes_foco]
    desp_m1 = desp_totais[mes_foco]
    saldo_m1 = rec_m1 - desp_m1
    
    rec_m2 = rec_totais[mes_foco_prox]
    desp_m2 = desp_totais[mes_foco_prox]
    saldo_m2 = rec_m2 - desp_m2
    
    with c_m1:
        st.markdown(f"### 🗓️ {mes_foco}")
        st.write(f"🟢 **Receitas:** R$ {rec_m1:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        st.write(f"🔴 **Despesas:** R$ {desp_m1:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        st.write(f"🔵 **Saldo:** R$ {saldo_m1:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        
    with c_m2:
        st.markdown(f"### 🗓️ {mes_foco_prox}")
        st.write(f"🟢 **Receitas:** R$ {rec_m2:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        st.write(f"🔴 **Despesas:** R$ {desp_m2:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        st.write(f"🔵 **Saldo:** R$ {saldo_m2:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

    st.divider()
    
    col_t1, col_t2 = st.columns(2)
    
    with col_t1:
        st.markdown("#### 🟢 Receitas")
        if not df_rec.empty:
            df_rec_foco = df_rec[["Item", mes_foco, mes_foco_prox]].copy()
            df_rec_foco = df_rec_foco[(df_rec_foco[mes_foco] > 0) | (df_rec_foco[mes_foco_prox] > 0)]
            st.dataframe(
                df_rec_foco.style.format({mes_foco: "R$ {:,.2f}", mes_foco_prox: "R$ {:,.2f}"}),
                use_container_width=True
            )
            
    with col_t2:
        st.markdown("#### 🔴 Despesas")
        if not df_desp.empty:
            df_desp_foco = df_desp[["Item", mes_foco, mes_foco_prox]].copy()
            df_desp_foco = df_desp_foco[(df_desp_foco[mes_foco] > 0) | (df_desp_foco[mes_foco_prox] > 0)]
            st.dataframe(
                df_desp_foco.style.format({mes_foco: "R$ {:,.2f}", mes_foco_prox: "R$ {:,.2f}"}),
                use_container_width=True
            )

# ----------------- ABA 2: NOVO LANÇAMENTO RÁPIDO -----------------
with tab_add:
    st.subheader(f"➕ Adicionar Lançamento Rápido em {ano_sel}")
    
    with st.form("form_novo_lancamento"):
        tipo_l = st.radio("Tipo de Lançamento:", ["🔴 Despesa", "🟢 Receita"], horizontal=True)
        
        target_df = df_desp if "Despesa" in tipo_l else df_rec
        itens_existentes = list(target_df["Item"].unique()) if not target_df.empty else []
        
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            item_sel = st.selectbox("Selecione o Item / Categoria Existente:", ["-- Criar Novo Item --"] + itens_existentes)
            if item_sel == "-- Criar Novo Item --":
                novo_item_nome = st.text_input("Nome do Novo Item:", placeholder="Ex: Farmácia, Cinema")
            else:
                novo_item_nome = item_sel
                
        with col_f2:
            mes_l = st.selectbox("Mês do Valor:", months, index=mes_atual_idx)
            valor_l = st.number_input("Valor (R$):", min_value=0.0, step=10.0, format="%.2f")
            
        btn_salvar_l = st.form_submit_button("💾 Salvar Lançamento", type="primary")
        
        if btn_salvar_l:
            if not novo_item_nome or novo_item_nome == "-- Criar Novo Item --":
                st.error("Por favor, digite ou selecione um nome de item válido!")
            elif valor_l <= 0:
                st.error("Digite um valor maior que R$ 0,00!")
            else:
                df_alvo = df_desp if "Despesa" in tipo_l else df_rec
                
                # Se o item já existir, atualiza/soma o valor do mês
                if novo_item_nome in df_alvo["Item"].values:
                    idx = df_alvo[df_alvo["Item"] == novo_item_nome].index[0]
                    df_alvo.loc[idx, mes_l] += valor_l
                else:
                    # Se for item novo, cria nova linha
                    nova_linha = {"Item": novo_item_nome}
                    for m in months:
                        nova_linha[m] = valor_l if m == mes_l else 0.0
                    if "Despesa" in tipo_l:
                        data[ano_sel]["despesas"] = pd.concat([df_desp, pd.DataFrame([nova_linha])], ignore_index=True)
                    else:
                        data[ano_sel]["receitas"] = pd.concat([df_rec, pd.DataFrame([nova_linha])], ignore_index=True)
                        
                st.cache_data.clear()
                st.success(f"Lançamento de R$ {valor_l:.2f} registrado em '{novo_item_nome}' ({mes_l}/{ano_sel})!")
                st.rerun()

# ----------------- ABA 3: EDITAR VALORES E CATEGORIAS -----------------
with tab_editar:
    st.subheader(f"✏️ Edição Direta da Planilha ({ano_sel})")
    st.info("💡 Altere qualquer valor nas tabelas abaixo e clique no botão verde no final para salvar.")

    st.markdown("### 🟢 Editar Receitas")
    df_rec_edit = st.data_editor(
        df_rec, 
        num_rows="dynamic",
        use_container_width=True,
        key=f"editor_rec_{ano_sel}"
    )

    st.markdown("### 🔴 Editar Despesas")
    df_desp_edit = st.data_editor(
        df_desp, 
        num_rows="dynamic",
        use_container_width=True,
        key=f"editor_desp_{ano_sel}"
    )

    if st.button("💾 Salvar Alterações da Tabela", type="primary"):
        data[ano_sel]["receitas"] = df_rec_edit
        data[ano_sel]["despesas"] = df_desp_edit
        st.cache_data.clear()
        st.success("Tabelas atualizadas com sucesso!")
        st.rerun()

# ----------------- ABA 4: VISÃO GERAL ANUAL -----------------
with tab_geral:
    st.subheader("Evolução Mensal das Finanças")
    
    df_chart = pd.DataFrame({
        "Mês": months,
        "Receita": rec_totais.values,
        "Despesa": desp_totais.values,
        "Saldo": saldo_mensal.values
    })
    
    fig_bar = px.bar(
        df_chart, x="Mês", y=["Receita", "Despesa"],
        barmode="group",
        color_discrete_map={"Receita": "#10b981", "Despesa": "#ef4444"}
    )
    fig_bar.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font_color="#f8fafc")
    st.plotly_chart(fig_bar, use_container_width=True)

# ----------------- ABA 5: TABELA COMPLETA -----------------
with tab_completa:
    st.subheader("Tabela Completa de Receitas (12 Meses)")
    st.dataframe(df_rec.style.format({m: "R$ {:,.2f}" for m in months}), use_container_width=True)
    
    st.subheader("Tabela Completa de Despesas (12 Meses)")
    st.dataframe(df_desp.style.format({m: "R$ {:,.2f}" for m in months}), use_container_width=True)
