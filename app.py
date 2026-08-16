import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os

# Configuração inicial da página
st.set_page_config(
    page_title="Gestão Financeira — Casal",
    page_icon="💑",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilização CSS personalizada para um visual escuro e limpo
st.markdown("""
<style>
    /* Card de métricas */
    .metric-card {
        background-color: #1e293b;
        border: 1px solid #334155;
        border-radius: 12px;
        padding: 18px;
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
        font-size: 1.6rem;
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

@st.cache_data
def load_data(file_path):
    if not os.path.exists(file_path):
        st.error(f"Arquivo '{file_path}' não foi encontrado.")
        st.stop()
        
    xls = pd.ExcelFile(file_path)
    months = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 
              'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']
    
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
                
                if any(v > 0 for v in vals.values()):
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

# Menu lateral de opções
st.sidebar.title("💑 Finanças do Casal")
ano_sel = st.sidebar.selectbox("📅 Selecione o Ano", list(data.keys()), index=0)

data_year = data[ano_sel]
df_rec = data_year["receitas"]
df_desp = data_year["despesas"]

months = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 
          'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']

rec_totais = df_rec[months].sum() if not df_rec.empty else pd.Series([0]*12, index=months)
desp_totais = df_desp[months].sum() if not df_desp.empty else pd.Series([0]*12, index=months)
saldo_mensal = rec_totais - desp_totais

tot_rec = rec_totais.sum()
tot_desp = desp_totais.sum()
saldo_anual = tot_rec - tot_desp
pct_comprometido = (tot_desp / tot_rec * 100) if tot_rec > 0 else 0

st.title(f"📊 Painel de Controle — {ano_sel}")

# Bloco de Métricas Principais
col1, col2, col3, col4 = st.columns(4)

with col1:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-title">🟢 Receita Total Anual</div>
        <div class="metric-value val-green">R$ {tot_rec:,.2f}</div>
    </div>
    """.replace(",", "X").replace(".", ",").replace("X", "."), unsafe_allow_html=True)

with col2:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-title">🔴 Despesa Total Anual</div>
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
        <div class="metric-title">⚖️ Renda Comprometida</div>
        <div class="metric-value val-amber">{pct_comprometido:.1f}%</div>
    </div>
    """, unsafe_allow_html=True)

st.write("")
st.divider()

# Navegação em Abas
tab1, tab2, tab3 = st.tabs(["📈 Visão Geral", "👥 Receita do Casal", "📑 Tabela Completa"])

with tab1:
    st.subheader("Evolução Financeira Mês a Mês")
    
    df_chart = pd.DataFrame({
        "Mês": months,
        "Receita": rec_totais.values,
        "Despesa": desp_totais.values,
        "Saldo": saldo_mensal.values
    })
    
    fig_bar = px.bar(
        df_chart, x="Mês", y=["Receita", "Despesa"],
        barmode="group",
        color_discrete_map={"Receita": "#10b981", "Despesa": "#ef4444"},
        labels={"value": "Valor (R$)", "variable": "Tipo"}
    )
    fig_bar.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font_color="#f8fafc")
    st.plotly_chart(fig_bar, use_container_width=True)
    
    c_left, c_right = st.columns(2)
    
    with c_left:
        st.subheader("Para onde está indo o dinheiro?")
        df_desp_sum = df_desp.copy()
        df_desp_sum["Total"] = df_desp_sum[months].sum(axis=1)
        df_desp_sum = df_desp_sum[df_desp_sum["Total"] > 0]
        
        fig_pie = px.pie(
            df_desp_sum, values="Total", names="Item",
            hole=0.4,
            color_discrete_sequence=px.colors.qualitative.Pastel
        )
        fig_pie.update_layout(paper_bgcolor="rgba(0,0,0,0)", font_color="#f8fafc")
        st.plotly_chart(fig_pie, use_container_width=True)
        
    with c_right:
        st.subheader("Economia Sobrando Mês a Mês")
        fig_line = px.line(
            df_chart, x="Mês", y="Saldo",
            markers=True,
            color_discrete_sequence=["#3b82f6"]
        )
        fig_line.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font_color="#f8fafc")
        st.plotly_chart(fig_line, use_container_width=True)

with tab2:
    st.subheader("Divisão de Receitas")
    if not df_rec.empty:
        df_rec_sum = df_rec.copy()
        df_rec_sum["Total"] = df_rec_sum[months].sum(axis=1)
        
        fig_rec_bar = px.bar(
            df_rec_sum, x="Item", y="Total",
            color="Item",
            title="Total Recebido por Origem no Ano",
            text_auto='.2f'
        )
        fig_rec_bar.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font_color="#f8fafc")
        st.plotly_chart(fig_rec_bar, use_container_width=True)

with tab3:
    st.subheader("Tabela de Receitas")
    st.dataframe(df_rec.style.format({m: "R$ {:,.2f}" for m in months}), use_container_width=True)
    
    st.subheader("Tabela de Despesas")
    st.dataframe(df_desp.style.format({m: "R$ {:,.2f}" for m in months}), use_container_width=True)
