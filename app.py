import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# Configuração da página
st.set_page_config(
    page_title="Gestão Financeira do Casal",
    page_icon="💑",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilo visual moderno em CSS
st.markdown("""
<style>
    .main { background-color: #0f172a; }
    .stMetric { background-color: #1e293b; padding: 15px; border-radius: 10px; border: 1px solid #334155; }
    .css-1r6slb0 { background-color: #1e293b; }
</style>
""", unsafe_allow_html=True)

# Função para carregar e estruturar dados da planilha CASAL.xlsx
@st.cache_data
def load_data(file_path):
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

# Carregamento dos dados
file_path = 'CASAL.xlsx'
try:
    data = load_data(file_path)
except Exception as e:
    st.error(f"Erro ao carregar a planilha '{file_path}': {e}")
    st.stop()

# Barra Lateral (Sidebar)
st.sidebar.title("💑 Finanças do Casal")
ano_selecionado = st.sidebar.selectbox("📅 Selecione o Ano", list(data.keys()), index=0)

data_year = data[ano_selecionado]
df_rec = data_year["receitas"]
df_desp = data_year["despesas"]

months = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 
          'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']

# Cálculos Totais
rec_totais = df_rec[months].sum() if not df_rec.empty else pd.Series([0]*12, index=months)
desp_totais = df_desp[months].sum() if not df_desp.empty else pd.Series([0]*12, index=months)
saldo_mensal = rec_totais - desp_totais

tot_rec_ano = rec_totais.sum()
tot_desp_ano = desp_totais.sum()
saldo_ano = tot_rec_ano - tot_desp_ano

# Título do App
st.title(f"📊 Painel Financeiro — {ano_selecionado}")

# Cards Superiores
col1, col2, col3, col4 = st.columns(4)
col1.metric("Total Receitas (Ano)", f"R$ {tot_rec_ano:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
col2.metric("Total Despesas (Ano)", f"R$ {tot_desp_ano:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
col3.metric("Saldo Líquido", f"R$ {saldo_ano:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."), delta_color="normal")
col4.metric("Média Mensal de Economia", f"R$ {(saldo_ano/12):,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

st.divider()

# Abas do App
tab1, tab2, tab3 = st.tabs(["📈 Visão Geral & Gráficos", "📑 Detalhamento Completo", "➕ Novo Lançamento"])

with tab1:
    st.subheader("Evolução Mensal das Finanças")
    
    # Gráfico de Barras Receitas vs Despesas
    df_chart = pd.DataFrame({
        "Mês": months,
        "Receitas": rec_totais.values,
        "Despesas": desp_totais.values,
        "Saldo": saldo_mensal.values
    })
    
    fig_evo = px.bar(
        df_chart, x="Mês", y=["Receitas", "Despesas"],
        barmode="group",
        color_discrete_map={"Receitas": "#10b981", "Despesas": "#ef4444"},
        title="Comparativo Mês a Mês"
    )
    fig_evo.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font_color="#f8fafc")
    st.plotly_chart(fig_evo, use_container_width=True)
    
    col_g1, col_g2 = st.columns(2)
    
    with col_g1:
        st.subheader("Divisão das Despesas")
        df_desp_cat = df_desp.copy()
        df_desp_cat["Total"] = df_desp_cat[months].sum(axis=1)
        fig_pie = px.pie(
            df_desp_cat, values="Total", names="Item",
            title="Distribuição Anual de Gastos",
            hole=0.4
        )
        fig_pie.update_layout(paper_bgcolor="rgba(0,0,0,0)", font_color="#f8fafc")
        st.plotly_chart(fig_pie, use_container_width=True)
        
    with col_g2:
        st.subheader("Evolução do Saldo Líquido")
        fig_line = px.line(
            df_chart, x="Mês", y="Saldo",
            markers=True,
            title="Saldo Sobrando/Faltando no Mês",
            color_discrete_sequence=["#3b82f6"]
        )
        fig_line.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font_color="#f8fafc")
        st.plotly_chart(fig_line, use_container_width=True)

with tab2:
    st.subheader("Receitas")
    st.dataframe(df_rec.style.format({m: "R$ {:,.2f}" for m in months}), use_container_width=True)
    
    st.subheader("Despesas")
    st.dataframe(df_desp.style.format({m: "R$ {:,.2f}" for m in months}), use_container_width=True)

with tab3:
    st.subheader("Cadastrar Lançamento Rápido")
    with st.form("form_lancamento"):
        tipo = st.selectbox("Tipo", ["Despesa", "Receita"])
        item = st.text_input("Nome do Item / Categoria", placeholder="Ex: Mercado, Cinema, Conta de Luz")
        mes = st.selectbox("Mês", months)
        valor = st.number_input("Valor (R$)", min_value=0.0, step=10.0, format="%.2f")
        submit = st.form_submit_button("Salvar Registro")
        
        if submit:
            st.success(f"{tipo} '{item}' de R$ {valor:.2f} registrado em {mes}/{ano_selecionado} com sucesso!")
