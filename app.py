import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
from supabase import create_client, Client
import os

# 1. Configuração Inicial da Página
st.set_page_config(
    page_title="Gestão Financeira — Casal",
    page_icon="💑",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 2. Conexão com Supabase
@st.cache_resource
def init_supabase():
    url = st.secrets.get("SUPABASE_URL", "")
    key = st.secrets.get("SUPABASE_KEY", "")
    if url and key:
        try:
            return create_client(url, key)
        except Exception as e:
            st.error(f"Erro ao conectar com Supabase: {e}")
            return None
    return None

supabase = init_supabase()

# 3. Estilização CSS
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
    .metric-title { font-size: 0.85rem; color: #94a3b8; text-transform: uppercase; letter-spacing: 0.05em; margin-bottom: 6px; }
    .metric-value { font-size: 1.5rem; font-weight: 700; }
    .val-green { color: #10b981; } 
    .val-red { color: #ef4444; }
    .val-blue { color: #3b82f6; } 
    .val-amber { color: #f59e0b; }
    .alert-box { padding: 12px 16px; border-radius: 8px; font-weight: 600; margin-bottom: 15px; }
    .alert-warning { background-color: rgba(245, 158, 11, 0.2); border: 1px solid #f59e0b; color: #fef08a; }
    .alert-danger { background-color: rgba(239, 68, 68, 0.2); border: 1px solid #ef4444; color: #fca5a5; }
    .alert-success { background-color: rgba(16, 185, 129, 0.2); border: 1px solid #10b981; color: #6ee7b7; }
</style>
""", unsafe_allow_html=True)

months = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 
          'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']

CATEGORY_MAP = {
    'CAERN': '🏠 Moradia & Contas', 'COSERN': '🏠 Moradia & Contas', 'INTERNET': '🏠 Moradia & Contas',
    'IPTV': '🏠 Moradia & Contas', 'RASTREADOR CARRO': '🏠 Moradia & Contas', 'RASTREADOR MOTO': '🏠 Moradia & Contas',
    'VIVO TIAGO': '🏠 Moradia & Contas', 'VIVO ESTEFANNY': '🏠 Moradia & Contas',
    'CARRO': '🚗 Veículos & Transporte', 'MOTO': '🚗 Veículos & Transporte', 'COMBUSTIVEL': '🚗 Veículos & Transporte',
    'ALIMENTAÇÃO': '🛒 Estilo de Vida & Saúde', 'PANOBIANCO': '🛒 Estilo de Vida & Saúde',
    'WELHUB': '🛒 Estilo de Vida & Saúde', 'BOTICARIO': '🛒 Estilo de Vida & Saúde', 'RACÃO': '🛒 Estilo de Vida & Saúde',
    'CARTÃO TIAGO': '💳 Cartões & Empréstimos', 'CARTÃO ESTEFANNY': '💳 Cartões & Empréstimos',
    'EMPRESTIMOS': '💳 Cartões & Empréstimos', 'OUTROS': '📦 Diversos'
}

def get_category(item_name):
    item_upper = str(item_name).upper().strip()
    for k, v in CATEGORY_MAP.items():
        if k in item_upper:
            return v
    return '📦 Diversos'

# 4. Leitura do Banco Supabase
@st.cache_data(ttl=5)
def load_data():
    if supabase:
        try:
            res_rec = supabase.table("receitas").select("*").execute()
            res_desp = supabase.table("despesas").select("*").execute()
            
            df_r = pd.DataFrame(res_rec.data) if res_rec.data else pd.DataFrame()
            df_d = pd.DataFrame(res_desp.data) if res_desp.data else pd.DataFrame()
            
            for m in months:
                col_key = m.lower()
                if not df_r.empty and col_key not in df_r.columns: df_r[col_key] = 0.0
                if not df_d.empty and col_key not in df_d.columns: df_d[col_key] = 0.0
            
            if not df_r.empty or not df_d.empty:
                rename_dict = {m.lower(): m for m in months}
                if not df_r.empty: 
                    df_r = df_r.rename(columns=rename_dict)
                    if 'item' in df_r.columns: df_r = df_r.rename(columns={'item': 'Item'})
                if not df_d.empty: 
                    df_d = df_d.rename(columns=rename_dict)
                    if 'item' in df_d.columns: df_d = df_d.rename(columns={'item': 'Item'})
                    if 'categoria' in df_d.columns: df_d = df_d.rename(columns={'categoria': 'Categoria'})
                
                data_by_year = {}
                for yr in ['2026', '2027']:
                    r_yr = df_r[df_r['ano'] == yr].copy() if ('ano' in df_r.columns and not df_r.empty) else pd.DataFrame(columns=['Item'] + months)
                    d_yr = df_d[df_d['ano'] == yr].copy() if ('ano' in df_d.columns and not df_d.empty) else pd.DataFrame(columns=['Item', 'Categoria'] + months)
                    data_by_year[yr] = {"receitas": r_yr, "despesas": d_yr}
                return data_by_year
        except Exception as e:
            st.error(f"Erro ao ler do Supabase: {e}")

    file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'CASAL.xlsx')
    if not os.path.exists(file_path):
        return {'2026': {"receitas": pd.DataFrame(), "despesas": pd.DataFrame()}}
    xls = pd.ExcelFile(file_path)
    data_by_year = {}
    for sheet in xls.sheet_names:
        if sheet in ['2026', '2027']:
            df = pd.read_excel(file_path, sheet_name=sheet)
            cat_col = df.columns[0]
            receitas, despesas, in_despesas = [], [], False
            for idx, row in df.iterrows():
                item_name = str(row[cat_col]).strip() if pd.notna(row[cat_col]) else ''
                if not item_name or item_name == 'nan': continue
                if item_name.lower() == 'despesas': in_despesas = True; continue
                if any(x in item_name.lower() for x in ['total', 'saldo', 'reserva', 'acumulado']): continue
                vals = {m: (float(row[m]) if m in df.columns and pd.notna(row[m]) else 0.0) for m in months}
                entry = {"Item": item_name, "ano": sheet, **vals}
                if in_despesas:
                    entry["Categoria"] = get_category(item_name)
                    despesas.append(entry)
                else:
                    receitas.append(entry)
            data_by_year[sheet] = {"receitas": pd.DataFrame(receitas), "despesas": pd.DataFrame(despesas)}
    return data_by_year

data = load_data()

# 5. Barra Lateral
st.sidebar.title("💑 Finanças do Casal")
ano_sel = st.sidebar.selectbox("📅 Selecione o Ano", list(data.keys()) if data else ['2026'], index=0)

data_year = data.get(ano_sel, {"receitas": pd.DataFrame(), "despesas": pd.DataFrame()})
df_rec = data_year["receitas"]
df_desp = data_year["despesas"]

mes_atual_idx = datetime.now().month - 1
mes_atual_nome = months[mes_atual_idx]
mes_prox_nome = months[(mes_atual_idx + 1) % 12]

with st.sidebar.expander("📱 Atalho no Celular"):
    st.write("• **iPhone:** Compartilhar -> Adicionar à Tela de Início\n• **Android:** Menu 3 Pontos -> Adicionar à Tela Inicial")

# 6. Processamento dos Totais
rec_totais = df_rec[months].sum() if not df_rec.empty and all(m in df_rec.columns for m in months) else pd.Series([0]*12, index=months)
desp_totais = df_desp[months].sum() if not df_desp.empty and all(m in df_desp.columns for m in months) else pd.Series([0]*12, index=months)
saldo_mensal = rec_totais - desp_totais

tot_rec, tot_desp = rec_totais.sum(), desp_totais.sum()
saldo_anual = tot_rec - tot_desp
pct_comp_ano = (tot_desp / tot_rec * 100) if tot_rec > 0 else 0

st.title(f"📊 Painel de Controle — {ano_sel}")

# Cards Superiores
col1, col2, col3, col4 = st.columns(4)
col1.markdown(f'<div class="metric-card"><div class="metric-title">🟢 Receita Anual</div><div class="metric-value val-green">R$ {tot_rec:,.2f}</div></div>'.replace(",", "X").replace(".", ",").replace("X", "."), unsafe_allow_html=True)
col2.markdown(f'<div class="metric-card"><div class="metric-title">🔴 Despesa Anual</div><div class="metric-value val-red">R$ {tot_desp:,.2f}</div></div>'.replace(",", "X").replace(".", ",").replace("X", "."), unsafe_allow_html=True)
col3.markdown(f'<div class="metric-card"><div class="metric-title">🔵 Saldo Anual Líquido</div><div class="metric-value val-blue">R$ {saldo_anual:,.2f}</div></div>'.replace(",", "X").replace(".", ",").replace("X", "."), unsafe_allow_html=True)
color_class = "val-green" if pct_comp_ano <= 75 else ("val-amber" if pct_comp_ano <= 90 else "val-red")
col4.markdown(f'<div class="metric-card"><div class="metric-title">⚖️ Renda Comprometida</div><div class="metric-value {color_class}">{pct_comp_ano:.1f}%</div></div>', unsafe_allow_html=True)

st.write("")

# Abas da Aplicação
tab_foco, tab_grupo, tab_add, tab_editar, tab_geral, tab_completa = st.tabs([
    "📌 Mês Atual & Próximo", "🏷️ Categorias & Metas", "➕ Novo Lançamento", "✏️ Editar Planilha", "📈 Visão Geral Anual", "📑 Tabela Completa 12M"
])

# ----------------- ABA 1: MÊS ATUAL & PRÓXIMO (FIXO) -----------------
with tab_foco:
    st.subheader(f"🔍 Comparativo Prático: {mes_atual_nome} vs. {mes_prox_nome}")
    
    c_m1, c_m2 = st.columns(2)
    
    rec_m1 = rec_totais[mes_atual_nome]
    desp_m1 = desp_totais[mes_atual_nome]
    
    rec_m2 = rec_totais[mes_prox_nome]
    desp_m2 = desp_totais[mes_prox_nome]
    
    with c_m1:
        st.markdown(f"### 🗓️ {mes_atual_nome} (Atual)")
        st.write(f"🟢 **Receitas:** R$ {rec_m1:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        st.write(f"🔴 **Despesas:** R$ {desp_m1:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        st.write(f"🔵 **Saldo:** R$ {rec_m1 - desp_m1:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        
    with c_m2:
        st.markdown(f"### 🗓️ {mes_prox_nome} (Próximo)")
        st.write(f"🟢 **Receitas:** R$ {rec_m2:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        st.write(f"🔴 **Despesas:** R$ {desp_m2:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        st.write(f"🔵 **Saldo:** R$ {rec_m2 - desp_m2:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

    st.divider()
    
    col_t1, col_t2 = st.columns(2)
    
    with col_t1:
        st.markdown("#### 🟢 Receitas")
        if not df_rec.empty:
            cols_rec = ["Item", mes_atual_nome, mes_prox_nome] if all(x in df_rec.columns for x in ["Item", mes_atual_nome, mes_prox_nome]) else df_rec.columns
            df_rec_foco = df_rec[cols_rec].copy()
            df_rec_foco = df_rec_foco[(df_rec_foco[mes_atual_nome] > 0) | (df_rec_foco[mes_prox_nome] > 0)]
            st.dataframe(
                df_rec_foco.style.format({mes_atual_nome: "R$ {:,.2f}", mes_prox_nome: "R$ {:,.2f}"}),
                use_container_width=True
            )
            
    with col_t2:
        st.markdown("#### 🔴 Despesas")
        if not df_desp.empty:
            cols_desp = ["Item", "Categoria", mes_atual_nome, mes_prox_nome] if all(x in df_desp.columns for x in ["Item", "Categoria", mes_atual_nome, mes_prox_nome]) else df_desp.columns
            df_desp_foco = df_desp[cols_desp].copy()
            df_desp_foco = df_desp_foco[(df_desp_foco[mes_atual_nome] > 0) | (df_desp_foco[mes_prox_nome] > 0)]
            st.dataframe(
                df_desp_foco.style.format({mes_atual_nome: "R$ {:,.2f}", mes_prox_nome: "R$ {:,.2f}"}),
                use_container_width=True
            )

# ----------------- ABA 2: CATEGORIAS E METAS (SELEÇÃO DE MÊS) -----------------
with tab_grupo:
    col_cat_header, col_cat_sel = st.columns([2, 1])
    with col_cat_sel:
        mes_cat_sel = st.selectbox("📅 Selecione o Mês para Análise:", months, index=mes_atual_idx, key="sel_mes_categorias")
    with col_cat_header:
        st.subheader(f"🏷️ Agrupamento de Gastos — {mes_cat_sel}")
        
    rec_cat_val = rec_totais[mes_cat_sel]
    desp_cat_val = desp_totais[mes_cat_sel]
    
    if not df_desp.empty and "Categoria" in df_desp.columns:
        df_cat = df_desp.groupby("Categoria")[mes_cat_sel].sum().reset_index()
        df_cat = df_cat[df_cat[mes_cat_sel] > 0]
        
        c_pie, c_bar = st.columns(2)
        with c_pie:
            fig_pie = px.pie(df_cat, values=mes_cat_sel, names="Categoria", hole=0.45, title=f"Distribuição de Gastos ({mes_cat_sel})")
            fig_pie.update_layout(paper_bgcolor="rgba(0,0,0,0)", font_color="#f8fafc")
            st.plotly_chart(fig_pie, use_container_width=True)
        with c_bar:
            fig_bar = px.bar(df_cat, x="Categoria", y=mes_cat_sel, color="Categoria", text_auto='.2f', title=f"Total por Categoria (R$)")
            fig_bar.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font_color="#f8fafc", showlegend=False)
            st.plotly_chart(fig_bar, use_container_width=True)
            
        st.divider()
        st.markdown(f"#### 🎯 Metas e Teto Recomendado para {mes_cat_sel} (Regra 50-30-20)")
        col_r1, col_r2, col_r3 = st.columns(3)
        col_r1.metric("Essenciais (Moradia, Contas, Alimentação)", f"R$ {rec_cat_val * 0.50:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."), "Meta 50%")
        col_r2.metric("Estilo de Vida & Lazer", f"R$ {rec_cat_val * 0.30:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."), "Meta 30%")
        col_r3.metric("Reserva & Investimento", f"R$ {rec_cat_val * 0.20:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."), "Meta 20%")

# ----------------- ABA 3: NOVO LANÇAMENTO -----------------
with tab_add:
    st.subheader(f"➕ Adicionar Lançamento Rápido em {ano_sel}")
    with st.form("form_novo_lancamento"):
        tipo_l = st.radio("Tipo:", ["🔴 Despesa", "🟢 Receita"], horizontal=True)
        target_df = df_desp if "Despesa" in tipo_l else df_rec
        itens = list(target_df["Item"].unique()) if ("Item" in target_df.columns and not target_df.empty) else []
        
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            item_sel = st.selectbox("Item / Categoria:", ["-- Criar Novo Item --"] + itens)
            novo_item = st.text_input("Nome do Novo Item:") if item_sel == "-- Criar Novo Item --" else item_sel
        with col_f2:
            mes_l = st.selectbox("Mês:", months, index=mes_atual_idx)
            valor_l = st.number_input("Valor (R$):", min_value=0.0, step=10.0, format="%.2f")
            
        btn_salvar = st.form_submit_button("💾 Salvar Lançamento Permanente", type="primary")
        
        if btn_salvar:
            if novo_item and valor_l > 0:
                tbl_name = "despesas" if "Despesa" in tipo_l else "receitas"
                col_mes = mes_l.lower()
                
                if supabase:
                    try:
                        res_check = supabase.table(tbl_name).select("*").eq("ano", ano_sel).eq("item", novo_item).execute()
                        if res_check.data:
                            row_id = res_check.data[0]['id']
                            val_atual = res_check.data[0].get(col_mes, 0) or 0
                            supabase.table(tbl_name).update({col_mes: val_atual + valor_l}).eq("id", row_id).execute()
                        else:
                            row_data = {"ano": ano_sel, "item": novo_item, col_mes: valor_l}
                            if "Despesa" in tipo_l: row_data["categoria"] = get_category(novo_item)
                            supabase.table(tbl_name).insert(row_data).execute()
                        
                        st.cache_data.clear()
                        st.success(f"Lançamento de R$ {valor_l:.2f} gravado com SUCESSO no Banco de Dados!")
                        st.rerun()
                    except Exception as e_db:
                        st.error(f"Erro ao salvar no banco: {e_db}")

# ----------------- ABA 4: EDITAR TABELA -----------------
with tab_editar:
    st.subheader(f"✏️ Edição Direta da Planilha ({ano_sel})")
    st.info("💡 Qualquer alteração feita aqui será sincronizada permanentemente com o banco de dados Supabase!")
    
    st.markdown("### 🟢 Editar Receitas")
    df_rec_edit = st.data_editor(df_rec, num_rows="dynamic", use_container_width=True, key=f"rec_{ano_sel}")
    
    st.markdown("### 🔴 Editar Despesas")
    df_desp_edit = st.data_editor(df_desp, num_rows="dynamic", use_container_width=True, key=f"desp_{ano_sel}")
    
    if st.button("💾 Sincronizar Alterações com o Banco", type="primary"):
        if supabase:
            try:
                for _, row in df_rec_edit.iterrows():
                    item_val = row.get("Item", "")
                    if not item_val: continue
                    u_data = {"ano": ano_sel, "item": item_val}
                    for m in months: u_data[m.lower()] = float(row.get(m, 0.0) or 0.0)
                    
                    res_c = supabase.table("receitas").select("id").eq("ano", ano_sel).eq("item", item_val).execute()
                    if res_c.data:
                        supabase.table("receitas").update(u_data).eq("id", res_c.data[0]['id']).execute()
                    else:
                        supabase.table("receitas").insert(u_data).execute()

                for _, row in df_desp_edit.iterrows():
                    item_val = row.get("Item", "")
                    if not item_val: continue
                    u_data = {"ano": ano_sel, "item": item_val, "categoria": row.get("Categoria", get_category(item_val))}
                    for m in months: u_data[m.lower()] = float(row.get(m, 0.0) or 0.0)
                    
                    res_c = supabase.table("despesas").select("id").eq("ano", ano_sel).eq("item", item_val).execute()
                    if res_c.data:
                        supabase.table("despesas").update(u_data).eq("id", res_c.data[0]['id']).execute()
                    else:
                        supabase.table("despesas").insert(u_data).execute()

                st.cache_data.clear()
                st.success("Tabelas totalmente atualizadas e salvas no banco!")
                st.rerun()
            except Exception as e_sync:
                st.error(f"Erro ao sincronizar: {e_sync}")

# ----------------- ABA 5 -----------------
with tab_geral:
    st.subheader("Evolução Mensal")
    df_chart = pd.DataFrame({"Mês": months, "Receita": rec_totais.values, "Despesa": desp_totais.values})
    fig = px.bar(df_chart, x="Mês", y=["Receita", "Despesa"], barmode="group", color_discrete_map={"Receita": "#10b981", "Despesa": "#ef4444"})
    fig.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font_color="#f8fafc")
    st.plotly_chart(fig, use_container_width=True)

# ----------------- ABA 6 -----------------
with tab_completa:
    st.subheader("Receitas (12M)")
    st.dataframe(df_rec, use_container_width=True)
    st.subheader("Despesas (12M)")
    st.dataframe(df_desp, use_container_width=True)
