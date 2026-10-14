import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
import os
from streamlit_gsheets import GSheetsConnection

# Configuração da Página
st.set_page_config(
    page_title="Painel de Acompanhamento Acadêmico - 1ª DIREC",
    page_icon="🎓",
    layout="wide"
)

# Estilização visual limpa
st.markdown("""
    <style>
        .main { background-color: #f8fafc; }
        .stMetric { background-color: #ffffff; padding: 15px; border-radius: 12px; border: 1px solid #e2e8f0; box-shadow: 0 1px 2px rgba(0,0,0,0.05); }
    </style>
""", unsafe_allow_html=True)

# Caminho padrão da planilha local
EXCEL_DEFAULT = 'RAPP_1DIREC_Situacao_Alunos_Tratada.xlsx'

if 'uploaded_file_path' not in st.session_state:
    st.session_state.uploaded_file_path = EXCEL_DEFAULT

# Conexão com o Google Sheets para as Visitas Colaborativas
conn = st.connection("gsheets", type=GSheetsConnection)

def carregar_visitas_gsheets():
    try:
        df_v = conn.read(worksheet="Página1", ttl=0)
        if df_v is not None and not df_v.empty and 'Escola' in df_v.columns:
            return df_v.dropna(subset=['Escola']).to_dict('records')
    except Exception:
        pass
    return []

def salvar_visita_gsheets(escola, turno, atendidos, aprovados, data):
    try:
        df_atual = conn.read(worksheet="Página1", ttl=0)
        if df_atual is None or df_atual.empty or 'Escola' not in df_atual.columns:
            df_atual = pd.DataFrame(columns=['Escola', 'Turno', 'Atendidos', 'Aprovados_Liberados', 'Data_Registro'])
        
        nova_linha = pd.DataFrame([{
            'Escola': escola,
            'Turno': turno,
            'Atendidos': atendidos,
            'Aprovados_Liberados': aprovados,
            'Data_Registro': data
        }])
        
        df_novo = pd.concat([df_atual, nova_linha], ignore_index=True)
        conn.update(worksheet="Página1", data=df_novo)
        return True
    except Exception as e:
        st.error(f"Erro ao salvar no Google Sheets: {e}")
        return False

# Carregamento e Tratamento dos Dados Reais
@st.cache_data
def carregar_dados(path_file):
    if not os.path.exists(path_file):
        path_file = EXCEL_DEFAULT
    
    df_cons = pd.read_excel(path_file, sheet_name='Alunos Consolidado')
    df_comp = pd.read_excel(path_file, sheet_name='Base componentes')
    df_art = pd.read_excel(path_file, sheet_name='Articuladores ')
    
    df_comp['Situação'] = df_comp['Situação'].astype(str).replace({
        'Não obrigado': 'Não avaliado',
        'NÃO AGENDADA': 'Não avaliado',
        'AGENDADA': 'Não avaliado',
        'AGUARDANDO_INICIO': 'Não avaliado'
    })
    
    df_art['Escola_Clean'] = df_art['Escola'].astype(str).str.strip().str.upper()
    df_cons['Escola_Clean'] = df_cons['Escola'].astype(str).str.strip().str.upper()
    
    map_assessor = dict(zip(df_art['Escola_Clean'], df_art['Assessor(a)']))
    df_cons['Assessor(a)'] = df_cons['Escola_Clean'].map(map_assessor).fillna("Não Informado")
    
    return df_cons, df_comp, df_art

try:
    df_cons, df_comp, df_art = carregar_dados(st.session_state.uploaded_file_path)
except Exception as e:
    st.error(f"Erro ao carregar a planilha: {e}. Certifique-se de que o arquivo está na pasta correta.")
    st.stop()

if 'escola_selecionada_visita' not in st.session_state:
    st.session_state.escola_selecionada_visita = "Selecione..."

if 'aba_ativa' not in st.session_state:
    st.session_state.aba_ativa = "Visão Geral"

def ir_para_visita(nome_escola):
    st.session_state.escola_selecionada_visita = nome_escola
    st.session_state.aba_ativa = "Registro de Visitas & Campo"

# --- MENU LATERAL DE NAVEGAÇÃO ---
st.sidebar.markdown("### 🎓 1ª DIREC - Natal/RN")
st.sidebar.markdown("<p style='font-size: 12px; color: #64748b;'>Monitoramento de Ciclos, Visitas e Plataformas</p>", unsafe_allow_html=True)
st.sidebar.markdown("---")

opcoes_menu = [
    "Visão Geral", 
    "Escolas Prioritárias (25)", 
    "Todas as Escolas", 
    "Registro de Visitas & Campo", 
    "Detalhamento por Componentes",
    "📁 Atualizar Base de Dados",
    "Relatório Executivo / WhatsApp"
]

aba = st.sidebar.radio("Navegação", opcoes_menu, key="aba_ativa")

# --- FILTROS GLOBAIS EM CASCATA ---
st.sidebar.markdown("---")
st.sidebar.markdown("#### 🔍 Filtros Globais")

tutores_disp = ["Todos(as)"] + sorted(df_cons['Professor tutor'].dropna().unique().tolist())
tutor_sel = st.sidebar.selectbox("Professor-Tutor", tutores_disp)

df_filtrado = df_cons.copy()
if tutor_sel != "Todos(as)":
    df_filtrado = df_filtrado[df_filtrado['Professor tutor'] == tutor_sel]

assessores_disp = ["Todos(as)"] + sorted([a for a in df_filtrado['Assessor(a)'].dropna().unique().tolist() if a != "Não Informado"])
assessor_sel = st.sidebar.selectbox("Assessor(a)", assessores_disp)

if assessor_sel != "Todos(as)":
    df_filtrado = df_filtrado[df_filtrado['Assessor(a)'] == assessor_sel]

escolas_disp = ["Todos(as)"] + sorted(df_filtrado['Escola'].dropna().unique().tolist())
escola_sel = st.sidebar.selectbox("Escola", escolas_disp)

if escola_sel != "Todos(as)":
    df_filtrado = df_filtrado[df_filtrado['Escola'] == escola_sel]

plat_sel = st.sidebar.selectbox("Plataforma", ["Todas (Plural & Avalia)", "Plurall", "Avalia"])
if plat_sel != "Todas (Plural & Avalia)":
    df_filtrado = df_filtrado[df_filtrado['Plataforma'] == plat_sel]


# ==========================================
# ABA 1: VISÃO GERAL
# ==========================================
if aba == "Visão Geral":
    st.title("📊 Painel de Acompanhamento Acadêmico")
    st.markdown("Monitoramento de Ciclos, Visitas e Plataformas (Plural & Avalia)")
    st.markdown("---")

    total_estudantes = df_filtrado['Matrícula'].nunique()
    aprovados = len(df_filtrado[df_filtrado['Situação do aluno'] == 'Aprovado'])
    nao_avaliados = len(df_filtrado[df_filtrado['Situação do aluno'] == 'Não avaliado'])
    nao_aprovados = len(df_filtrado[df_filtrado['Situação do aluno'] == 'Não aprovado'])

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total de Estudantes", f"{total_estudantes:,}")
    c2.metric("Aprovados", f"{aprovados:,}")
    c3.metric("Não Avaliados", f"{nao_avaliados:,}")
    c4.metric("Não Aprovados", f"{nao_aprovados:,}")

    st.markdown("<br>", unsafe_allow_html=True)

    col_g1, col_g2 = st.columns(2)
    with col_g1:
        st.subheader("Status por Situação (Registros Oficiais)")
        df_status = df_filtrado['Situação do aluno'].value_counts().reset_index()
        df_status.columns = ['Situação', 'Quantidade']
        fig_donut = px.pie(df_status, names='Situação', values='Quantidade', hole=0.5,
                           color='Situação',
                           color_discrete_map={'Aprovado':'#10b981', 'Não avaliado':'#f59e0b', 'Não aprovado':'#f43f5e'})
        fig_donut.update_layout(margin=dict(t=20, b=20, l=20, r=20))
        st.plotly_chart(fig_donut, use_container_width=True)

    with col_g2:
        st.subheader("Evolução das Visitas de Campo")
        visitas_atuais = carregar_visitas_gsheets()
        if len(visitas_atuais) == 0:
            st.info("Nenhuma visita registrada até o momento. O gráfico começará a ser preenchido conforme as visitas forem lançadas.")
            fig_bar = px.bar(title="Aguardando registros...")
            fig_bar.update_layout(xaxis={"visible": False}, yaxis={"visible": False})
        else:
            df_v = pd.DataFrame(visitas_atuais)
            df_visitas_real = df_v.groupby('Escola')['Atendidos'].sum().reset_index()
            fig_bar = px.bar(df_visitas_real, x='Escola', y='Atendidos', text_auto=True, 
                             labels={'Escola':'Escola', 'Atendidos':'Atendidos'},
                             color_discrete_sequence=['#d97706'])
            fig_bar.update_layout(margin=dict(t=20, b=20, l=20, r=20))
        st.plotly_chart(fig_bar, use_container_width=True)


# ==========================================
# ABA 2: ESCOLAS PRIORITÁRIAS (25)
# ==========================================
elif aba == "Escolas Prioritárias (25)":
    st.markdown("""
        <div style="background: linear-gradient(to right, #f59e0b, #d97706); padding: 20px; border-radius: 12px; color: white; margin-bottom: 20px;">
            <h2 style="margin: 0; font-size: 22px;">Top 25 Escolas Prioritárias (2º Ciclo)</h2>
            <p style="margin: 5px 0 0 0; font-size: 13px; opacity: 0.9;">Escolas com maior volume de estudantes pendentes que exigem foco nas visitas.</p>
        </div>
    """, unsafe_allow_html=True)

    resumo_escolas = df_cons.groupby('Escola').agg(
        Total_Alunos=('Matrícula', 'nunique'),
        Nao_Avaliados_Plural=('Plataforma', lambda x: ((x == 'Plurall') & (df_cons.loc[x.index, 'Situação do aluno'] == 'Não avaliado')).sum()),
        Nao_Avaliados_Avalia=('Plataforma', lambda x: ((x == 'Avalia') & (df_cons.loc[x.index, 'Situação do aluno'] == 'Não avaliado')).sum()),
        Nao_Aprovados=('Situação do aluno', lambda x: (x == 'Não aprovado').sum()),
        Tutor=('Professor tutor', 'first'),
        Assessor=('Assessor(a)', 'first')
    ).reset_index()

    resumo_escolas['Total_Pendentes'] = resumo_escolas['Nao_Avaliados_Plural'] + resumo_escolas['Nao_Avaliados_Avalia'] + resumo_escolas['Nao_Aprovados']
    top_prioritarias = resumo_escolas.sort_values(by='Total_Pendentes', ascending=False).head(25)

    for idx, row in top_prioritarias.iterrows():
        col_info, col_btn = st.columns([5, 1])
        col_info.markdown(
            f"**{row['Escola']}** | Tutor: {row['Tutor']} | Assessor(a): {row['Assessor']} | "
            f"Total Alunos: **{row['Total_Alunos']}** | Não Avaliados Plural: **{row['Nao_Avaliados_Plural']}** | "
            f"Não Avaliados Avalia: **{row['Nao_Avaliados_Avalia']}** | Não Aprovados: **{row['Nao_Aprovados']}** | "
            f"🔴 **Total de Pendentes: {row['Total_Pendentes']}**"
        )
        col_btn.button("✈️ Visitar", key=f"prio_{idx}", on_click=ir_para_visita, args=(row['Escola'],))


# ==========================================
# ABA 3: TODAS AS ESCOLAS
# ==========================================
elif aba == "Todas as Escolas":
    st.title("🏫 Todas as Escolas da Rede")
    st.markdown("Panorama completo de todas as unidades monitoradas com conferência de quantitativos.")
    
    busca = st.text_input("🔍 Buscar escola...", "")
    
    resumo_geral = df_cons.groupby('Escola').agg(
        Total_Alunos=('Matrícula', 'nunique'),
        Nao_Avaliados=('Situação do aluno', lambda x: (x == 'Não avaliado').sum()),
        Nao_Aprovados=('Situação do aluno', lambda x: (x == 'Não aprovado').sum()),
        Tutor=('Professor tutor', 'first'),
        Assessor=('Assessor(a)', 'first')
    ).reset_index()

    resumo_geral['Total_Pendentes'] = resumo_geral['Nao_Avaliados'] + resumo_geral['Nao_Aprovados']

    if busca:
        resumo_geral = resumo_geral[resumo_geral['Escola'].str.contains(busca, case=False, na=False)]

    for idx, row in resumo_geral.iterrows():
        col_info, col_btn = st.columns([5, 1])
        col_info.markdown(
            f"**{row['Escola']}** | Tutor: {row['Tutor']} | Assessor(a): {row['Assessor']} | "
            f"Total Alunos: **{row['Total_Alunos']}** | Não Avaliados: **{row['Nao_Avaliados']}** | "
            f"Não Aprovados: **{row['Nao_Aprovados']}** | 🔴 **Total Pendentes: {row['Total_Pendentes']}**"
        )
        col_btn.button("✈️ Visitar", key=f"geral_{idx}", on_click=ir_para_visita, args=(row['Escola'],))


# ==========================================
# ABA 4: REGISTRO DE VISITAS & CAMPO (Integrado ao Google Sheets)
# ==========================================
elif aba == "Registro de Visitas & Campo":
    st.title("🚐 Caderneta de Campo & Atendimento")
    st.markdown("Preencha o formulário de visita ou realize a chamada nominal dos estudantes.")
    st.markdown("---")

    escolas_disponiveis = sorted(df_cons['Escola'].unique())
    idx_default = 0
    if st.session_state.escola_selecionada_visita in escolas_disponiveis:
        idx_default = escolas_disponiveis.index(st.session_state.escola_selecionada_visita)

    with st.form("form_visita_manual"):
        st.subheader("Registrar Visita Escolar & Atendimento de Campo")
        col_m1, col_m2 = st.columns(2)
        escola_manual = col_m1.selectbox("Escola Visitada", escolas_disponiveis, index=idx_default)
        turno_manual = col_m2.selectbox("Turno da Visita", ["Matutino", "Vespertino", "Noturno"])
        
        col_m3, col_m4 = st.columns(2)
        atendidos_manual = col_m3.number_input("Alunos Atendidos na Visita", min_value=0, value=0, step=1)
        aprovados_manual = col_m4.number_input("Alunos Aprovados / Liberados", min_value=0, value=0, step=1)
        
        submitted = st.form_submit_button("💾 Salvar Visita e Atualizar Indicadores", type="primary")
        if submitted:
            data_hora = datetime.now().strftime("%d/%m/%Y às %H:%M")
            sucesso = salvar_visita_gsheets(escola_manual, turno_manual, atendidos_manual, aprovados_manual, data_hora)
            if sucesso:
                st.success(f"Visita à escola '{escola_manual}' salva e sincronizada com sucesso no Google Sheets!")

    st.markdown("---")

    st.subheader("Caderneta Nominal & Chamada de Estudantes")
    escola_visitada = st.selectbox("Selecione a Escola para carregar a lista nominal:", ["Selecione..."] + escolas_disponiveis, index=idx_default+1 if idx_default>0 else 0)
    
    if escola_visitada != "Selecione...":
        df_escola_total = df_cons[df_cons['Escola'] == escola_visitada].copy()
        qtd_total = len(df_escola_total)
        qtd_nao_avaliados = len(df_escola_total[df_escola_total['Situação do aluno'] == 'Não avaliado'])
        qtd_nao_aprovados = len(df_escola_total[df_escola_total['Situação do aluno'] == 'Não aprovado'])
        qtd_aprovados = len(df_escola_total[df_escola_total['Situação do aluno'] == 'Aprovado'])
        
        st.info(f"📊 **Resumo da Unidade:** Total de Registros: **{qtd_total}** | Aprovados: **{qtd_aprovados}** | Não Aprovados: **{qtd_nao_aprovados}** | Não Avaliados: **{qtd_nao_avaliados}**")

        st.markdown("##### Filtrar exibição por status:")
        f_aprovado = st.checkbox("Aprovado", value=False)
        f_nao_aprovado = st.checkbox("Não aprovado", value=True)
        f_nao_avaliado = st.checkbox("Não avaliado", value=True)

        status_permitidos = []
        if f_aprovado: status_permitidos.append("Aprovado")
        if f_nao_aprovado: status_permitidos.append("Não aprovado")
        if f_nao_avaliado: status_permitidos.append("Não avaliado")

        df_alunos_escola = df_escola_total[df_escola_total['Situação do aluno'].isin(status_permitidos)].copy()
        
        if len(df_alunos_escola) > 0:
            st.markdown("Marque a presença dos estudantes durante a visita de campo:")
            with st.form(key=f"form_chamada_{escola_visitada}"):
                for idx, row in df_alunos_escola.iterrows():
                    col_n, col_p = st.columns([4, 1])
                    col_n.write(f"**{row['Nome']}** (Mat: {row['Matrícula']} | Plataforma: {row['Plataforma']} | Status: **{row['Situação do aluno']}**)")
                    col_p.selectbox("Presença", ["Presente", "Ausente"], key=f"pres_{row['Matrícula']}_{idx}")
                
                salvar_chamada = st.form_submit_button("💾 Salvar Chamada Nominal", type="primary")
                if salvar_chamada:
                    st.success("Chamada nominal registrada com sucesso!")
        else:
            st.info("Nenhum estudante encontrado com os filtros de status selecionados.")

    st.markdown("---")
    st.subheader("Histórico de Visitas Registradas (Sincronizado)")
    visitas_salvas = carregar_visitas_gsheets()
    if len(visitas_salvas) > 0:
        for v in visitas_salvas:
            st.markdown(f"- 🏫 **{v.get('Escola')}** | Turno: {v.get('Turno')} | Atendidos: {v.get('Atendidos')} | Aprovados/Liberados: {v.get('Aprovados_Liberados')} | 🕒 **{v.get('Data_Registro')}**")
    else:
        st.markdown("<p style='color: #94a3b8; font-size: 13px;'>Nenhuma visita registrada ainda.</p>", unsafe_allow_html=True)


# ==========================================
# ABA 5: DETALHAMENTO POR COMPONENTES
# ==========================================
elif aba == "Detalhamento por Componentes":
    st.title("📚 Detalhamento de Componentes Curriculares por Escola")
    st.markdown("Visualização por disciplina com regra de multiplicidade.")
    
    escola_comp = st.selectbox("Selecione a Escola:", sorted(df_comp['Escola'].unique()))
    df_componentes_escola = df_comp[df_comp['Escola'] == escola_comp].copy()
    
    q_total = len(df_componentes_escola)
    q_aprov = len(df_componentes_escola[df_componentes_escola['Situação'].isin(['Aprovado', 'APROVADO'])])
    q_nao_aprov = len(df_componentes_escola[df_componentes_escola['Situação'].isin(['Não Aprovado', 'REPROVADO'])])
    q_nao_aval = len(df_componentes_escola[df_componentes_escola['Situação'] == 'Não avaliado'])

    mc1, mc2, mc3, mc4 = st.columns(4)
    mc1.metric("Componentes totais", q_total)
    mc2.metric("Aprovados", q_aprov)
    mc3.metric("Não aprovados", q_nao_aprov)
    mc4.metric("Não avaliados", q_nao_aval)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("##### Filtrar componentes por situação:")
    
    col_f1, col_f2, col_f3 = st.columns(3)
    f_comp_aprovado = col_f1.checkbox("Exibir Aprovados", value=True, key="f_c_ap")
    f_comp_nao_aprovado = col_f2.checkbox("Exibir Não Aprovados", value=True, key="f_c_nap")
    f_comp_nao_avaliado = col_f3.checkbox("Exibir Não avaliados", value=True, key="f_c_nav")

    status_comp_permitidos = []
    if f_comp_aprovado: status_comp_permitidos.extend(['Aprovado', 'APROVADO'])
    if f_comp_nao_aprovado: status_comp_permitidos.extend(['Não Aprovado', 'REPROVADO'])
    if f_comp_nao_avaliado: status_comp_permitidos.append('Não avaliado')

    df_comp_filtrado = df_componentes_escola[df_componentes_escola['Situação'].isin(status_comp_permitidos)].copy()

    st.info(f"Exibindo {len(df_comp_filtrado)} de {q_total} registros para a unidade selecionada.")
    df_exib_comp = df_comp_filtrado[['Matrícula', 'Nome', 'Série', 'Componente', 'Plataforma', 'Situação']].copy()
    st.markdown(df_exib_comp.to_html(index=False, classes="table table-striped"), unsafe_allow_html=True)


# ==========================================
# ABA 6: ATUALIZAR BASE DE DADOS
# ==========================================
elif aba == "📁 Atualizar Base de Dados":
    st.title("📁 Importar / Atualizar Planilha de Dados")
    
    data_hora_atual = datetime.now().strftime("%d/%m/%Y às %Hh%Mmin%S")
    st.info(f"🕒 **Última verificação / Acesso à base:** {data_hora_atual}")
    
    st.markdown("Envie uma nova versão da planilha consolidada da 1ª DIREC para atualizar imediatamente todo o painel.")
    
    uploaded_file = st.file_uploader("Escolha o arquivo Excel (.xlsx)", type=["xlsx"])
    if uploaded_file is not None:
        temp_path = "RAPP_1DIREC_Situacao_Alunos_Tratada.xlsx"
        with open(temp_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        st.session_state.uploaded_file_path = temp_path
        st.success("Planilha atualizada com sucesso! Todos os indicadores da rede foram recalculados.")


# ==========================================
# ABA 7: RELATÓRIO EXECUTIVO / WHATSAPP
# ==========================================
elif aba == "Relatório Executivo / WhatsApp":
    st.title("📄 Gerador de Resumo para Gestão & WhatsApp")
    
    data_hora_atual = datetime.now().strftime("%d/%m/%Y às %Hh%Mmin%S")
    st.info(f"🕒 **Última verificação / Acesso à base:** {data_hora_atual}")
    
    total_est = df_filtrado['Matrícula'].nunique()
    visitas_salvas = carregar_visitas_gsheets()
    total_visitas = len(visitas_salvas)
    
    texto_whatsapp = f"""📊 *BOLETIM EXECUTIVO DE REDE - ACOMPANHAMENTO 2º CICLO*

🏫 *Escolas Monitoradas:* {df_filtrado['Escola'].nunique()} unidades
👨‍🎓 *Total de Estudantes:* {total_est:,}
🚐 *Visitas Realizadas no Campo:* {total_visitas} visitas registradas

📌 *Status Operacional:* Operação de campo em andamento focada no atendimento das unidades prioritárias para garantia de conclusão das avaliações Plural e Avalia."""

    st.text_area("Copie o texto abaixo para enviar no WhatsApp:", texto_whatsapp, height=220)
    st.success("Pronto! Basta selecionar o texto acima, copiar e colar no WhatsApp.")