import streamlit as st
import requests
import pandas as pd
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext
from datetime import datetime, timedelta
import os
import io 

# --- Configurações e Variáveis de Ambiente ---
# Use st.secrets para acessar variáveis de ambiente de forma segura no Streamlit
# Para configurar isso no Streamlit Cloud:
# Vá em "Settings" -> "Secrets" no seu aplicativo e adicione as chaves/valores
# Exemplo:
# MOVIDESK_TOKEN="34779acb-809d-4628-8594-441fa68dc694"
# SHAREPOINT_URL="https://dellavolpecombr.sharepoint.com/sites/DellaVolpe"
# SHAREPOINT_USERNAME="marcos.silva@dellavolpe.com.br"
# SHAREPOINT_PASSWORD="38213824rR!!"
# SHAREPOINT_FOLDER="/sites/DellaVolpe/Documentos%20Compartilhados/Planejamentos/Dados_PVD/"

MOVIDESK_TOKEN = st.secrets["MOVIDESK_TOKEN"]
SHAREPOINT_URL = st.secrets["SHAREPOINT_URL"]
SHAREPOINT_USERNAME = st.secrets["SHAREPOINT_USERNAME"]
SHAREPOINT_PASSWORD = st.secrets["SHAREPOINT_PASSWORD"]
SHAREPOINT_FOLDER = st.secrets["SHAREPOINT_FOLDER"]

# --- Funções do seu código original ---

def uploadSharePoint(file_content, file_name, sharepoint_folder, url_sharepoint, username, password):
    try:
        ctx_auth = AuthenticationContext(url_sharepoint)
        if ctx_auth.acquire_token_for_user(username, password):
            ctx = ClientContext(url_sharepoint, ctx_auth)
            target_folder = ctx.web.get_folder_by_server_relative_url(sharepoint_folder)
            target_folder.upload_file(file_name, file_content).execute_query()
            st.success(f"Arquivo **{file_name}** enviado com sucesso para o SharePoint!")
            return True
        else:
            st.error("Autenticação no SharePoint falhou. Verifique as credenciais.")
            return False
    except Exception as e:
        st.error(f"Erro ao fazer upload para o SharePoint: {e}")
        return False

def get_tickets_for_date(date, movidesk_token):
    start_of_day = date.strftime("%Y-%m-%d") + "T00:00:00.00z"
    end_of_day = date.strftime("%Y-%m-%d") + "T23:59:59.99z"

    api_url = (
        "https://api.movidesk.com/public/v1/tickets?"
        f"token={movidesk_token}"
        "&$select=id,type,origin,status,urgency,type,originEmailAccount,"
        "serviceFirstLevelId,serviceFull,createdBy,owner,ownerTeam,createdDate,"
        "lastUpdate,cc,originEmailAccount,clients,actions,parentTickets,"
        "childrenTickets,statusHistories,customFieldValues,assets,chatWaitingTime,"
        "resolvedIn,subject"
        "&$expand=owner,createdBy,customFieldValues($expand=items)"
        f"&$filter=createdDate ge {start_of_day} and createdDate le {end_of_day} and ownerTeam ne 'Agente - CRC'"
    )

    try:
        response = requests.get(api_url, timeout=30) # Adicione um timeout
        response.raise_for_status() # Lança um erro para status de erro (4xx ou 5xx)
        return response.json()
    except requests.exceptions.RequestException as e:
        st.error(f"Erro ao conectar à API do Movidesk para a data {date.strftime('%Y-%m-%d')}: {e}")
        return None

def extract_custom_fields(custom_field_values):
    custom_fields = {}
    if custom_field_values:
        for field in custom_field_values:
            field_id = field.get('customFieldId')
            value = field.get('value')
            if value is None and field.get('items'):
                value = field['items'][0].get('customFieldItem') if field['items'][0].get('customFieldItem') else None
            custom_fields[f'customField_{field_id}'] = value
    return custom_fields

def expand_owner(owner):
    if owner is None:
        return {
            'owner_id': None, 'owner_personType': None, 'owner_profileType': None,
            'owner_businessName': None, 'owner_email': None, 'owner_phone': None,
            'owner_pathPicture': None
        }
    return {
        'owner_id': owner.get('id'), 'owner_personType': owner.get('personType'),
        'owner_profileType': owner.get('profileType'), 'owner_businessName': owner.get('businessName'),
        'owner_email': owner.get('email'), 'owner_phone': owner.get('phone'),
        'owner_pathPicture': owner.get('pathPicture')
    }

def expand_createdby(createdby):
    if createdby is None:
        return {
            'createdBy_id': None, 'createdBy_businessName': None, 'createdBy_email': None,
            'createdBy_phone': None, 'createdBy_profileType': None, 'createdBy_personType': None
        }
    return {
        'createdBy_id': createdby.get('id'), 'createdBy_businessName': createdby.get('businessName'),
        'createdBy_email': createdby.get('email'), 'createdBy_phone': createdby.get('phone'),
        'createdBy_profileType': createdby.get('profileType'), 'createdBy_personType': createdby.get('personType')
    }

def get_first_action_description(actions):
    if actions and isinstance(actions, list) and len(actions) > 0:
        return actions[0].get('description', None)
    return None

# Mapeamento dos nomes das colunas customizadas (mantido como no seu código)
de_para_customField = {
    'customField_177683': 'CON - ANO', 'customField_178151': 'CON - ID DE CUSTO',
    'customField_177671': 'CON - MOTIVO DO TICKET', 'customField_177678': 'CON - PLACA',
    'customField_178156': 'CON - PRIMEIRO CONTATO?', 'customField_177801': 'CON - REGIÃO',
    'customField_177682': 'CON - TIPO DE VEICULO', 'customField_177804': 'CON - UF',
    'customField_186813': 'CON - VALOR DO FRENTE NEGOCIADO?', 'customField_178097': 'CON - VALOR MESA',
    'customField_177685': 'CON - VALOR NEGOCIADO', 'customField_174358': 'CRC - Clientes',
    'customField_179575': 'CRC - E-mail', 'customField_185839': 'CRC - Filial',
    'customField_175192': 'CRC - Motivo', 'customField_175190': 'CRC - Origem',
    'customField_175170': 'CRC - Tipo', 'customField_175125': 'DEV - Agente Auxiliar',
    'customField_175121': 'DEV - Classificação', 'customField_175118': 'DEV - Lista de Agentes',
    'customField_177084': 'DEV - Observação', 'customField_175123': 'DEV - Tipo de Chamado',
    'customField_177132': 'DEV -- Tipo de Chamado', 'customField_189130': 'FAT - Armazém Filial',
    'customField_189129': 'FAT - Carregamento', 'customField_189128': 'FAT - CNPJ cliente sacado',
    'customField_189139': 'FAT - Contrato cliente sacado', 'customField_189133': 'FAT - Grupo cliente sacado',
    'customField_190227': 'FAT - Número do CT-e/RPS', 'customField_189131': 'FAT - Rota emissão',
    'customField_189189': 'FAT - Status de finalização - Descrição', 'customField_189188': 'FAT - Status de finalização - Motivo',
    'customField_189184': 'FAT - Status de finalização - Setor responsável pela divergência', 'customField_190231': 'FAT - Usuário',
    'customField_188182': 'GPI - Data de Inicio', 'customField_188183': 'GPI - Data Final',
    'customField_188179': 'GPI - Observação', 'customField_188180': 'GPI - Responsável',
    'customField_188181': 'GPI - Status', 'customField_174495': 'RDA - Agente Auxiliar',
    'customField_174494': 'RDA - Assunto', 'customField_174501': 'RDA - Prazo',
    'customField_174493': 'RDA - Responsável', 'customField_174489': 'RDA - Tipo de Chamado',
    'customField_187641': 'RDA - tipo de evento', 'customField_178941': 'RDA - Tipo de evento',
    'customField_174488': 'RDA - Tipo de requisição', 'customField_174487': 'RDA - Tipo de Serviço',
    'customField_189009': 'SAC - Categoria', 'customField_189005': 'SAC - Observação',
    'customField_189008': 'SAC - Produto', 'customField_189007': 'SAC - Tipo de Atendimento',
    'customField_189010': 'SAC - Tipo de Problema', 'customField_174486': 'SAC - Tipo de Solução',
    'customField_174485': 'SAC - Tipo de Subproblema', 'customField_174484': 'SAC - Tipo de Ticket',
    'customField_188495': 'SAC - Tipo de Ticket', 'customField_177674': 'SIG - Agente Auxiliar',
    'customField_177675': 'SIG - Motivo', 'customField_177676': 'SIG - Observação',
    'customField_177677': 'SIG - Responsável', 'customField_177673': 'SIG - Tipo de Ticket'
}

# --- Lógica do Streamlit ---

st.set_page_config(
    page_title="Movidesk Data Extractor & Uploader",
    layout="wide"
)

st.title("📊 Extrator e Uploader de Dados do Movidesk")
st.markdown("Selecione um intervalo de datas para coletar tickets do Movidesk e fazer upload para o SharePoint.")

# Seleção de datas
col1, col2 = st.columns(2)
with col1:
    start_date = st.date_input("Data de Início", datetime.now().date() - timedelta(days=7))
with col2:
    end_date = st.date_input("Data Final", datetime.now().date())

# Garantir que a data de início não seja posterior à data final
if start_date > end_date:
    st.warning("A data de início não pode ser posterior à data final. Ajustando data de início...")
    start_date = end_date

# Botão para iniciar a coleta e processamento
if st.button("Coletar Dados e Gerar CSV"):
    st.info(f"Coletando dados do Movidesk de {start_date.strftime('%d/%m/%Y')} a {end_date.strftime('%d/%m/%Y')}...")
    progress_bar = st.progress(0)
    status_text = st.empty()

    dates = [start_date + timedelta(days=i) for i in range((end_date - start_date).days + 1)]
    all_data = []

    for idx, date in enumerate(dates):
        status_text.text(f"Progresso: Coletando dados para {date.strftime('%d/%m/%Y')} ({idx + 1}/{len(dates)} datas)")
        data = get_tickets_for_date(date, MOVIDESK_TOKEN)
        if data is not None: # Verifica se a chamada à API retornou algo
            if isinstance(data, list):
                all_data.extend(data)
            else:
                st.warning(f"Resposta inesperada da API para {date.strftime('%Y-%m-%d')}: {data}")
        progress_bar.progress((idx + 1) / len(dates))

    if not all_data:
        st.warning("Nenhum dado encontrado para o intervalo de datas selecionado.")
    else:
        st.success("Dados coletados com sucesso! Processando DataFrame...")
        
        # Corrigindo campos ausentes
        for item in all_data:
            if 'actions' not in item:
                item['actions'] = None
            if 'customFieldValues' not in item:
                item['customFieldValues'] = []
            if 'owner' not in item:
                item['owner'] = None
            if 'createdBy' not in item:
                item['createdBy'] = None

        df = pd.DataFrame(all_data)

        # Extraindo e expandindo campos (com tratamento de erro)
        df['first_action_description'] = df['actions'].apply(get_first_action_description)

        # Usando .apply(lambda x: x if isinstance(x, list) else []) para customFieldValues
        # para garantir que é uma lista antes de passar para extract_custom_fields
        expanded_fields = df['customFieldValues'].apply(lambda x: extract_custom_fields(x) if isinstance(x, list) else {})
        custom_fields_df = pd.DataFrame(expanded_fields.tolist())

        expanded_owners = df['owner'].apply(expand_owner)
        owner_fields_df = pd.DataFrame(expanded_owners.tolist())

        expanded_createdBy = df['createdBy'].apply(expand_createdby)
        createdBy_fields_df = pd.DataFrame(expanded_createdBy.tolist())
        
        # Juntando DataFrames, lidando com colunas duplicadas
        # Crie uma lista de dataframes para concatenação
        dfs_to_concat = [
            df.drop(columns=['owner', 'customFieldValues', 'createdBy', 'actions'], errors='ignore'),
            owner_fields_df,
            custom_fields_df,
            createdBy_fields_df
        ]

        # Use axis=1 para concatenação lateral
        df_final = pd.concat(dfs_to_concat, axis=1)

        # Formatando datas (com tratamento de erro para colunas que podem não existir)
        if 'createdDate' in df_final.columns:
            df_final['createdDate'] = pd.to_datetime(df_final['createdDate'], errors='coerce').dt.strftime('%d/%m/%Y %H:%M')
        if 'resolvedIn' in df_final.columns:
            df_final['resolvedIn'] = pd.to_datetime(df_final['resolvedIn'], errors='coerce').dt.strftime('%d/%m/%Y %H:%M')
        
        # Renomeando as colunas dos campos customizados
        df_final = df_final.rename(columns=de_para_customField)

        st.subheader("Prévia dos Dados Coletados")
        st.dataframe(df_final.head())

        # Botão de download
        csv_file_name = f"TicketsMovidesk_{start_date.strftime('%Y%m%d')}_to_{end_date.strftime('%Y%m%d')}.csv"
        csv_buffer = io.StringIO()
        df_final.to_csv(csv_buffer, index=False, encoding='utf-8-sig') # Usar utf-8-sig para compatibilidade com Excel
        st.download_button(
            label="💾 Baixar CSV",
            data=csv_buffer.getvalue(),
            file_name=csv_file_name,
            mime="text/csv",
        )

        # Botão para upload para o SharePoint
        if st.button("📤 Fazer Upload para o SharePoint"):
            with st.spinner("Enviando arquivo para o SharePoint..."):
                # Passar o conteúdo do CSV em memória e o nome do arquivo
                upload_success = uploadSharePoint(
                    io.BytesIO(csv_buffer.getvalue().encode('utf-8-sig')), # Converter para bytes
                    csv_file_name,
                    SHAREPOINT_FOLDER,
                    SHAREPOINT_URL,
                    SHAREPOINT_USERNAME,
                    SHAREPOINT_PASSWORD
                )
                if upload_success:
                    st.balloons() # Uma pequena celebração visual
