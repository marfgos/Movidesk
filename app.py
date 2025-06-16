import streamlit as st
import pandas as pd
import requests
from datetime import datetime, timedelta
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext
import os

# --- Configurações SharePoint ---
sharepoint_folder = '/sites/DellaVolpe/Documentos%20Compartilhados/Planejamentos/Dados_PVD/'
url_sharepoint = 'https://dellavolpecombr.sharepoint.com/sites/DellaVolpe'
username = 'marcos.silva@dellavolpe.com.br'
password = '38213824rR!!'

# --- Função para upload no SharePoint ---
def uploadSharePoint(filename, pasta_destino):
    ctx_auth = AuthenticationContext(url_sharepoint)
    if ctx_auth.acquire_token_for_user(username, senha):
        ctx = ClientContext(url_sharepoint, ctx_auth)
        with open(filename, 'rb') as content_file:
            file_content = content_file.read()
        target_folder = ctx.web.get_folder_by_server_relative_url(pasta_destino)
        target_file = target_folder.upload_file(os.path.basename(filename), file_content).execute_query()
        st.success(f"📁 Arquivo '{filename}' enviado para o SharePoint com sucesso.")
    else:
        st.error("❌ Falha na autenticação do SharePoint.")

# --- INTERFACE STREAMLIT ---
st.set_page_config(layout="wide")
st.title("🔄 Extração de Dados Movidesk + Upload para SharePoint")

# --- INTERVALO DE DATAS PADRÃO ---
data_inicial = datetime(2025, 4, 1)

# === ETAPA 1: Extração principal de tickets ===
if st.button("📦 Rodar Extração de Tickets (PART1)"):
    with st.spinner("🔄 Coletando tickets..."):

        # Configurações da API
        api_token = "34779acb-809d-4628-8594-441fa68dc694"
        top = 1000
        base_url = "https://api.movidesk.com/public/v1/tickets"

        start_date = data_inicial.strftime("%Y-%m-%dT00:00:00Z")
        end_date = datetime.now().strftime("%Y-%m-%dT23:59:59Z")

        def montar_url(skip):
            return (
                f"{base_url}?token={api_token}&$filter=createdDate ge {start_date} and createdDate le {end_date}"
                f"&$top={top}&$skip={skip}"
            )

        def get_page(skip):
            url = montar_url(skip)
            r = requests.get(url)
            return r.json() if r.status_code == 200 else []

        def get_all():
            skip = 0
            all_data = []
            while True:
                page = get_page(skip)
                if not page:
                    break
                all_data.extend(page)
                if len(page) < top:
                    break
                skip += top
            return all_data

        tickets = get_all()
        df_tickets = pd.json_normalize(tickets)
        csv_file_1 = "Tickets_Movidesk.csv"
        df_tickets.to_csv(csv_file_1, index=False)
        st.success(f"✅ Arquivo '{csv_file_1}' salvo.")
        uploadSharePoint(csv_file_1, sharepoint_folder)
        st.dataframe(df_tickets.head())

    st.balloons()

# === ETAPA 2: Extração de Ações ===
if st.button("📝 Rodar Extração de Ações dos Tickets"):
    with st.spinner("🔄 Coletando ações dos tickets..."):

        def montar_url(skip):
            return (
                f"{base_url}?token={api_token}&$select=id,actions"
                f"&$expand=actions"
                f"&$filter=createdDate ge {start_date} and createdDate le {end_date}"
                f"&$top={top}&$skip={skip}"
            )

        def get_page(skip):
            url = montar_url(skip)
            r = requests.get(url)
            return r.json() if r.status_code == 200 else []

        def get_all_actions():
            skip = 0
            all_data = []
            while True:
                page = get_page(skip)
                if not page:
                    break
                all_data.extend(page)
                if len(page) < top:
                    break
                skip += top
            return all_data

        raw_data = get_all_actions()
        flat_rows = []
        for item in raw_data:
            ticket_id = item.get('id')
            actions = item.get('actions', [])
            for action in actions:
                action_record = {'TicketId': ticket_id}
                for key, value in action.items():
                    action_record[f'Action_{key}'] = value
                flat_rows.append(action_record)

        df_actions = pd.DataFrame(flat_rows)

        # Conversão de datas
        for col in df_actions.columns:
            if "Date" in col:
                df_actions[col] = pd.to_datetime(df_actions[col], errors='coerce')

        csv_file_2 = "Tickets_Actions.csv"
        df_actions.to_csv(csv_file_2, index=False)
        st.success(f"✅ Arquivo '{csv_file_2}' salvo.")
        uploadSharePoint(csv_file_2, sharepoint_folder)
        st.dataframe(df_actions.head())

    st.balloons()
