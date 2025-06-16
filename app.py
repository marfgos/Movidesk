import streamlit as st
import requests
import pandas as pd
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext
from datetime import datetime, timedelta
import os

# --- Configurações SharePoint ---
sharepoint_folder = '/sites/DellaVolpe/Documentos%20Compartilhados/Planejamentos/Dados_PVD/'
url_sharepoint = 'https://dellavolpecombr.sharepoint.com/sites/DellaVolpe'
username = 'SEU_EMAIL@empresa.com.br'
senha = 'SUA_SENHA_APP'  # App password ou método autenticado

# --- Função para upload no SharePoint ---
def uploadSharePoint(nome_arquivo, pasta):
    ctx_auth = AuthenticationContext(url_sharepoint)
    if ctx_auth.acquire_token_for_user(username, senha):
        ctx = ClientContext(url_sharepoint, ctx_auth)
        with open(nome_arquivo, 'rb') as f:
            target_folder = ctx.web.get_folder_by_server_relative_url(pasta)
            target_file = target_folder.upload_file(nome_arquivo, f.read())
            ctx.execute_query()
    else:
        st.error("❌ Falha na autenticação com o SharePoint.")

# --- Interface Streamlit ---
st.title("📊 Coleta de Dados Movidesk")

# Seleção interativa de data inicial
data_inicial = st.date_input("📅 Data inicial da consulta:", value=datetime(2025, 4, 1))
data_final = datetime.now().date()

# Conversão para ISO format
start_date = datetime.combine(data_inicial, datetime.min.time()).strftime("%Y-%m-%dT00:00:00Z")
end_date = datetime.combine(data_final, datetime.max.time()).strftime("%Y-%m-%dT23:59:59Z")

api_token = "34779acb-809d-4628-8594-441fa68dc694"
top = 1000

# --- ETAPA 1: Consulta principal (tickets com campos específicos) ---
if st.button("📥 Rodar Coleta Principal (PART1)"):
    with st.spinner("🔄 Coletando dados principais..."):

        def montar_url_part1(skip):
            return (
                f"https://api.movidesk.com/public/v1/tickets?"
                f"token={api_token}&$top={top}&$skip={skip}"
                f"&$filter=createdDate ge {start_date} and createdDate le {end_date}"
                f"&$select=id,subject,createdDate,status,businessEmail,owner,urgency,category,serviceFirstLevel"
            )

        def get_all_part1():
            skip = 0
            all_data = []
            while True:
                url = montar_url_part1(skip)
                r = requests.get(url)
                if r.status_code != 200:
                    break
                page = r.json()
                if not page:
                    break
                all_data.extend(page)
                if len(page) < top:
                    break
                skip += top
            return all_data

        dados = get_all_part1()
        df_part1 = pd.json_normalize(dados)
        nome_arquivo = "Tickets_Principal.csv"
        df_part1.to_csv(nome_arquivo, index=False)
        st.success("✅ Arquivo principal salvo com sucesso.")
        uploadSharePoint(nome_arquivo, sharepoint_folder)
        st.dataframe(df_part1.head())

# --- ETAPA 2: Ações dos tickets ---
if st.button("📄 Rodar Extração de Ações (PART2)"):
    with st.spinner("🔄 Coletando ações dos tickets..."):

        def montar_url_part2(skip):
            return (
                f"https://api.movidesk.com/public/v1/tickets?"
                f"token={api_token}&$select=id,actions&$expand=actions"
                f"&$filter=createdDate ge {start_date} and createdDate le {end_date}"
                f"&$top={top}&$skip={skip}"
            )

        def get_all_part2():
            skip = 0
            all_data = []
            while True:
                url = montar_url_part2(skip)
                r = requests.get(url)
                if r.status_code != 200:
                    break
                page = r.json()
                if not page:
                    break
                all_data.extend(page)
                if len(page) < top:
                    break
                skip += top
            return all_data

        raw_tickets = get_all_part2()
        registros = []

        for ticket in raw_tickets:
            ticket_id = ticket.get("id")
            for action in ticket.get("actions", []):
                registro = {"TicketId": str(ticket_id)}
                for campo, valor in action.items():
                    registro["Action_" + campo] = valor
                registros.append(registro)

        df_acoes = pd.DataFrame(registros)

        for col in df_acoes.columns:
            if "Date" in col:
                df_acoes[col] = pd.to_datetime(df_acoes[col], errors="coerce")

        nome_arquivo = "Tickets_Actions.csv"
        df_acoes.to_csv(nome_arquivo, index=False)
        st.success("✅ Arquivo de ações gerado com sucesso.")
        uploadSharePoint(nome_arquivo, sharepoint_folder)
        st.dataframe(df_acoes.head())
