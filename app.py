import streamlit as st
import pandas as pd
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext
from datetime import datetime
import os

# --- Configurações SharePoint ---
sharepoint_folder = '/sites/DellaVolpe/Documentos%20Compartilhados/Planejamentos/Dados_PVD/'
url_sharepoint = 'https://dellavolpecombr.sharepoint.com/sites/DellaVolpe'
username = 'marcos.silva@dellavolpe.com.br'
password = '38213824rR!!'

def uploadSharePoint(nome_arquivo, pasta):
    ctx_auth = AuthenticationContext(url_sharepoint)
    if ctx_auth.acquire_token_for_user(username, password):
        ctx = ClientContext(url_sharepoint, ctx_auth)
        with open(nome_arquivo, 'rb') as f:
            target_folder = ctx.web.get_folder_by_server_relative_url(pasta)
            target_folder.upload_file(nome_arquivo, f.read()).execute_query()
    else:
        st.error("Falha na autenticação com o SharePoint")

# --- Parâmetros da interface ---
data_inicio = st.date_input("Data de início", datetime(2025, 4, 1))
data_fim = datetime.now().date()

# Botão 1 - Coleta Principal
if st.button("Rodar Coleta Principal"):
    df_principal = pd.DataFrame({  # Simule aqui a coleta de dados como no PART1
        "TicketId": [1, 2],
        "Status": ["Aberto", "Fechado"]
    })

    nome_arquivo = f"tickets_principal_{data_inicio}_{data_fim}.csv"
    df_principal.to_csv(nome_arquivo, index=False)
    uploadSharePoint(nome_arquivo, sharepoint_folder)
    st.success("Arquivo da coleta principal enviado com sucesso.")

# Botão 2 - Coleta Ações
if st.button("Rodar Coleta Ações"):
    df_acoes = pd.DataFrame({  # Simule aqui o que sua consulta M fazia
        "TicketId": [1, 1, 2],
        "ActionType": [1, 2, 1],
        "ActionDate": [data_inicio, data_fim, data_fim]
    })

    nome_arquivo = f"tickets_acoes_{data_inicio}_{data_fim}.csv"
    df_acoes.to_csv(nome_arquivo, index=False)
    uploadSharePoint(nome_arquivo, sharepoint_folder)
    st.success("Arquivo das ações enviado com sucesso.")
