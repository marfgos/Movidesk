import streamlit as st
import requests
import pandas as pd
from datetime import datetime, timedelta
import os

# --- Caminho desejado (apenas usado se rodar localmente no Windows) ---
DOWNLOADS_PATH = r"C:\Users\MArcos.Silva\Downloads\TicketsMovidesk.csv"

# --- Lista fixa de e-mails permitidos ---
ALLOWED_EMAILS = [
    "karina.viana@dellavolpe.com.br",
    "danillo.silva@dellavolpe.com.br",
    "thayane.jesus@dellavolpe.com.br",
    "ana.jesus@dellavolpe.com.br",
    "thicyane.pena@dellavolpe.com.br",
    "brenda.felgueiras@dellavolpe.com.br",
    "erick.martini@dellavolpe.com.br",
    "marcos.silva@dellavolpe.com.br"
]
ALLOWED_EMAILS = [e.strip().lower() for e in ALLOWED_EMAILS]

# --- Funções auxiliares ---
def get_tickets_for_date(date):
    start_of_day = date.strftime("%Y-%m-%d") + "T00:00:00.00z"
    end_of_day = date.strftime("%Y-%m-%d") + "T23:59:59.99z"
    api_url = (
        "https://api.movidesk.com/public/v1/tickets?"
        "token=34779acb-809d-4628-8594-441fa68dc694"
        "&$select=id,type,origin,status,urgency,type,originEmailAccount,"
        "serviceFirstLevelId,serviceFull,createdBy,owner,ownerTeam,createdDate,"
        "lastUpdate,cc,originEmailAccount,clients,actions,parentTickets,"
        "childrenTickets,statusHistories,customFieldValues,assets,chatWaitingTime,"
        "resolvedIn,subject"
        "&$expand=owner,createdBy,customFieldValues($expand=items)"
        f"&$filter=createdDate ge {start_of_day} and createdDate le {end_of_day} and ownerTeam ne 'Agente - CRC'"
    )
    response = requests.get(api_url)
    return response.json()

def extract_custom_fields(custom_field_values):
    custom_fields = {}
    for field in custom_field_values:
        field_id = field.get('customFieldId')
        value = field.get('value')
        if not value and field.get('items'):
            value = field['items'][0].get('customFieldItem')
        custom_fields[f'customField_{field_id}'] = value
    return custom_fields

def expand_owner(owner):
    if owner is None:
        return {}
    return {
        'owner_id': owner.get('id'),
        'owner_businessName': owner.get('businessName'),
        'owner_email': owner.get('email')
    }

def expand_createdby(createdby):
    if createdby is None:
        return {}
    return {
        'createdBy_id': createdby.get('id'),
        'createdBy_businessName': createdby.get('businessName'),
        'createdBy_email': createdby.get('email')
    }

def get_first_action_description(actions):
    if actions and len(actions) > 0:
        return actions[0].get('description')
    return None

# --- Streamlit app ---
st.title("📊 Coleta de Tickets Movidesk")

# 🔹 SELETOR DE BASE
tipo_base = st.radio(
    "Selecione a base desejada:",
    (
        "Somente chamados criados pelos usuários autorizados",
        "Somente chamados do serviço \"Agendamento\""
    )
)

data_inicial = st.date_input(
    "Selecione a data inicial:",
    value=datetime(2025, 6, 1).date(),
    min_value=datetime(2025, 1, 1).date(),
    max_value=datetime.now().date()
)

if st.button("🚀 Extrair, filtrar e salvar/baixar CSV"):
    from zoneinfo import ZoneInfo
    execution_timestamp = datetime.now(
        ZoneInfo("America/Sao_Paulo")
    ).strftime('%d/%m/%Y %H:%M:%S')

    st.info(f"🕒 Data/hora da execução: {execution_timestamp}")

    with st.spinner("Extraindo base..."):
        start_date = datetime.combine(data_inicial, datetime.min.time())
        end_date = datetime.now()
        dates = [start_date + timedelta(days=i) for i in range((end_date - start_date).days + 1)]

        all_data = []
        progress = st.progress(0)

        for idx, date in enumerate(dates, 1):
            data = get_tickets_for_date(date)
            if isinstance(data, list):
                all_data.extend(data)
            progress.progress(idx / len(dates))

        df = pd.DataFrame(all_data)
        df['first_action_description'] = df['actions'].apply(get_first_action_description)

        custom_fields_df = pd.DataFrame(df['customFieldValues'].apply(extract_custom_fields).tolist())
        owner_df = pd.DataFrame(df['owner'].apply(expand_owner).tolist())
        createdby_df = pd.DataFrame(df['createdBy'].apply(expand_createdby).tolist())

        df = df.drop(columns=['customFieldValues', 'owner', 'createdBy', 'actions'])
        df_final = pd.concat([df, owner_df, createdby_df, custom_fields_df], axis=1)

        df_final['execution_timestamp'] = execution_timestamp

        # 🔥 FILTRO CONFORME BASE ESCOLHIDA
        before = len(df_final)

        if tipo_base == "Somente chamados criados pelos usuários autorizados":
            df_final['createdBy_email'] = df_final['createdBy_email'].str.lower().str.strip()
            df_final = df_final[df_final['createdBy_email'].isin(ALLOWED_EMAILS)]
        else:
            df_final['serviceFull'] = df_final['serviceFull'].str.lower().str.strip()
            df_final = df_final[df_final['serviceFull'] == 'agendamento']

        after = len(df_final)

        st.success(f"{after} chamados mantidos de {before} ({tipo_base})")

        # --- SALVAR LOCAL ---
        try:
            if os.name == 'nt' and os.path.exists(os.path.dirname(DOWNLOADS_PATH)):
                df_final.to_csv(DOWNLOADS_PATH, index=False)
                st.success(f"Arquivo salvo em {DOWNLOADS_PATH}")
        except Exception as e:
            st.error(str(e))

        # --- DOWNLOAD ---
        st.download_button(
            "⬇️ Baixar CSV",
            data=df_final.to_csv(index=False).encode("utf-8"),
            file_name="TicketsMovidesk_filtrado.csv",
            mime="text/csv"
        )

        st.dataframe(df_final.head())

    st.balloons()
