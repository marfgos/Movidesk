import streamlit as st
import requests
import pandas as pd
from datetime import datetime, timedelta
import os

# --- Caminho desejado (apenas usado se rodar localmente no Windows) ---
DOWNLOADS_PATH = r"C:\Users\MArcos.Silva\Downloads\TicketsMovidesk.csv"

# --- Lista fixa de e-mails que você solicitou (serão os únicos mantidos) ---
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
# Normalize allowed emails (lowercase, strip)
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
        value = field.get('value', None)
        if not value and field.get('items') and isinstance(field['items'], list) and len(field['items']) > 0:
            item = field['items'][0]
            value = item.get('customFieldItem') if isinstance(item, dict) else None
        custom_fields[f'customField_{field_id}'] = value
    return custom_fields

def expand_owner(owner):
    if owner is None:
        return dict.fromkeys(['owner_id', 'owner_personType', 'owner_profileType',
                              'owner_businessName', 'owner_email', 'owner_phone', 'owner_pathPicture'], None)
    return {
        'owner_id': owner.get('id'),
        'owner_personType': owner.get('personType'),
        'owner_profileType': owner.get('profileType'),
        'owner_businessName': owner.get('businessName'),
        'owner_email': owner.get('email'),
        'owner_phone': owner.get('phone'),
        'owner_pathPicture': owner.get('pathPicture')
    }

def expand_createdby(createdby):
    if createdby is None:
        return dict.fromkeys(['createdBy_id', 'createdBy_businessName', 'createdBy_email',
                              'createdBy_phone', 'createdBy_profileType', 'createdBy_personType'], None)
    return {
        'createdBy_id': createdby.get('id'),
        'createdBy_businessName': createdby.get('businessName'),
        'createdBy_email': createdby.get('email'),
        'createdBy_phone': createdby.get('phone'),
        'createdBy_profileType': createdby.get('profileType'),
        'createdBy_personType': createdby.get('personType')
    }

def get_first_action_description(actions):
    if actions and isinstance(actions, list) and len(actions) > 0:
        return actions[0].get('description', None)
    return None

# --- Streamlit app ---
st.title("📊 Coleta de Tickets Movidesk (filtrado por createdBy_email)")

data_inicial = st.date_input(
    "Selecione a data inicial:",
    value=datetime(2025, 6, 1).date(),
    min_value=datetime(2025, 1, 1).date(),
    max_value=datetime.now().date()
)

if st.button("🚀 Extrair, filtrar por e-mails e salvar/baixar CSV"):
    from zoneinfo import ZoneInfo
    execution_timestamp = datetime.now(ZoneInfo("America/Sao_Paulo")).strftime('%d/%m/%Y %H:%M:%S')
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

        for item in all_data:
            if 'actions' not in item:
                item['actions'] = None

        df = pd.DataFrame(all_data)
        df['first_action_description'] = df['actions'].apply(get_first_action_description)

        # safety checks for optional columns
        custom_fields_df = pd.DataFrame()
        owner_fields_df = pd.DataFrame()
        createdBy_fields_df = pd.DataFrame()

        if 'customFieldValues' in df.columns:
            expanded_fields = df['customFieldValues'].apply(extract_custom_fields)
            custom_fields_df = pd.DataFrame(expanded_fields.tolist())

        if 'owner' in df.columns:
            expanded_owners = df['owner'].apply(expand_owner)
            owner_fields_df = pd.DataFrame(expanded_owners.tolist())

        if 'createdBy' in df.columns:
            expanded_createdBy = df['createdBy'].apply(expand_createdby)
            createdBy_fields_df = pd.DataFrame(expanded_createdBy.tolist())

        drop_cols = [c for c in ['owner', 'customFieldValues', 'createdBy', 'actions'] if c in df.columns]
        left_df = df.drop(drop_cols, axis=1) if not df.drop(drop_cols, axis=1).empty else df
        df_final = pd.concat([left_df, owner_fields_df, custom_fields_df, createdBy_fields_df], axis=1)

        if 'createdDate' in df_final.columns:
            df_final['createdDate'] = pd.to_datetime(df_final['createdDate'], errors='coerce').dt.strftime('%d/%m/%Y %H:%M')
        if 'resolvedIn' in df_final.columns:
            df_final['resolvedIn'] = pd.to_datetime(df_final['resolvedIn'], errors='coerce').dt.strftime('%d/%m/%Y %H:%M')

        # rename custom fields if needed (kept minimal)
        df_final = df_final.rename(columns={}, errors='ignore')

        df_final['execution_timestamp'] = execution_timestamp

        # --- FILTRAR por createdBy_email usando a lista ALLOWED_EMAILS ---
        if 'createdBy_email' in df_final.columns:
            # normaliza para comparação segura
            df_final['createdBy_email_norm'] = df_final['createdBy_email'].astype(str).str.strip().str.lower()
            before_count = len(df_final)
            df_final = df_final[df_final['createdBy_email_norm'].isin(ALLOWED_EMAILS)].copy()
            after_count = len(df_final)
            st.success(f"Filtro aplicado: {after_count} chamados mantidos de {before_count} originais.")
            # remove coluna auxiliar
            df_final.drop(columns=['createdBy_email_norm'], inplace=True)
        else:
            st.error("A coluna 'createdBy_email' não foi encontrada nos dados. Nenhum filtro aplicado.")

        # --- TENTAR SALVAR LOCAL (apenas se o caminho existir) ---
        try:
            saved_locally = False
            downloads_dir = os.path.dirname(DOWNLOADS_PATH) if DOWNLOADS_PATH else ""
            if downloads_dir and os.name == 'nt' and os.path.exists(downloads_dir):
                df_final.to_csv(DOWNLOADS_PATH, index=False)
                st.success(f"✅ Arquivo salvo em: {DOWNLOADS_PATH}")
                saved_locally = True
            else:
                st.info("A pasta local de Downloads não está acessível neste ambiente. Use o botão de download.")
        except Exception as e:
            st.error("❌ Falha ao salvar localmente: " + str(e))

        # --- Sempre ofereço um botão de download ---
        try:
            to_download = df_final.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="⬇️ Baixar CSV pelo navegador",
                data=to_download,
                file_name="TicketsMovidesk_filtrado.csv",
                mime="text/csv"
            )
        except Exception as e:
            st.error("Erro ao criar botão de download: " + str(e))

        # mostra um trecho
        st.dataframe(df_final.head())

    st.balloons()
