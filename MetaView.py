import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import io

st.set_page_config(layout="wide")
data_nome_anexo = datetime.now().strftime("%d.%m.%Y")

def lerexcel(dia, arquivo, arquivo_guia):
    df = pd.read_excel(arquivo, header=None)
    df_guia = pd.read_excel(arquivo_guia)
    quantidade_digitado = dia
    df.columns = df.iloc[2]
    df = df.iloc[3:]
    df = df[['INSTRUTOR', 'COD. LOJA/LOCAL', 'STATUS', 'FORMATO', 'AVALIAÇÃO TREINAMENTO',
             'AVALIAÇÃO CONHECIMENTO','LOJA/LOCAL', 'PART. INSTRUTOR']]
    df["AVALIAÇÃO CONHECIMENTO"] = df["AVALIAÇÃO CONHECIMENTO"].astype(float).fillna(0)
    df_guia['Lojas meta'] = df_guia['Lojas/dia'] * quantidade_digitado
    df_guia['Pessoas meta'] = df_guia['Pessoas/dia'] * quantidade_digitado
    return df, df_guia

def meta1(df, df_guia, status_selecionados, formatos_selecionados, filtrar_por_id):
    if not status_selecionados or not formatos_selecionados:
        return pd.DataFrame(columns=['Instrutor', 'Lojas por Dia', 'Meta', 'Lojas Treinadas', '%'])
    if filtrar_por_id:
        df = df.drop_duplicates(subset='COD. LOJA/LOCAL')
    df = df[df['STATUS'].isin(status_selecionados)]
    df = df[df['FORMATO'].isin(formatos_selecionados)]
    if df.empty:
        return pd.DataFrame(columns=['Instrutor', 'Lojas por Dia', 'Meta', 'Lojas Treinadas', '%'])
    df = df[df['LOJA/LOCAL'] != "Home Office"]
    df_meta = df.groupby('INSTRUTOR')["LOJA/LOCAL"].count().reset_index()
    df_final = df_meta.merge(df_guia, on='INSTRUTOR', how='outer')
    df_final['%'] = df_final['LOJA/LOCAL'] / df_final['Lojas meta'] * 100
    df_final['%'] = df_final['%'].astype(float).round(0)
    df_final = df_final[['INSTRUTOR', 'Lojas/dia', 'Lojas meta', 'LOJA/LOCAL', '%']]
    df_final.columns = ['Instrutor', 'Lojas por Dia', 'Meta', 'Lojas Treinadas', '%']
    return df_final

def meta2(df, df_guia, status_selecionados, formatos_selecionados, filtrar_por_id):
    if not status_selecionados or not formatos_selecionados:
        return pd.DataFrame(columns=['Instrutor', 'Pessoas por Dia', 'Meta', 'Pessoas Treinadas', '%'])
    if filtrar_por_id:
        df = df.drop_duplicates(subset='COD. LOJA/LOCAL')
    df = df[df['STATUS'].isin(status_selecionados)]
    df = df[df['FORMATO'].isin(formatos_selecionados)]
    if df.empty:
        return pd.DataFrame(columns=['Instrutor', 'Pessoas por Dia', 'Meta', 'Pessoas Treinadas', '%'])
    df_meta = df.groupby('INSTRUTOR')["PART. INSTRUTOR"].sum().reset_index()
    df_merged = df_meta.merge(df_guia, on='INSTRUTOR', how='outer')
    df_merged['%'] = df_merged['PART. INSTRUTOR'] / df_merged['Pessoas meta'] * 100
    df_merged['%'] = df_merged['%'].astype(float).round(0)
    df_merged = df_merged[['INSTRUTOR', 'Pessoas/dia', 'Pessoas meta', 'PART. INSTRUTOR', '%']]
    df_merged.columns = ['Instrutor', 'Pessoas por Dia', 'Meta', 'Pessoas Treinadas', '%']
    return df_merged

def meta3(df, status_selecionados, filtrar_por_id):
    if not status_selecionados:
        return pd.DataFrame(columns=['Instrutor', 'Media de AVALIAÇÃO TREINAMENTO'])
    if filtrar_por_id:
        df = df.drop_duplicates(subset='COD. LOJA/LOCAL')
    df = df[df['FORMATO'].isin(status_selecionados)]
    df = df[df['AVALIAÇÃO TREINAMENTO'].notnull()]
    if df.empty:
        return pd.DataFrame(columns=['Instrutor', 'Media de AVALIAÇÃO TREINAMENTO'])
    df['AVALIAÇÃO TREINAMENTO'] = pd.to_numeric(df['AVALIAÇÃO TREINAMENTO'], errors='coerce').astype(float)
    df_meta = df.groupby('INSTRUTOR')['AVALIAÇÃO TREINAMENTO'].mean().reset_index()
    df_meta['AVALIAÇÃO TREINAMENTO'] = df_meta['AVALIAÇÃO TREINAMENTO'].round(2)
    df_meta.columns = ['Instrutor', 'Media de AVALIAÇÃO TREINAMENTO']
    return df_meta

def meta4(df, status_selecionados, filtrar_por_id):
    if not status_selecionados:
        return pd.DataFrame(columns=['INSTRUTOR', 'FORMATO', 'Quantidade'])
    df = df[df['STATUS'].isin(status_selecionados)]
    if filtrar_por_id:
        df = df.drop_duplicates(subset='COD. LOJA/LOCAL')
    if df.empty:
        return pd.DataFrame(columns=['INSTRUTOR', 'FORMATO', 'Quantidade'])
    try:
        df_grouped = df.groupby(['INSTRUTOR', 'FORMATO']).size().reset_index(name='Quantidade')
        if df_grouped.empty:
            return pd.DataFrame(columns=['INSTRUTOR', 'FORMATO', 'Quantidade'])
        df_pivot = df_grouped.pivot(index='FORMATO', columns='INSTRUTOR', values='Quantidade').fillna(0)
        df_pivot['Total Geral'] = df_pivot.sum(axis=1)
        totais = df_pivot.sum(axis=0)
        df_pivot.loc['Total'] = totais
        return df_pivot
    except Exception as e:
        st.error(f"Erro ao criar pivot table: {e}")
        return pd.DataFrame(columns=['INSTRUTOR', 'FORMATO', 'Quantidade'])

def meta5(df, status_selecionados, formatos_selecionados, filtrar_por_id):
    if filtrar_por_id:
        df = df.drop_duplicates(subset='COD. LOJA/LOCAL')
    df = df[df['STATUS'].isin(status_selecionados)]
    df = df[df['FORMATO'].isin(formatos_selecionados)]
    df = df[['INSTRUTOR', 'LOJA/LOCAL', 'AVALIAÇÃO CONHECIMENTO']]
    df1 = df.copy()
    df = df[df['LOJA/LOCAL'] != 'Home Office']
    df['AVALIAÇÃO CONHECIMENTO'] = df['AVALIAÇÃO CONHECIMENTO'].fillna(0).astype(float)
    df['Quantidade de Avaliacoes'] = np.where(df['AVALIAÇÃO CONHECIMENTO'] > 0, 1, 0)
    instrutor_counts = df['INSTRUTOR'].value_counts().copy()
    df = df.groupby(['INSTRUTOR', 'LOJA/LOCAL'])['Quantidade de Avaliacoes'].sum().reset_index()
    df = df.groupby('INSTRUTOR')['Quantidade de Avaliacoes'].sum().reset_index()
    df['Quantidade de lojas treinadas'] = df['INSTRUTOR'].map(instrutor_counts)
    df['%'] = (df['Quantidade de Avaliacoes'] / df['Quantidade de lojas treinadas']) * 100
    df['%'] = df['%'].round(2)
    df = df.sort_values(by='INSTRUTOR', ascending=True)
    df1 = df1.sort_values(by='INSTRUTOR', ascending=True)
    return df1, df

st.image("Logo Saiba+ Brastemp novo.png", width=350)

st.markdown("""
    <style>
        .block-container {
            max-width: 1500px;
            padding-left: 2rem;
            padding-right: 2rem;
        }
    </style>
""", unsafe_allow_html=True)

dias = st.number_input("Digite a quantidade de dias úteis:", min_value=1, step=1)
st.success(f"Você escolheu {dias} dias úteis.")

st.markdown("""<h3 style='font-size: 24px; color: #58595b;'>📁 Insira o arquivo Excel dos <strong>DADOS</strong></h3>""", unsafe_allow_html=True)
arquivo = st.file_uploader("Arquivo de dados", type=["xls", "xlsx"], key="dados_file", label_visibility="collapsed")

st.markdown("""<h3 style='font-size: 24px; color: #58595b;'>📋 Insira o arquivo Excel do <strong>GUIA INSTRUTORES</strong></h3>""", unsafe_allow_html=True)
arquivo_guia = st.file_uploader("Guia de instrutores", type=["xls", "xlsx"], key="guia_file", label_visibility="collapsed")

if arquivo and arquivo_guia:
    try:
        df, df_guia = lerexcel(dias, arquivo, arquivo_guia)
        if df.empty or df_guia.empty:
            st.error("Um ou ambos os arquivos estão vazios ou não contêm dados válidos.")
            st.stop()
    except Exception as e:
        st.error(f"Erro ao ler arquivos: {e}")
        st.stop()

    status_opcoes_formato = ['ADM', 'Evento', 'On The Job', 'Online', 'Presencial']
    status_opcoes_status = ['Realizado', 'No-show', 'Cancelado']

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("META POR LOJA")
        filtrar1 = st.checkbox("Filtrar IDs únicos - Meta 1")
        status1 = st.multiselect("Status para META 1:", status_opcoes_status, default=status_opcoes_status)
        formatos_selecionados1 = st.multiselect("Formatos para META 1:", status_opcoes_formato, default=status_opcoes_formato)
        df_meta1 = meta1(df, df_guia, status1, formatos_selecionados1, filtrar1)
        if not df_meta1.empty:
            st.dataframe(df_meta1, hide_index=True)
        else:
            st.warning("Nenhum dado encontrado para META 1.")

        st.subheader("MÉDIA NPS")
        filtrar3 = st.checkbox("Filtrar IDs únicos - Meta 3")
        status3 = st.multiselect("Status para META 3:", status_opcoes_formato, default=status_opcoes_formato)
        df_meta3 = meta3(df, status3, filtrar3)
        if not df_meta3.empty:
            st.dataframe(df_meta3, hide_index=True)
        else:
            st.warning("Nenhum dado encontrado para META 3.")

    with col2:
        st.subheader("META POR PESSOAS TREINADAS")
        filtrar2 = st.checkbox("Filtrar IDs únicos - Meta 2")
        status2 = st.multiselect("Status para META 2:", status_opcoes_status, default=status_opcoes_status)
        formatos_selecionados2 = st.multiselect("Formatos para META 2:", status_opcoes_formato, default=status_opcoes_formato)
        df_meta2 = meta2(df, df_guia, status2, formatos_selecionados2, filtrar2)
        if not df_meta2.empty:
            st.dataframe(df_meta2, hide_index=True)
        else:
            st.warning("Nenhum dado encontrado para META 2.")

        st.subheader("LOJAS POR FORMATO")
        filtrar4 = st.checkbox("Filtrar IDs únicos - Meta 4")
        status4 = st.multiselect("Status para META 4:", status_opcoes_status, default=status_opcoes_status)
        df_meta4 = meta4(df, status4, filtrar4)
        if not df_meta4.empty:
            st.dataframe(df_meta4)
        else:
            st.warning("Nenhum dado encontrado para a META 4 com os filtros aplicados.")

    st.subheader("AVALIAÇÃO DE CONHECIMENTO")
    filtrar5 = st.checkbox("Filtrar IDs únicos - Meta 5")
    status5 = st.multiselect("Status para META 5:", status_opcoes_status, default=status_opcoes_status)
    formatos_selecionados5 = st.multiselect("Formatos para META 5:", status_opcoes_formato, default=status_opcoes_formato)
    df_meta51, df_meta5 = meta5(df, status5, formatos_selecionados5, filtrar5)

    if not df_meta5.empty:
        st.dataframe(df_meta5, hide_index=True)
        instrutor_selecionado = st.selectbox("Selecione um Instrutor", df_meta5['INSTRUTOR'].unique())
        df_filtrado = df_meta51[df_meta51['INSTRUTOR'] == instrutor_selecionado]
        st.dataframe(df_filtrado, hide_index=True)
    else:
        st.warning("Nenhum dado encontrado para META 5.")

    with io.BytesIO() as buffer:
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='DADOS', index=False)
            if not df_meta1.empty:
                df_meta1.to_excel(writer, sheet_name='META 1', index=False)
            if not df_meta2.empty:
                df_meta2.to_excel(writer, sheet_name='META 2', index=False)
            if not df_meta3.empty:
                df_meta3.to_excel(writer, sheet_name='META 3', index=False)
            if not df_meta4.empty:
                df_meta4.to_excel(writer, sheet_name='META 4')
            if not df_meta5.empty:
                df_meta5.to_excel(writer, sheet_name='META 5', index=False)
            if not df_meta51.empty:
                df_meta51.to_excel(writer, sheet_name='META 5.1', index=False)
        buffer.seek(0)
        st.download_button(
            "📥 Baixar Excel", 
            buffer, 
            file_name=f"Dashboard_{data_nome_anexo}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

st.markdown("""
---
<div style='text-align: center; font-size: 20px; color: #58595b;'>
    Desenvolvido por Zoom Educação Corporativa © 2025
</div>
""", unsafe_allow_html=True)
