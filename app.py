import streamlit as st
import pandas as pd
from datetime import datetime
import os

def padronizar_nomes_colunas(df):
    df.columns = [col.strip().upper() for col in df.columns]
    return df

def adicionar_coluna_verificacao(df):
    if 'ID VERIFICACAO' not in df.columns:
        df['ID VERIFICACAO'] = ''
    colunas = list(df.columns)
    colunas.remove('ID VERIFICACAO')
    colunas.append('ID VERIFICACAO')
    return df[colunas]

def valor_valido(valor):
    valor = str(valor).strip() if not pd.isna(valor) else ""
    return valor not in ["", "00:00"]

def processar_planilha(uploaded_file):
    faltas_data = pd.read_excel(uploaded_file)
    faltas_data = padronizar_nomes_colunas(faltas_data)
    faltas_data = adicionar_coluna_verificacao(faltas_data)

    nao_processadas = faltas_data[faltas_data['ID VERIFICACAO'] != 'PROCESSADO']
    registros = []
    colunas_exportadas = []
    data_atual = datetime.now().strftime('%Y-%m-%d')

    for _, row in nao_processadas.iterrows():
        descricao = (
            f"Matrícula: {row.get('MATRÍCULA', '')}\n"
            f"Localização: {row.get('LOCALIZAÇÃO', '')}\n"
            f"Dia: {row.get('DIA', '')}\n"
        )

        batidas_colunas = [
            'BATIDAS', 'ENTRADA 1', 'SAÍDA 1', 'ENTRADA 2', 'SAÍDA 2',
            'ENTRADA 3', 'SAÍDA 3', 'ENTRADA 4', 'SAÍDA 4'
        ]
        
        if all(not valor_valido(row.get(col, '')) for col in batidas_colunas):
            registros.append({
                'list': 'SEM BATIDA',
                'Card Name': row.get('NOME', 'Sem Nome'),
                'desc': descricao,
                'checklist': 'Sem registros de batida',
                'Data': data_atual
            })
            colunas_exportadas.append('SEM BATIDA')

        campos_verificacao = {
            'ATRASO': 'ATRASO',
            'FALTA': 'FALTA',
            'BANCO DE HORAS': 'BANCO DE HORAS',
            'HORA EXTRA 50% (N.A.)': 'HORA EXTRA 50%',
            'HORA EXTRA 100% (N.A.)': 'HORA EXTRA 100%',
            'DSR DESCONTADO': 'DSR DESCONTADO',
            'ADICIONAL NOTURNO': 'ADICIONAL NOTURNO',
            'EXPEDIENTE': 'EXPEDIENTE'
        }

        for campo, lista in campos_verificacao.items():
            if valor_valido(row.get(campo, '')):
                registros.append({
                    'list': lista,
                    'Card Name': row.get('NOME', 'Sem Nome'),
                    'desc': descricao,
                    'checklist': str(row.get(campo, '')).strip(),
                    'Data': data_atual
                })
                colunas_exportadas.append(lista)

        faltas_data.loc[row.name, 'ID VERIFICACAO'] = 'PROCESSADO'

    return pd.DataFrame(registros), faltas_data, list(set(colunas_exportadas))

def main():
    st.title("Automação Trello")
    st.write("Faça upload do arquivo Excel para processar")

    uploaded_file = st.file_uploader("Escolha um arquivo Excel", type=['xlsx'])

    if uploaded_file is not None:
        if st.button("Processar Arquivo"):
            try:
                trello_data, faltas_atualizadas, colunas_exportadas = processar_planilha(uploaded_file)
                
                st.success("Arquivo processado com sucesso!")
                st.write("Colunas exportadas:", ", ".join(colunas_exportadas))

                # Download dos arquivos processados
                st.download_button(
                    label="Baixar arquivo Trello formatado",
                    data=trello_data.to_excel(index=False, engine='openpyxl'),
                    file_name=f"Trello_Formatado_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                st.download_button(
                    label="Baixar planilha atualizada",
                    data=faltas_atualizadas.to_excel(index=False, engine='openpyxl'),
                    file_name=f"Faltas_Atualizadas_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"Erro ao processar arquivo: {str(e)}")

if __name__ == "__main__":
    main()