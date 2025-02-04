import streamlit as st
import pandas as pd
from datetime import datetime
import io
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

def save_files(trello_data, faltas_atualizadas, save_path):
    try:
        if not os.path.exists(save_path):
            os.makedirs(save_path)
            
        trello_file = os.path.join(save_path, f"Trello_Formatado_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        faltas_file = os.path.join(save_path, f"Faltas_Atualizadas_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        
        trello_data.to_excel(trello_file, index=False)
        faltas_atualizadas.to_excel(faltas_file, index=False)
        return trello_file, faltas_file
    except Exception as e:
        st.error(f"Erro ao salvar arquivos: {e}")
        return None, None

def main():
    st.set_page_config(page_title="Automação Trello", layout="wide")
    
    st.title("Automação Trello")
    
    col1, col2 = st.columns(2)
    
    with col1:
        save_path = st.text_input(
            "Pasta para salvar arquivos:",
            value="C:/projeto_atendimento/Pronto_para_Trello",
            help="Digite o caminho completo da pasta onde deseja salvar os arquivos"
        )
    
    with col2:
        uploaded_file = st.file_uploader(
            "Escolha um arquivo Excel",
            type=['xlsx'],
            help="Selecione o arquivo Excel com os dados para processamento"
        )

    if uploaded_file is not None:
        if st.button("Processar Arquivo", type="primary"):
            with st.spinner("Processando arquivo..."):
                try:
                    trello_data, faltas_atualizadas, colunas_exportadas = processar_planilha(uploaded_file)
                    
                    # Salvar arquivos localmente
                    trello_file, faltas_file = save_files(trello_data, faltas_atualizadas, save_path)
                    
                    if trello_file and faltas_file:
                        st.success("✅ Arquivos salvos com sucesso!")
                        st.info(f"📁 Local dos arquivos:\n- {trello_file}\n- {faltas_file}")
                        
                        st.write("📊 Colunas exportadas:", ", ".join(colunas_exportadas))
                        
                        # Preparar buffers para download
                        buffer_trello = io.BytesIO()
                        trello_data.to_excel(buffer_trello, index=False, engine='openpyxl')
                        buffer_trello.seek(0)

                        buffer_faltas = io.BytesIO()
                        faltas_atualizadas.to_excel(buffer_faltas, index=False, engine='openpyxl')
                        buffer_faltas.seek(0)

                        col1, col2 = st.columns(2)
                        with col1:
                            st.download_button(
                                "⬇️ Download arquivo Trello",
                                data=buffer_trello,
                                file_name=f"Trello_Formatado_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                        
                        with col2:
                            st.download_button(
                                "⬇️ Download planilha atualizada",
                                data=buffer_faltas,
                                file_name=f"Faltas_Atualizadas_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                            
                except Exception as e:
                    st.error(f"❌ Erro ao processar arquivo: {str(e)}")

    st.sidebar.header("Sobre")
    st.sidebar.info(
        """
        Este aplicativo processa planilhas de faltas e gera arquivos formatados para o Trello.
        
        Como usar:
        1. Digite o caminho da pasta para salvar os arquivos
        2. Faça upload do arquivo Excel
        3. Clique em "Processar Arquivo"
        4. Faça download dos arquivos processados
        """
    )

if __name__ == "__main__":
    main()
