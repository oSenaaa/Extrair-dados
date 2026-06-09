import streamlit as st
import json
import pandas as pd
import io

st.set_page_config(page_title="Extrator de Dados", page_icon="📊")

def extract_all_data_from_json(uploaded_file):
    """Lê o arquivo e extrai TODAS as colunas dinamicamente."""
    # Lemos os bytes do arquivo e decodificamos
    file_content = uploaded_file.read().decode("utf-8")
    data = json.loads(file_content)
    
    # Identifica se os dados estão dentro de 'content' ou soltos no JSON
    content_list = data.get('content', []) if isinstance(data, dict) else data
    
    # O pd.json_normalize lê a lista e cria colunas para TUDO o que existir lá dentro,
    # inclusive abrindo chaves aninhadas (dados dentro de dados).
    df = pd.json_normalize(content_list)
    
    return df

# Interface do App
st.title("📊 Extrator de Dados Universal")
st.write("Faça o upload do arquivo para convertê-lo em Excel com TODAS as colunas disponíveis.")

input_file = st.file_uploader("Escolha o arquivo de dados")

if input_file is not None:
    try:
        with st.spinner('Processando todas as colunas...'):
            df = extract_all_data_from_json(input_file)
            
            st.subheader("Pré-visualização dos Dados")
            st.dataframe(df.head())
            
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            
            st.download_button(
                label="📥 Baixar Planilha Excel Completa",
                data=buffer.getvalue(),
                file_name="extracao_completa.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.success("Pronto! Todas as colunas foram lidas e convertidas.")
            
    except Exception as e:
        st.error(f"Erro ao processar. Certifique-se de que o conteúdo do arquivo é um JSON válido. Detalhes: {e}")