import streamlit as st
import json
import pandas as pd
import io

st.set_page_config(page_title="Extrator de Dados", page_icon="📊")

def extract_data_from_json(uploaded_file):
    """Lê o arquivo (independente da extensão) e extrai as informações."""
    # Lemos os bytes do arquivo e decodificamos em texto (utf-8)
    # Isso evita erros se o arquivo for .txt ou não tiver extensão
    file_content = uploaded_file.read().decode("utf-8")
    data = json.loads(file_content)
    
    extracted_data = []
    for item in data.get('content', []):
        extracted_data.append({
            "uuid": item.get("uuid"),
            "name": item.get("fullName") or item.get("name"),
            "active": item.get("active"),
            "externalId": item.get("externalId")
        })
    
    return extracted_data

# Interface do App
st.title("📊 Extrator de Dados")
st.write("Faça o upload do arquivo de resposta do sistema para convertê-lo em Excel.")

# ALTERAÇÃO AQUI: Sem o 'type', ele aceita .json, .txt ou arquivos sem extensão nenhuma!
input_file = st.file_uploader("Escolha o arquivo de dados")

if input_file is not None:
    try:
        with st.spinner('Processando os dados...'):
            extracted_data = extract_data_from_json(input_file)
            df = pd.DataFrame(extracted_data)
            
            st.subheader("Pré-visualização dos Dados")
            st.dataframe(df.head())
            
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            
            st.download_button(
                label="📥 Baixar Planilha Excel",
                data=buffer.getvalue(),
                file_name="jornada.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.success("Pronto!")
            
    except Exception as e:
        st.error(f"Erro ao processar. Certifique-se de que o conteúdo do arquivo é um JSON válido. Detalhes: {e}")
