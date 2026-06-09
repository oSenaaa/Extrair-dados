import streamlit as st
import json
import pandas as pd
import io

st.set_page_config(page_title="Extrator de Dados Universal", page_icon="📊", layout="wide")

st.title("📊 Extrator de Dados com Filtro Dinâmico")
st.write("Faça o upload do seu arquivo. O sistema identificará os campos personalizados automaticamente para você escolher.")

input_file = st.file_uploader("Escolha o arquivo de dados (JSON, TXT ou sem extensão)")

if input_file is not None:
    try:
        # 1. Ler e carregar o JSON do arquivo
        file_content = input_file.read().decode("utf-8")
        data = json.loads(file_content)
        content_list = data.get('content', []) if isinstance(data, dict) else data
        
        # 2. Criar o DataFrame base com todas as colunas normais
        df_base = pd.json_normalize(content_list)
        
        # 3. Mapear DINAMICAMENTE quais campos existem dentro de 'customFields'
        available_custom_fields = set()
        if 'customFields' in df_base.columns:
            for row in df_base['customFields']:
                if isinstance(row, list):
                    for campo in row:
                        if isinstance(campo, dict) and 'name' in campo:
                            available_custom_fields.add(campo['name'])
        
        # 4. Criar o componente na Interface (UI) do Streamlit
        campos_para_extrair = []
        if available_custom_fields:
            st.info(f"💡 Campos personalizados detectados neste arquivo: {', '.join(available_custom_fields)}")
            
            # O multiselect permite escolher um ou vários ao mesmo tempo
            campos_para_extrair = st.multiselect(
                "Selecione quais campos personalizados você deseja transformar em colunas no Excel:",
                options=sorted(list(available_custom_fields)),
                default=[c for c in ["CID"] if c in available_custom_fields] # Pré-seleciona "CID" se ele existir
            )
        else:
            st.warning("⚠️ Nenhum campo mapeado dentro de 'customFields' foi encontrado neste arquivo específico.")
        
        # 5. Processar a criação das colunas com base na escolha da UI
        df_final = df_base.copy()
        
        if campos_para_extrair:
            for campo_nome in campos_para_extrair:
                def extrair_valor(lista_de_campos):
                    if isinstance(lista_de_campos, list):
                        for c in lista_de_campos:
                            if isinstance(c, dict) and c.get('name') == campo_nome:
                                return c.get('value')
                    return ""
                
                # Cria a coluna nova na planilha com o nome do campo escolhido
                df_final[f'Personalizado_{campo_nome}'] = df_final['customFields'].apply(extrair_valor)
        
        # 6. Melhoria de segurança: Converte APENAS colunas complexas (listas/dicionários) para texto.
        # Isso evita o erro do Excel sem transformar seus números reais em texto!
        for col in df_final.columns:
            if df_final[col].apply(lambda x: isinstance(x, (list, dict))).any():
                df_final[col] = df_final[col].astype(str)
        
        # 7. Mostrar Pré-visualização e Botão de Download
        st.subheader("👀 Pré-visualização da Planilha")
        st.dataframe(df_final.head())
        
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_final.to_excel(writer, index=False)
        
        st.download_button(
            label="📥 Baixar Planilha Excel Customizada",
            data=buffer.getvalue(),
            file_name="extracao_customizada.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.success("Pronto! Planilha gerada com as suas escolhas.")
        
    except Exception as e:
        st.error(f"Erro ao processar o arquivo. Detalhes: {e}")