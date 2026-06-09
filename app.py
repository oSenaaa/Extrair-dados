import streamlit as st
import json
import pandas as pd
import io

st.set_page_config(page_title="Extrator de Dados Universal", page_icon="📊", layout="wide")

st.title("📊 Extrator de Dados com Filtro Dinâmico")
st.write("Faça o upload do seu arquivo para gerar um Excel totalmente limpo e organizado.")

input_file = st.file_uploader("Escolha o arquivo de dados (JSON, TXT ou sem extensão)")

if input_file is not None:
    try:
        # 1. Ler e carregar o JSON do arquivo
        file_content = input_file.read().decode("utf-8")
        data = json.loads(file_content)
        content_list = data.get('content', []) if isinstance(data, dict) else data
        
        # 2. Criar o DataFrame base
        df_base = pd.json_normalize(content_list)
        
        # 3. Mapear quais campos existem dentro de 'customFields'
        available_custom_fields = set()
        if 'customFields' in df_base.columns:
            for row in df_base['customFields']:
                if isinstance(row, list):
                    for campo in row:
                        if isinstance(campo, dict) and 'name' in campo:
                            available_custom_fields.add(campo['name'])
        
        # 4. Criar os componentes na Interface (UI)
        campos_para_extrair = []
        remover_coluna_original = True
        
        if available_custom_fields:
            st.success(f"🔍 Campos personalizados detetados: {', '.join(available_custom_fields)}")
            
            campos_para_extrair = st.multiselect(
                "Selecione quais campos deseja transformar em colunas próprias:",
                options=sorted(list(available_custom_fields)),
                default=[c for c in ["CID"] if c in available_custom_fields] # Pré-seleciona "CID"
            )
            
            # Opção para deletar a coluna bruta e deixar o Excel perfeito
            remover_coluna_original = st.checkbox(
                "Ocultar/Remover a coluna original 'customFields' do Excel gerado", 
                value=True
            )
        else:
            st.warning("⚠️ Nenhum campo mapeado dentro de 'customFields' foi encontrado.")
        
        # 5. Processar a criação das colunas limpas
        df_final = df_base.copy()
        
        if campos_para_extrair:
            for campo_nome in campos_para_extrair:
                def extrair_valor(lista_de_campos):
                    # Garante a leitura correta mesmo que o formato varie
                    if isinstance(lista_de_campos, list):
                        for c in lista_de_campos:
                            if isinstance(c, dict) and c.get('name') == campo_nome:
                                return c.get('value')
                    elif isinstance(lista_de_campos, str):
                        try:
                            import ast
                            dados = ast.literal_eval(lista_de_campos)
                            if isinstance(dados, list):
                                for c in dados:
                                    if isinstance(c, dict) and c.get('name') == campo_nome:
                                        return c.get('value')
                        except:
                            pass
                    return ""
                
                # NOVIDADE: Define o nome da coluna exatamente como o campo (ex: "CID")
                # Se já existir uma coluna com esse nome por padrão, adiciona um sufixo
                nome_nova_coluna = campo_nome if campo_nome not in df_base.columns else f"{campo_nome}_custom"
                df_final[nome_nova_coluna] = df_final['customFields'].apply(extrair_valor)
        
        # NOVIDADE: Remove a coluna bruta "customFields" para não poluir o Excel do usuário
        if remover_coluna_original and 'customFields' in df_final.columns:
            df_final = df_final.drop(columns=['customFields'])
        
        # 6. Converte outras colunas complexas restantes (se houverem) para texto para evitar erros
        for col in df_final.columns:
            if df_final[col].apply(lambda x: isinstance(x, (list, dict))).any():
                df_final[col] = df_final[col].astype(str)
        
        # 7. Mostrar Pré-visualização e Botão de Download
        st.subheader("👀 Pré-visualização da Planilha Limpa")
        st.dataframe(df_final.head())
        
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_final.to_excel(writer, index=False)
        
        st.download_button(
            label="📥 Baixar Planilha Excel Limpa",
            data=buffer.getvalue(),
            file_name="extracao_perfeita.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.success("Concluído! O CID foi separado numa coluna exclusiva e os dados brutos foram limpos.")
        
    except Exception as e:
        st.error(f"Erro ao processar o arquivo. Detalhes: {e}")