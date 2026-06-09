import streamlit as st
import json
import pandas as pd
import io
import ast

st.set_page_config(page_title="Extrator de Dados Universal", page_icon="📊", layout="wide")

st.title("📊 Extrator de Dados com Filtro Dinâmico")
st.write("Faça o upload do seu arquivo para gerar um Excel totalmente limpo e organizado.")

input_file = st.file_uploader("Escolha o arquivo de dados (JSON, TXT ou sem extensão)")

if input_file is not None:
    try:
        # 🛡️ SEGURANÇA 1: Garante que o ponteiro de leitura do arquivo volte para o início
        input_file.seek(0)
        
        # 🛡️ SEGURANÇA 2: Cria uma chave única para o arquivo para salvar na memória (Session State)
        # Isso impede que o Streamlit re-mude os dados e dê reset nas caixas ao clicar em Baixar
        file_id = f"{input_file.name}_{input_file.size}"
        
        if "current_file" not in st.session_state or st.session_state["current_file"] != file_id:
            file_content = input_file.read().decode("utf-8")
            data = json.loads(file_content)
            content_list = data.get('content', []) if isinstance(data, dict) else data
            
            # Guardamos o DataFrame bruto na memória do navegador
            st.session_state["df_base"] = pd.json_normalize(content_list)
            st.session_state["current_file"] = file_id
            
        # Recuperamos o DataFrame estável da memória
        df_base = st.session_state["df_base"]
        
        # 3. Mapear quais campos existem dentro de 'customFields'
        available_custom_fields = set()
        if 'customFields' in df_base.columns:
            for row in df_base['customFields']:
                if isinstance(row, list):
                    for campo in row:
                        if isinstance(campo, dict) and 'name' in campo:
                            available_custom_fields.add(campo['name'])
        
        df_final = df_base.copy()
        
        # 4. Extração de Campos Personalizados (ex: CID)
        if available_custom_fields:
            st.success(f"🔍 Campos personalizados detectados: {', '.join(available_custom_fields)}")
            
            # Adicionada uma chave fixa (key) para salvar o estado do componente
            campos_para_extrair = st.multiselect(
                "1️⃣ Selecione quais campos personalizados deseja transformar em colunas:",
                options=sorted(list(available_custom_fields)),
                default=[c for c in ["CID"] if c in available_custom_fields],
                key="selecao_campos_custom"
            )
            
            remover_coluna_original = st.checkbox(
                "Ocultar a coluna original 'customFields' para limpar os dados brutos", 
                value=True,
                key="check_remover_original"
            )
            
            if campos_para_extrair:
                for campo_nome in campos_para_extrair:
                    def extrair_valor(lista_de_campos):
                        if isinstance(lista_de_campos, list):
                            for c in lista_de_campos:
                                if isinstance(c, dict) and c.get('name') == campo_nome:
                                    return c.get('value')
                        elif isinstance(lista_de_campos, str):
                            try:
                                dados = ast.literal_eval(lista_de_campos)
                                if isinstance(dados, list):
                                    for c in dados:
                                        if isinstance(c, dict) and c.get('name') == campo_nome:
                                            return c.get('value')
                            except:
                                pass
                        return ""
                    
                    nome_nova_coluna = campo_nome if campo_nome not in df_base.columns else f"{campo_nome}_custom"
                    df_final[nome_nova_coluna] = df_final['customFields'].apply(extrair_valor)
            
            if remover_coluna_original and 'customFields' in df_final.columns:
                df_final = df_final.drop(columns=['customFields'])
        
        # 5. Filtro mestre de colunas para exportação
        st.divider()
        st.write("### 🛠️ Ajuste Fino da Planilha")
        
        todas_as_colunas = df_final.columns.tolist()
        
        # 🛡️ SEGURANÇA 3: Adicionada a propriedade `key`. 
        # Isso amarra as colunas desmarcadas na memória e o download respeitará o filtro!
        colunas_selecionadas = st.multiselect(
            "2️⃣ Escolha exatamente quais colunas você quer manter no Excel final:",
            options=todas_as_colunas,
            default=todas_as_colunas,
            key="filtro_colunas_finais",
            help="Clique no 'X' para remover as colunas que você não quer exportar."
        )
        
        # Aplica o filtro selecionado pelo usuário
        if colunas_selecionadas:
            df_export = df_final[colunas_selecionadas]
        else:
            df_export = df_final
            st.warning("⚠️ Nenhuma coluna foi selecionada. Mostrando todas por padrão.")
        
        # 6. Converte colunas complexas restantes para texto para evitar erro no Excel
        for col in df_export.columns:
            if df_export[col].apply(lambda x: isinstance(x, (list, dict))).any():
                df_export[col] = df_export[col].astype(str)
        
        # 7. Mostrar Pré-visualização Exata
        st.subheader("👀 Pré-visualização Exata")
        st.write("A planilha gerada será 100% igual ao que está aparecendo abaixo:")
        st.dataframe(df_export)
        
        # Gerando o Excel estritamente com o DF filtrado
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_export.to_excel(writer, index=False)
        
        st.download_button(
            label="📥 Baixar Planilha Excel Escolhida",
            data=buffer.getvalue(),
            file_name="extracao_perfeita.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="botao_download_final"
        )
        
    except Exception as e:
        st.error(f"Erro ao processar o arquivo. Detalhes: {e}")