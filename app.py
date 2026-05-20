import streamlit as st
import os
from book_writer import BookWriter

st.set_page_config(page_title="Escritor de Livros KDP", page_icon="📚")

st.title("📚 Escritor de Livros Automático")
st.markdown("""
Transforme seu tema em um livro completo formatado para Amazon KDP (6"x9").
Basta digitar o tema abaixo e clicar em **Gerar Livro**.
""")

with st.sidebar:
    st.header("Configurações")
    user_api_key = st.text_input("OpenAI API Key (Opcional)", type="password", help="Se fornecido, o livro será escrito por IA. Caso contrário, será usado um texto de exemplo.")

theme = st.text_input("Qual o tema do seu livro?", placeholder="Ex: O Guia Definitivo da Culinária Vegana")

if st.button("🚀 Gerar Livro"):
    if not theme:
        st.error("Por favor, insira um tema para o livro.")
    else:
        try:
            # Use the key provided by the user in this session, or from environment if not provided
            effective_api_key = user_api_key if user_api_key else os.getenv("OPENAI_API_KEY")
            
            writer = BookWriter(theme, api_key=effective_api_key)
            
            with st.status("Gerando seu livro... Isso pode levar alguns minutos.", expanded=True) as status:
                st.write("Processando conteúdo...")
                writer.create_book()
                status.update(label="✅ Livro Gerado com Sucesso!", state="complete", expanded=False)
            
            # Use in-memory buffer to avoid file collisions on server
            book_buffer = writer.save_to_buffer()
            
            st.download_button(
                label="📥 Baixar Arquivo DOCX",
                data=book_buffer,
                file_name=writer.filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            st.success(f"O livro '{theme}' foi criado com sucesso!")
            
        except Exception as e:
            st.error(f"Ocorreu um erro ao gerar o livro: {e}")

st.divider()
st.info("Nota: O arquivo gerado já está configurado no tamanho 6\"x9\" (padrão Amazon KDP) com margens e fontes apropriadas.")
