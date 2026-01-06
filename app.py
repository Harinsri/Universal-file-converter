import streamlit as st
from markitdown import MarkItDown
import os

# Initialize the MarkItDown engine
# Note: MarkItDown handles Word, Excel, PPT, PDF, and HTML natively
mid = MarkItDown()

# Page configuration
st.set_page_config(page_title="Universal Doc Converter", page_icon="📄")

def main():
    st.title("📄 Universal Document Reader")
    st.markdown("Convert Office docs, PDFs, and HTML into clean Markdown instantly.")

    # [2] Upload Area - Multiple files supported
    uploaded_files = st.file_uploader(
        "Drag and drop files here", 
        type=["docx", "xlsx", "pptx", "pdf", "html", "zip"],
        accept_multiple_files=True
    )

    if uploaded_files:
        for uploaded_file in uploaded_files:
            file_name = uploaded_file.name
            base_name = os.path.splitext(file_name)[0]

            with st.expander(f"👁️ Preview: {file_name}", expanded=True):
                try:
                    # [3] Resilience & Processing
                    # MarkItDown can process file-like objects (BytesIO)
                    # We pass the file stream directly to the engine
                    result = mid.convert(uploaded_file)
                    converted_text = result.text_content

                    # [2] Instant Preview
                    st.text_area(
                        label="Converted Content",
                        value=converted_text,
                        height=300,
                        key=f"text_{file_name}"
                    )

                    # [2] & [4] Download Options
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.download_button(
                            label="Download as Markdown (.md)",
                            data=converted_text,
                            file_name=f"{base_name}_converted.md",
                            mime="text/markdown",
                            key=f"md_{file_name}"
                        )
                    
                    with col2:
                        st.download_button(
                            label="Download as Text (.txt)",
                            data=converted_text,
                            file_name=f"{base_name}_converted.txt",
                            mime="text/plain",
                            key=f"txt_{file_name}"
                        )

                except Exception as e:
                    # [3] Error Handling
                    st.error(f"⚠️ Could not read {file_name}. Please check the format.")
                    # Log the error for the developer (optional)
                    # st.write(f"Debug Info: {str(e)}")

if __name__ == "__main__":
    main()
