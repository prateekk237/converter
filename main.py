import streamlit as st
from pdf2pptx import convert

def main():
    st.title("PDF to PPT Converter (Using pdf2pptx)")

    uploaded_file = st.file_uploader("Upload a PDF file", type=["pdf"])

    if uploaded_file is not None:
        # Save the uploaded PDF to a file
        with open("input.pdf", "wb") as f:
            f.write(uploaded_file.getbuffer())

        # Convert the PDF to PowerPoint
        output_pptx = "output.pptx"
        convert("input.pdf", output_pptx)

        # Provide download button for the generated PowerPoint
        with open(output_pptx, "rb") as ppt_file:
            st.download_button(
                label="Download PowerPoint",
                data=ppt_file,
                file_name="output.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )

if __name__ == "__main__":
    main()
