import streamlit as st
import pandas as pd


from checkreport import doctable,download_svg_link,download_word_link


def main():
    st.subheader("Convert Table to Word")
    # upload excel file
    file = st.file_uploader("Upload excel file", type=["xlsx"])
    if file is not None:
        # get sheet names list from excel file
        xls = pd.ExcelFile(file)
        sheets = xls.sheet_names
        # choose sheet name and click button
        sheet_name = st.selectbox('Choose sheetname', sheets)
        # choose header row number and click button
        header_row = st.number_input('Choose header row',
                                        min_value=0,
                                        max_value=10,
                                        value=0)

        df = pd.read_excel(file, header=header_row, sheet_name=sheet_name)
        st.write(df.astype(str))
 
        cols = df.columns
        # input title
        title = st.sidebar.text_input('Title')

        # select columns on the left
        hcols = st.sidebar.multiselect('Section column', cols)

        # select columns on the right
        dcols = st.sidebar.multiselect('Content column', cols)

        # click button to generate docx
        if st.sidebar.button('Generate report', key='generate'):
            document,fig = doctable(df, title, hcols, dcols)
            # download docx document
            filelink = download_word_link(document, 'table2doc.docx', 'Download docx')
            st.sidebar.markdown(filelink, unsafe_allow_html=True)
            # show figure
            svglink=download_svg_link(fig, 'graph.svg', 'Download svg')
            st.sidebar.markdown(svglink, unsafe_allow_html=True)


if __name__ == '__main__':
    main()