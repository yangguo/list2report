import streamlit as st
import pandas as pd


from checkreport import doctable,download_svg_link,download_word_link


def main():

    st.subheader("发现清单转文档")
    # upload excel file
    file = st.file_uploader("选择文件", type=["xlsx"])
    if file is not None:
        # get sheet names list from excel file
        xls = pd.ExcelFile(file)
        sheets = xls.sheet_names
        # choose sheet name and click button
        sheet_name = st.selectbox('选择工作表', sheets)
        # choose header row number and click button
        header_row = st.number_input('选择表头行',
                                        min_value=0,
                                        max_value=10,
                                        value=0)

        df = pd.read_excel(file, header=header_row, sheet_name=sheet_name)
        st.write(df.astype(str))
 
        cols = df.columns
        # input title
        title = st.sidebar.text_input('输入标题')

        # select columns on the left
        hcols = st.sidebar.multiselect('选择结构字段', cols)

        # select columns on the right
        dcols = st.sidebar.multiselect('选择数据字段', cols)

        # click button to generate docx
        if st.sidebar.button('生成文档', key='generate'):
            document,fig = doctable(df, title, hcols, dcols)
            # download docx document
            filelink = download_word_link(document, '发现清单转文档.docx', '下载文档')
            st.sidebar.markdown(filelink, unsafe_allow_html=True)
            # show figure
            svglink=download_svg_link(fig, '发现清单.svg', '下载图片')
            st.sidebar.markdown(svglink, unsafe_allow_html=True)


if __name__ == '__main__':
    main()