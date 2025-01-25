import streamlit as st
import pandas as pd
import xlsxwriter
import io

def main():
    st.title("Efficient Spreadsheet Organizer")

    # Upload file
    uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx", "xls"])

    if uploaded_file is not None:
        # Read the file
        df = pd.read_excel(uploaded_file)

        # Sidebar for selecting operations
        st.sidebar.header("Step 1: Select Operations")

        # Column reordering
        st.sidebar.subheader("Reorder Columns")
        new_column_order = st.sidebar.multiselect(
            "Select and reorder columns", 
            df.columns.tolist(), 
            default=df.columns.tolist()
        )

        # Select columns to filter
        st.sidebar.subheader("Select Columns to Filter")
        filter_columns = st.sidebar.multiselect("Choose columns for filtering", df.columns.tolist())

        # Select columns to group
        st.sidebar.subheader("Select Columns to Group By")
        group_columns = st.sidebar.multiselect("Choose columns for grouping", df.columns.tolist())

        # Select columns to sort
        st.sidebar.subheader("Select Columns to Sort By")
        sort_columns = st.sidebar.multiselect("Choose columns for sorting", df.columns.tolist())

        st.sidebar.subheader("Step 2: Configure Options")

        # Apply column reordering
        if new_column_order:
            df = df[new_column_order]

        # Dynamic filtering options
        filters = {}
        for col in filter_columns:
            if df[col].dtype == 'object' or df[col].dtype == 'category':
                filters[col] = st.sidebar.multiselect(f"Filter {col}", df[col].unique(), default=df[col].unique())
            else:
                min_val, max_val = st.sidebar.slider(
                    f"Filter {col}", 
                    float(df[col].min()), 
                    float(df[col].max()), 
                    (float(df[col].min()), float(df[col].max()))
                )
                filters[col] = (min_val, max_val)

        # Apply filters
        for col, values in filters.items():
            if isinstance(values, tuple):
                df = df[(df[col] >= values[0]) & (df[col] <= values[1])]
            else:
                df = df[df[col].isin(values)]

        # Apply grouping
        if group_columns:
            df = df.groupby(group_columns).sum().reset_index()

        # Dynamic sorting options
        if sort_columns:
            sort_orders = [st.sidebar.radio(f"Sort {col}", ["Ascending", "Descending"]) == "Ascending" for col in sort_columns]
            df = df.sort_values(by=sort_columns, ascending=sort_orders)

        # Display the processed data
        st.subheader("Processed Data")
        st.dataframe(df)

        # Download the processed data
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name="Processed Data")
        processed_data = output.getvalue()

        st.download_button(
            label="Download Processed Excel File",
            data=processed_data,
            file_name="processed_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()
