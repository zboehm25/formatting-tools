import pandas as pd
import streamlit as st
import io
from dateutil import parser

def format_qb_tb(input_file):
    """
    Formats a Trial Balance Report from QuickBooks Online.
    Report must be ran for the date range and have columns separated by month.
    Important to do before using this tool is to open the excel file and copy-paste everything as values.
    QBO exports reports as formulas (=number) and won't be read unless pasted as a value.
    """
    df = pd.read_excel(input_file, skiprows=3)
    df.iloc[0] = df.iloc[0].ffill()
    df.iloc[0] = df.iloc[0].astype(str) + ' ' + df.iloc[1].astype(str)
    df = df.drop(1).reset_index(drop=True)
    df.iloc[0, 0] = 'Account'
    df = df.fillna(0)
    df.columns = df.iloc[0]
    df = df.drop(0).reset_index(drop=True)

    total_index = df[df.iloc[:, 0] == 'TOTAL'].index
    if not total_index.empty:
        df = df.iloc[:total_index[0]]
        
    return df

def calculate_activity(df, last_bal_sheet_acc):
    months = []
    for col in df.columns:
        if 'Debit' in col or 'Credit' in col:
            month_year = ' '.join(col.split(' ')[:-1])
            if month_year not in months:
                months.append(month_year)

    for month in months:
        debit_col = f'{month} Debit'
        credit_col = f'{month} Credit'
        df[f'{month} Ending Balance'] = df[debit_col] - df[credit_col]
        if month != months[0]:
            previous_month = months[months.index(month) - 1]
            df[f'{month} Activity'] = df[f'{month} Ending Balance'] - df[f'{previous_month} Ending Balance']

    new_columns = []
    for i in range(len(months)):
        month = months[i]
        new_columns.append(f'{month} Debit')
        new_columns.append(f'{month} Credit')
        new_columns.append(f'{month} Ending Balance')
        if i > 0:
            new_columns.append(f'{month} Activity')

    df = df[['Account'] + new_columns]

    specified_value = last_bal_sheet_acc
    specified_index = df.index[df['Account'] == specified_value].tolist()[0]

    result_df = pd.DataFrame(df['Account'])
    for i, month in enumerate(months):
        ending_balance_col = f'{month} Ending Balance'
        activity_col = f'{month} Activity'
        result_df[month] = None
        if i == 0:
            result_df[month] = df[ending_balance_col]
        else:
            result_df.loc[:specified_index, month] = df.loc[:specified_index, ending_balance_col]
            result_df.loc[specified_index + 1:, month] = df.loc[specified_index + 1:, activity_col]
            
    return result_df

def unpivot_and_date(df):
    df_melted = df.melt(id_vars=['Account'], var_name='Date', value_name='Value')
    df_melted['Date'] = df_melted['Date'].apply(lambda x: parser.parse(x).strftime('%Y-%m'))
    df_melted['Year'] = df_melted['Date'].apply(lambda x: x.split('-')[0])
    df_melted['Month'] = df_melted['Date'].apply(lambda x: x.split('-')[1])
    df_melted = df_melted.drop(columns=['Date'])
    df_melted['Scenario'] = 'Actual'
    
    # Convert Year and Month to numeric
    df_melted['Year'] = pd.to_numeric(df_melted['Year'], errors='coerce')
    df_melted['Month'] = pd.to_numeric(df_melted['Month'], errors='coerce')
    
    # Filter out rows where Value is zero
    df_melted = df_melted[df_melted['Value'] != 0]
    
    df_melted = df_melted[['Account', 'Year', 'Month', 'Scenario', 'Value']]
    return df_melted

# Streamlit app
st.title("QuickBooks Trial Balance Formatter")
st.write("""
    Upload a QuickBooks Online Trial Balance Report (Excel file) with columns separated by month.
    Ensure the Excel file has been copied and pasted as values (not formulas) before uploading.
""")

uploaded_file = st.file_uploader("Upload your QuickBooks Trial Balance Excel file", type=["xlsx"])
last_bal_sheet_acc = st.text_input("Enter the last balance sheet account (must match exactly):", placeholder="e.g., Retained Earnings")

# Disable the process button until both file and text input are provided
process_button = st.button("Process File", disabled=not (uploaded_file and last_bal_sheet_acc))

if process_button:
    try:
        # Process the uploaded file
        tb = format_qb_tb(uploaded_file)
        tb1 = calculate_activity(tb, last_bal_sheet_acc)
        tb2 = unpivot_and_date(tb1)

        # Create a buffer for the output Excel file
        output = io.BytesIO()
        tb2.to_excel(output, index=False)
        output.seek(0)

        # Provide download link
        st.download_button(
            label="Download Processed File",
            data=output,
            file_name="Processed_Trial_Balance.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.success("File processed successfully!")
    except Exception as e:
        st.error(f"An error occurred: {str(e)}")
