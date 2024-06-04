import pandas as pd

# ==== ==== ====
# Function to process MAN report
# ==== ==== ====
def process_man_report(selected_sheet):
    # Read MAN delivery report
    MAN_Report = 'MAN Delivery Report 2024.xlsx'
    sheet_data = pd.read_excel(MAN_Report, sheet_name=selected_sheet)
    
    # Filter to display only wanted data (manually change iloc)
    filtered = (sheet_data.iloc[1:29, 4:15]
                .fillna(0)
                .rename(columns=sheet_data.iloc[1, 4:15])
                .iloc[1:]
                .reset_index(drop=True)
    )
    
    # New DF to display summation
    column_sum = filtered.sum(axis=0).rename('OK')
    filtered = pd.concat([filtered, pd.DataFrame(column_sum).T])

    # Create final grouped DF
    grouped_columns = filtered.columns.str[:13]
    grouped_sums = (filtered.T
                    .groupby(grouped_columns)
                    .sum()
                    .T
    )

    grouped_sum_df = grouped_sums.loc[['OK']].T

    # Read NOK data (manually change iloc)
    MAN_NOK = pd.read_excel('NOK tracking.xlsx')
    MAN_NOK = MAN_NOK.iloc[25:31][['NOK FULL BOXES', 'Unnamed: 9']].reset_index(drop=True)
    MAN_NOK.columns = ['P/N', 'NOK']
    MAN_NOK = MAN_NOK.reindex([5] + list(range(5))).reset_index(drop=True)

    # Finalize monthly report for MAN
    MAN_NOK.reset_index(drop=True, inplace=True)
    grouped_sum_df.reset_index(drop=True, inplace=True)
    MAN_NOK['OK'] = grouped_sum_df['OK']
    MAN_NOK['Controlled'] = MAN_NOK['NOK'] + MAN_NOK['OK']
    MAN_NOK['NOK rate(%)'] = (MAN_NOK['NOK'] / MAN_NOK['Controlled']) * 100
    MAN_NOK['NOK'], MAN_NOK['OK'] = MAN_NOK['OK'], MAN_NOK['NOK']
    MAN_NOK = MAN_NOK.rename(columns={'NOK': 'OK', 'OK': 'NOK'})

    total_row = pd.DataFrame(
        {'P/N': ['Total'], 
         'OK': '', 
         'NOK': '', 
         'Controlled': '', 
         'NOK rate(%)': MAN_NOK['NOK rate(%)'].mean()
        }, 
         index=[6]
    )

    MAN_NOK = pd.concat(
        [MAN_NOK.iloc[:6], 
         total_row, MAN_NOK.iloc[6:]
        ]
    ).reset_index(drop=True)
    
    # Export
    MAN_NOK.to_excel('MAN QC Report.xlsx', index=False)


# ==== ==== ====
# Function to process VOLVO report
# ==== ==== ====
def process_volvo_report(selected_sheet):
    # Read VOLVO delivery report
    VOLVO_Report = 'VOLVO Delivery Report 2024.xlsx'
    sheet_data = pd.read_excel(VOLVO_Report, sheet_name=selected_sheet)
    
    # Filter to display only wanted data (manually change iloc)
    filtered = (sheet_data.iloc[0:35, 4:16]
                .fillna(0)
    )

    column_sum = filtered.sum(axis=0).rename('OK')
    filtered = pd.concat([filtered, pd.DataFrame(column_sum).T])
    grouped_columns = filtered.columns.str[-5:-1]
    grouped_sums = (filtered.T
                    .groupby(grouped_columns)
                    .sum()
                    .T
    )

    grouped_sum_df = grouped_sums.loc[['OK']].T

    # Read NOK data (Manually change iloc)
    VOLVO_NOK = pd.read_excel('NOK tracking.xlsx')
    VOLVO_NOK = VOLVO_NOK.iloc[34:46][['NOK FULL BOXES', 'Unnamed: 9']].reset_index(drop=True)
    VOLVO_NOK.columns = ['P/N', 'NOK']

    # Finalize monthly report for VOLVO
    VOLVO_NOK.reset_index(drop=True, inplace=True)
    grouped_sum_df.reset_index(drop=True, inplace=True)
    VOLVO_NOK['OK'] = grouped_sum_df['OK']
    VOLVO_NOK['Controlled'] = VOLVO_NOK['NOK'] + VOLVO_NOK['OK']
    VOLVO_NOK['NOK rate(%)'] = (VOLVO_NOK['NOK'] / VOLVO_NOK['Controlled']) * 100
    VOLVO_NOK['NOK'], VOLVO_NOK['OK'] = VOLVO_NOK['OK'], VOLVO_NOK['NOK']
    VOLVO_NOK = VOLVO_NOK.rename(columns={'NOK': 'OK', 'OK': 'NOK'})

    total_row = pd.DataFrame(
        {'P/N': ['Total'], 
         'OK': '', 
         'NOK': '', 
         'Controlled': '', 
         'NOK rate(%)': VOLVO_NOK['NOK rate(%)'].mean()
        }, index=[12]
    )

    VOLVO_NOK = pd.concat(
        [VOLVO_NOK.iloc[:12], 
         total_row, 
         VOLVO_NOK.iloc[12:]
        ]
    ).reset_index(drop=True)

    # Export
    VOLVO_NOK.to_excel('VOLVO QC Report.xlsx', index=False)


# ==== ==== ====
# Get data from sheet of desired excel file
# ==== ==== ====
def get_sheet(file):
    available_sheets = file.sheet_names
    sheet = input("Enter sheet name: ")
    if sheet in available_sheets:
        return sheet
    else:
        print(f"Available sheets: {available_sheets}")


# ==== ==== ====
# Read delivery report sheet
# ==== ==== ====
delivery = pd.ExcelFile('VOLVO Delivery Report 2024.xlsx')
available_sheets = delivery.sheet_names

manufacturer = input("Enter manufacturer (MAN or VOLVO): ").upper()
print(f"Available sheets: {available_sheets}")
selected_sheet = get_sheet(delivery)
if selected_sheet:
    if manufacturer == 'MAN':
        process_man_report(selected_sheet)
    elif manufacturer == 'VOLVO':
        process_volvo_report(selected_sheet)
    else:
        print("Invalid manufacturer specified.")