import pandas as pd

def read_from_excel(file_path, sheet_name, start_row, start_col, end_row, end_col):
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    data = df.iloc[start_row-1:end_row, start_col-1:end_col].values
    return data

def write_to_excel(file_path, sheet_name, data, start_row, start_col):
    try:
        df = pd.DataFrame(data)
        with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            df.to_excel(writer, sheet_name=sheet_name, startrow=start_row-1, startcol=start_col-1, header=False, index=False)
    except PermissionError:
        print("Permission denied: Please make sure the file is not open in another application and you have write access.")
        raise
    except Exception as e:
        print(f"An error occurred: {e}")
        raise
