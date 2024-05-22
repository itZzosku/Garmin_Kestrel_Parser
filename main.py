import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os


def find_closest(row, df):
    """
    Find the closest timestamp in kestrel_df to the timestamp in the current row of garmin_df.
    """
    try:
        row_time = pd.to_datetime(row['Timestamp'], format='%H:%M:%S')
        df['Timestamp'] = pd.to_datetime(df['Timestamp'], format='%H:%M:%S')
        time_diff = (df['Timestamp'] - row_time).abs()
        closest_index = time_diff.idxmin()
        return df.loc[closest_index]
    except Exception as e:
        print(f"Error in find_closest: {e}")
        raise


def process_garmin_sheet(sheet_name, garmin_df, kestrel_df):
    """
    Process each Garmin sheet by renaming columns, formatting timestamps, and finding closest matches in Kestrel data.
    """
    try:
        print(f"Processing Garmin sheet: {sheet_name}")
        garmin_df.rename(columns={garmin_df.columns[5]: 'Timestamp'}, inplace=True)
        garmin_df['Timestamp'] = pd.to_datetime(garmin_df['Timestamp'], format='%H:%M:%S', errors='coerce').dt.strftime('%H:%M:%S')
        garmin_df.dropna(subset=['Timestamp'], inplace=True)
        closest_rows = garmin_df.apply(lambda row: find_closest(row, kestrel_df), axis=1)
        combined_df = pd.concat([garmin_df.reset_index(drop=True), closest_rows.reset_index(drop=True)], axis=1)
        return combined_df
    except Exception as e:
        print(f"Error in process_garmin_sheet: {e}")
        raise


def read_file(file_path, is_kestrel):
    """
    Read the Kestrel or Garmin file based on the file extension and is_kestrel flag.
    """
    try:
        if file_path.endswith('.xlsx'):
            if is_kestrel:
                df = pd.read_excel(file_path, skiprows=5, usecols=[0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14])
                df.rename(columns={df.columns[0]: 'Timestamp', df.columns[1]: 'Temperature',
                                   df.columns[2]: 'Relative Humidity', df.columns[3]: 'Station Pressure'}, inplace=True)
                df = df[['Timestamp', 'Temperature', 'Relative Humidity', 'Station Pressure'] + df.columns[4:].tolist()]
                df['Timestamp'] = pd.to_datetime(df['Timestamp'], format='%Y-%m-%d %I:%M:%S %p', errors='coerce').dt.strftime('%H:%M:%S')
            else:
                xl = pd.ExcelFile(file_path)
                df_dict = {sheet: xl.parse(sheet, skiprows=1, usecols=[0, 1, 2, 3, 4, 5, 6, 7, 8]) for sheet in xl.sheet_names}
                for sheet, df in df_dict.items():
                    df.rename(columns={df.columns[5]: 'Timestamp'}, inplace=True)
                    df['Timestamp'] = pd.to_datetime(df['Timestamp'], format='%H:%M:%S', errors='coerce').dt.strftime('%H:%M:%S')
                return df_dict
        elif file_path.endswith('.csv'):
            if is_kestrel:
                df = pd.read_csv(file_path, skiprows=5)
                df.columns = [
                    'Timestamp', 'Temperature', 'Relative Humidity', 'Station Pressure',
                    'Heat Index', 'Dew Point', 'Density Altitude', 'Data Type',
                    'Record name', 'Start time', 'Duration (H:M:S)', 'Location description',
                    'Location address', 'Location coordinates', 'Notes'
                ]
                df['Timestamp'] = pd.to_datetime(df['Timestamp'], format='%Y-%m-%d %I:%M:%S %p', errors='coerce').dt.strftime('%H:%M:%S')
                df = df[['Timestamp', 'Temperature', 'Relative Humidity', 'Station Pressure'] + df.columns[4:].tolist()]
            else:
                df = pd.read_csv(file_path, skiprows=1, usecols=[0, 1, 2, 3, 4, 5, 6, 7, 8])
                df['Timestamp'] = pd.to_datetime(df['Timestamp'], format='%H:%M:%S', errors='coerce').dt.strftime('%H:%M:%S')
        else:
            raise ValueError("Unsupported file format. Please provide an Excel or CSV file.")
        return df
    except Exception as e:
        print(f"Error in read_file: {e}")
        raise


def generate_unique_filename(filepath):
    """
    Generate a unique filename by appending a counter if the file already exists.
    """
    base, extension = os.path.splitext(filepath)
    counter = 1
    new_filepath = filepath
    while os.path.exists(new_filepath):
        new_filepath = f"{base}_{counter}{extension}"
        counter += 1
    return new_filepath


def process_files(kestrel_path, garmin_path):
    """
    Process the selected Kestrel and Garmin files, combine the data, and save it to a new file.
    """
    try:
        print(f"Reading Kestrel file: {kestrel_path}")
        kestrel_df = read_file(kestrel_path, is_kestrel=True)
        print(f"Kestrel DataFrame:\n{kestrel_df.head()}")

        print(f"Reading Garmin file: {garmin_path}")
        garmin_sheets = read_file(garmin_path, is_kestrel=False)

        all_combined_dfs = []
        if isinstance(garmin_sheets, dict):
            for sheet_name, garmin_df in garmin_sheets.items():
                combined_df = process_garmin_sheet(sheet_name, garmin_df, kestrel_df)
                all_combined_dfs.append((sheet_name, combined_df))
        else:
            combined_df = process_garmin_sheet("Sheet1", garmin_sheets, kestrel_df)
            all_combined_dfs.append(("Sheet1", combined_df))

        output_file = generate_unique_filename(os.path.join(os.path.dirname(kestrel_path), 'Combined_Output.xlsx'))
        with pd.ExcelWriter(output_file) as writer:
            for sheet_name, combined_df in all_combined_dfs:
                headers = ['Shot Count', 'Speed (MPS)', 'Δ AVG (MPS)', 'KE (J)', 'Power Factor (N⋅s)', 'Time',
                           'Clean Bore', 'Cold Bore', 'Shot Notes', 'Timestamp', 'Temperature', 'Relative Humidity',
                           'Station Pressure', 'Heat Index', 'Dew Point', 'Density Altitude', 'Data Type', 'Record name',
                           'Start time', 'Duration (H:M:S)', 'Location description', 'Location address', 'Location coordinates', 'Notes']
                combined_df.columns = headers[:combined_df.shape[1]]  # Ensure only the number of columns present
                combined_df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=5)

        messagebox.showinfo("Success", f"The combined data has been saved to {output_file}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")
        print(f"Error in process_files: {e}")


def select_kestrel_file():
    """
    Open file dialog to select the Kestrel file.
    """
    try:
        kestrel_path = filedialog.askopenfilename(title="Select Kestrel File", filetypes=[("Excel and CSV files", "*.xlsx;*.csv")])
        if kestrel_path:
            kestrel_entry.delete(0, tk.END)
            kestrel_entry.insert(0, kestrel_path)
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while selecting the Kestrel file: {e}")
        print(f"Error in select_kestrel_file: {e}")


def select_garmin_file():
    """
    Open file dialog to select the Garmin file.
    """
    try:
        garmin_path = filedialog.askopenfilename(title="Select Garmin File", filetypes=[("Excel and CSV files", "*.xlsx;*.csv")])
        if garmin_path:
            garmin_entry.delete(0, tk.END)
            garmin_entry.insert(0, garmin_path)
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while selecting the Garmin file: {e}")
        print(f"Error in select_garmin_file: {e}")


def combine_files():
    """
    Validate file paths and formats, then initiate file processing.
    """
    kestrel_path = kestrel_entry.get()
    garmin_path = garmin_entry.get()
    if not kestrel_path or not garmin_path:
        messagebox.showwarning("Warning", "Please select both files")
        return

    if not (kestrel_path.endswith(('.xlsx', '.csv')) and garmin_path.endswith(('.xlsx', '.csv'))):
        messagebox.showwarning("Warning", "Selected files must be Excel or CSV files")
        return

    process_files(kestrel_path, garmin_path)


if __name__ == "__main__":
    root = tk.Tk()
    root.title("Excel and CSV Combiner")
    root.geometry("500x300")

    lbl = tk.Label(root, text="Select Kestrel and Garmin files to combine them", pady=10)
    lbl.pack()

    kestrel_frame = tk.Frame(root)
    kestrel_frame.pack(pady=5)
    kestrel_label = tk.Label(kestrel_frame, text="Kestrel File:")
    kestrel_label.pack(side=tk.LEFT)
    kestrel_entry = tk.Entry(kestrel_frame, width=50)
    kestrel_entry.pack(side=tk.LEFT, padx=5)
    kestrel_button = tk.Button(kestrel_frame, text="Browse", command=select_kestrel_file)
    kestrel_button.pack(side=tk.LEFT)

    garmin_frame = tk.Frame(root)
    garmin_frame.pack(pady=5)
    garmin_label = tk.Label(garmin_frame, text="Garmin File:")
    garmin_label.pack(side=tk.LEFT)
    garmin_entry = tk.Entry(garmin_frame, width=50)
    garmin_entry.pack(side=tk.LEFT, padx=5)
    garmin_button = tk.Button(garmin_frame, text="Browse", command=select_garmin_file)
    garmin_button.pack(side=tk.LEFT)

    combine_button = tk.Button(root, text="Combine Files", command=combine_files, pady=10, padx=20)
    combine_button.pack(pady=20)

    root.mainloop()
