import os
import pandas as pd
from datetime import datetime
import shutil
import openpyxl

# Define the paths
client_folder = r"E:\Python Task\Client Folder"
dev_folder = r"E:\Python Task\Dev Team Folder"
excel_file = r"E:\Python Task\folder_sync.xlsx"

def take_snapshot(folder_path):
    """Take a snapshot of files in the specified folder, capturing their names and last modified times."""
    snapshot = []
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            file_path = os.path.join(root, file)
            last_modified = datetime.fromtimestamp(os.path.getmtime(file_path))
            snapshot.append({'File': file, 'Last Modified': last_modified})
    return pd.DataFrame(snapshot)

def save_snapshot(merged_df, file_path):
    """Save the merged DataFrame to the Excel file, replacing the 'Last Snapshot' sheet if it exists."""
    # Load the workbook and remove the sheet if it exists
    workbook = openpyxl.load_workbook(file_path)
    if 'Last Snapshot' in workbook.sheetnames:
        del workbook['Last Snapshot']
    workbook.save(file_path)
    
    # Now write the new data
    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
        merged_df.to_excel(writer, sheet_name='Last Snapshot', index=False)

def compare_snapshots(client_df, dev_df):
    """Compare snapshots of client and dev folders, determining which files need to be synchronized."""
    merged_df = pd.merge(client_df, dev_df, on='File', how='outer', suffixes=('_Client', '_Dev'))
    merged_df['Sync Action'] = merged_df.apply(determine_sync_action, axis=1)
    merged_df = merged_df.sort_values(by=['Last Modified_Client', 'Last Modified_Dev'], ascending=False)
    return merged_df

def determine_sync_action(row):
    """Determine the synchronization action based on the modification times."""
    if pd.isna(row['Last Modified_Client']):
        return 'Copy to Client'
    elif pd.isna(row['Last Modified_Dev']):
        return 'Copy to Dev'
    elif row['Last Modified_Client'] > row['Last Modified_Dev']:
        return 'Copy to Dev'
    elif row['Last Modified_Client'] < row['Last Modified_Dev']:
        return 'Copy to Client'
    else:
        return 'In Sync'

def sync_files(merged_df):
    """Sync files based on the determined sync actions."""
    for index, row in merged_df.iterrows():
        if row['Sync Action'] == 'Copy to Client':
            src_path = os.path.join(dev_folder, row['File'])
            dest_path = os.path.join(client_folder, row['File'])
            shutil.copy2(src_path, dest_path)
            print(f"Copied {row['File']} to Client Folder.")
        elif row['Sync Action'] == 'Copy to Dev':
            src_path = os.path.join(client_folder, row['File'])
            dest_path = os.path.join(dev_folder, row['File'])
            shutil.copy2(src_path, dest_path)
            print(f"Copied {row['File']} to Dev Folder.")

if __name__ == "__main__":
    # Take snapshots of the client and dev folders
    client_snapshot = take_snapshot(client_folder)
    dev_snapshot = take_snapshot(dev_folder)
    
    # Compare the snapshots
    comparison_result = compare_snapshots(client_snapshot, dev_snapshot)
    
    # Save the result to the Excel file
    save_snapshot(comparison_result, excel_file)
    
    # Sync the files based on the comparison
    sync_files(comparison_result)

    print("Folder synchronization completed successfully.")
