import pandas as pd
import os
import shutil


def tracking():

    # Define the path to the Excel file containing the subfolder names
    excel_file_path = r'C:\Users\Barkhatov Sergei\OneDrive - ' \
                      r'ТОО «ХЬЮМАН КАПИТАЛ ГРУП»\Документы\Work\ЭТ\Tracking\C500.xlsx'

    # Read the Excel file into a DataFrame
    df = pd.read_excel(excel_file_path)

    # Assuming the column containing subfolder names is named "Repair ID"
    subfolder_column = 'Repair ID'

    # Get the list of subfolder names from the DataFrame
    subfolder_names = df[subfolder_column].tolist()

    # Define the base directory where the subfolders are located
    base_directory = r'C:\Users\Barkhatov Sergei\OneDrive - ' \
                     r'ТОО «ХЬЮМАН КАПИТАЛ ГРУП»\Документы\Work\ЭТ\Tracking\Reports_19092023'

    # Define the path to the new folder where subfolders will be copied
    new_folder_path = r'C:\Users\Barkhatov Sergei\OneDrive - ' \
                      r'ТОО «ХЬЮМАН КАПИТАЛ ГРУП»\Документы\Work\ЭТ\Tracking\C500_photo'

    # Iterate through the subfolder names and check if they exist in the base directory
    missing = []
    for subfolder_name in subfolder_names:
        subfolder_path = os.path.join(base_directory, subfolder_name)
        new_subfolder_path = os.path.join(new_folder_path, subfolder_name)
        if os.path.exists(subfolder_path) and os.path.isdir(subfolder_path):
            if not os.path.exists(new_subfolder_path):
                shutil.copytree(subfolder_path, new_subfolder_path)
        else:
            missing.append(subfolder_name)

    # Using DataFrame.insert() to add a column
    df2 = pd.DataFrame(missing)

    writer = pd.ExcelWriter(r'C:\Users\Barkhatov Sergei\OneDrive - '
                            r'ТОО «ХЬЮМАН КАПИТАЛ ГРУП»\Документы\Work\ЭТ\Tracking\Missing_photo.xlsx',
                            engine='xlsxwriter')
    df2.to_excel(writer, sheet_name='Missing', index=False)
    writer._save()

    print("Copy process completed")


if __name__ == '__main__':
    tracking()
