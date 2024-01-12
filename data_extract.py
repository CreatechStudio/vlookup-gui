import pandas as pd
import numpy as np
import datetime
import os

# Enter file name
# file_name = input("Please export xlsm to xlsx first, put it into the same folder as this program, and enter the file name (without the suffix):")
# input_file_path = f'{file_name}.xlsx'
def data_extract_function(input_file_path):
    print(f'Extracting {input_file_path}...')

    # Extract file
    df = pd.read_excel(input_file_path)

    # Set pandas option to avoid scientific notation display
    pd.set_option('display.float_format', lambda x: f'{int(x)}' if np.isfinite(x) else '')

    # Define the rule validation function
    def validateData(row):
        invalid_rows = {}

        return invalid_rows

    # Store the line number of the line that does not meet the rules and the header of the problem
    invalid_rows = {}

    # Traverse each line and check the rules
    for index, row in df.iterrows():
        row = row.apply(lambda x: x.strip() if isinstance(x, str) else x)
        row_invalid = validateData(row)
        if row_invalid:
            for key, value in row_invalid.items():
                if value - 1 in invalid_rows:
                    invalid_rows[value - 1].append(key)
                else:
                    invalid_rows[value - 1] = [key]
        df.loc[index] = row

    # Extract header
    headers = df.columns

    # Create an empty dictionary to store data
    data_dict = {}

    # Loop through each row/column
    for index, row in df.iterrows():
        for header in headers:
            value = row[header]
            replaced_value = None
            # Skip empty values
            if pd.isna(value):
                value = None  # 将缺失值设为None
            if isinstance(value, float):
                value = int(value)
            # Store in dictionary
            if header not in data_dict:
                data_dict[header] = []
            data_dict[header].append(value if value is not None else [])  # 添加空列表，而不是None

    # Check if any column has all empty lists and convert to empty list
    for header in headers:
        if all(isinstance(value, list) and not value for value in data_dict[header]):
            data_dict[header] = []

    # Store the line number of the line that does not meet the rules
    invalid_row_number_only = list(invalid_rows.keys())


    if invalid_rows:
        print("Invalid rows:")
        for row, headers in invalid_rows.items():
            row += 1
            print(f"Row {row}: {', '.join(headers)}")
        invalid_row_number_only = list(invalid_rows.keys())

        folder_path = "error_log"
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)
            print(f"Folder {folder_path} create Successfully!")
        else:
            print(f"Folder {folder_path} already exists!")
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        filename = f"error_log/error_{timestamp}.txt"
        with open(filename, "w") as file:
            file.write("Invalid rows:\n")
            for row, headers in invalid_rows.items():
                row += 1
                file.write(f"Invalid rows {row}: {', '.join(headers)}\n")

        print(f"Invalid data has beed stored in {filename}")
        valid_flag = False
    else:
        print("Congratulations! No errors found! Start running the automation program!")
        valid_flag = True

    print(data_dict)

    # Publish variables
    return data_dict, valid_flag

# data_extract_function("/Users/jessiezhu/Documents/GitHub/vlookup-gui/target_result.xlsx")

    # 打印全部原始数据
    # for header, values in data_dict.items():
    #     print(f"{header}: {values}")
    # 打印pub_data_dict
    # print(pub_data_dict)
    # print(pub_invalid_rows)
    # 测试输出特定数据
    # print(pub_data_dict['Year_Group'][0])
    # print(pub_data_dict['1st_Surname'][0])
    # print(pub_data_dict['StuEmail'][0])
    # print(pub_replaced_data_dict['txtForename'][0])
    # print(pub_replaced_data_dict['Year_Group'][0])
    # print(pub_replaced_data_dict['txtForm'][0])
    # print(pub_replaced_data_dict['intForm'][0])
    # print(pub_replaced_data_dict['Day'][0])
    # print(pub_replaced_data_dict['txtStuEmail'][0])
    # print(pub_data_dict['SchoolID'][0])
    # for i in range(len(pub_data_dict['SchoolID'])):
    #             row_data = {header: values[i] for header, values in pub_data_dict.items()}
    #             replaced_data = {header: values[i] for header, values in pub_replaced_data_dict.items()}
    #             print(row_data)
    # row_data = {header: values[2] for header, values in pub_data_dict.items()}
    # print(row_data)