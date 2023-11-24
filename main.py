import pandas as pd
from datetime import datetime, timedelta
import numpy as np
import matplotlib.pyplot as plt


class DataInsert:
    def __init__(self, min_value, max_value, label):
        self.min_value = min_value
        self.max_value = max_value
        self.label = label


df_input = pd.read_csv("sp500_30m_1m.csv")
original_indices = range(0, df_input["TxDateTime"].size)
min_check_value = 45000
max_check_value = 60001
step_check_value = 100
df_output = pd.DataFrame({"Number check": np.arange(min_check_value, max_check_value, step_check_value)})
sub_arrays = []
sub_array_length = 5
letter_to_color = {
    'O': (215, 215, 215),
    'A': (215, 215, 215),
    'B': (195, 195, 195),
    'C': (175, 175, 175),
    'D': (155, 155, 155),
    'E': (135, 135, 135),
    'F': (115, 115, 115),
    'G': (95, 95, 95),
    'H': (75, 75, 75),
    'I': (55, 55, 55),
    'J': (35, 35, 35),
    'Z': (35, 35, 35),
    
    'N': (0, 255, 0) # Background
}

sheet_name = 'Sheet1' 
excel_file_path = 'input_fix.xlsx'
excel_file_path_output = 'output.xlsx'
excel_file_path_letter_to_color = 'letter_to_color.png'

def fix_tx_date_time():
    for idx in original_indices:
        date_str = df_input["TxDateTime"][idx]
        date_obj = datetime.strptime(date_str, '%m/%d/%Y %H:%M')
        start_time = date_obj - timedelta(minutes=29)
        output_str = f"{start_time.strftime('%m/%d/%Y %H:%M')} - {date_str}"
        df_input["TxDateTime"][idx] = output_str
        df_input.to_excel(excel_file_path)


def create_template():
    df_input_fix = pd.read_excel(excel_file_path)
    for i in range(0, len(original_indices), sub_array_length):
        sub_array = original_indices[i: i + sub_array_length]
        sub_arrays.append(sub_array)

    label_index = 0
    label_list = ["A", "B", "C", "D", "E"]
    for sub_array in sub_arrays:
        label_index = 0
        data_insert_list = []
        for i in sub_array:
            data_insert = DataInsert(
                df_input_fix["LowPrice"][i], df_input_fix["HighPrice"][i], label_list[label_index]
            )
            label_index += 1
            data_insert_list.append(data_insert)
        insert_data_frame(
            df_input_fix["TxDateTime"][sub_array[0]],
            df_input_fix["TxDateTime"][sub_array[-1]],
            data_insert_list,
        )

    df_output.to_excel(excel_file_path_output, index=False)


def insert_data_frame(start_tx_date_time, end_tx_date_time, data_insert_list):
    data = []
    for check_value in np.arange(min_check_value, max_check_value, step_check_value):
        label = ""
        for data_insert in data_insert_list:
            if data_insert.min_value / 1000 < check_value < data_insert.max_value / 1000:
                label += data_insert.label
        data.append(label)
    df_output[f"{start_tx_date_time} - {end_tx_date_time}"] = data
    
    
def change_color_based_on_mapping():
    excel_data = pd.read_excel(excel_file_path_output)
    excel_data.fillna('N', inplace=True)

    image_data = np.zeros((excel_data.shape[0], excel_data.shape[1], 3), dtype=np.uint8)
    for row in range(excel_data.shape[0]):
        for col in range(excel_data.shape[1]):
            cell_value = excel_data.iloc[row, col]
            color = letter_to_color.get(cell_value, (0, 0, 0))
            image_data[row, col] = color

    plt.figure(figsize=(10, 10))
    plt.imshow(image_data)
    plt.axis('off') 
    plt.tight_layout() 
    plt.savefig(excel_file_path_letter_to_color)
    plt.show()

# fix_tx_date_time()
# create_template()
change_color_based_on_mapping()