import datetime
import os
import pandas as pd


class DataManager:
    @staticmethod
    def load_and_prepare_data(file_path):
        try:
            df = pd.read_excel(file_path, engine='openpyxl')
            selected_columns = df[['功能類別', '測試個案編號', '個案說明', '預期結果', '測試日期', '測試結果', '備註']]
        except KeyError as e:
            print("標頭不符合指定格式")
            return None

        renamed_columns = selected_columns.rename(columns={
            '功能類別': '功能\n類別',
            '測試個案編號': '測試個案\n編號',
            '測試結果': '測試\n結果'
        })
        cleaned_data = renamed_columns.dropna(how='all')
        current_date = datetime.datetime.now().strftime('%Y/%m/%d')
        cleaned_data['測試日期'] = current_date
        cleaned_data['測試\n結果'] = '通過'
        base_name = os.path.basename(file_path.split('.')[0])
        cleaned_data['測試個案\n編號'] = [f"{base_name}-{i + 1:02}" for i in range(len(cleaned_data))]
        return cleaned_data.fillna('')
