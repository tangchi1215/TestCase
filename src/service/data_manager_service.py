import datetime
import os
import pandas as pd


class DataManagerService:
    @staticmethod
    def load_and_prepare_data(request):
        try:
            df = pd.read_excel(request.file_path, engine='openpyxl')
            selected_columns = df[['功能類別', '測試個案編號', '個案說明', '預期結果', '測試日期', '測試結果', '備註']]
        except KeyError as e:
            print(e, "標頭不符合指定格式")
            return None

        renamed_columns = selected_columns.rename(columns={
            '功能類別': '功能\n類別',
            '測試個案編號': '測試個案\n編號',
            '測試結果': '測試\n結果'
        })
        cleaned_data = renamed_columns.dropna(how='all')

        if request.test_date is None:
            request.test_date = datetime.datetime.now()

        if request.seq_no_option == "filename":
            base_name = os.path.basename(request.file_path).split('.')[0]
            cleaned_data['測試個案\n編號'] = [f"{base_name}-{i + 1:02}" for i in range(len(cleaned_data))]
        elif request.seq_no_option == "custom":
            cleaned_data['測試個案\n編號'] = [f"{request.custom_prefix}-{i + 1:02}" for i in range(len(cleaned_data))]
        elif request.seq_no_option == "excel":
            # 保持從Excel中讀取的測試編號
            pass

        cleaned_data['測試日期'] = request.test_date.strftime('%Y/%m/%d')
        cleaned_data['測試\n結果'] = '通過'
        return cleaned_data.fillna('')
