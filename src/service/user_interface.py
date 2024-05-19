import os

class UserInterface:
    @staticmethod
    def display_files(files):
        if not files:
            print("找不到任何 .xlsx 檔案QAQ")
            return None
        
        print("以下為找到的 .xlsx 檔案：")
        for idx, file in enumerate(files, start=1):
            print(f"{idx}. {os.path.basename(file)}")

        return files

    @staticmethod
    def user_select_file(files):
        try:
            selection = int(input("請輸入文件編號（ex:1, 2 ...）: "))
            if 1 <= selection <= len(files):
                return files[selection - 1]
            else:
                print("超出範圍，重來")
        except ValueError:
            print("無效的數值 = =")
            
        return None
