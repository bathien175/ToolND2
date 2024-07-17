#cào hóa đơn và dịch vụ đã sử dụng
#using
import csv
import os
import time
import tkinter as tk
from tkinter import messagebox
from datetime import datetime, timedelta
import customtkinter
from tkinter import filedialog
from seleniumwire import webdriver
from selenium.webdriver.common.by import By
from PIL import ImageTk, Image
from tkcalendar import Calendar
from babel.numbers import *
import json
import pandas as pd
from tkinter import ttk
from openpyxl.cell import MergedCell
from openpyxl import Workbook, load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment, PatternFill
import threading
#global
customtkinter.set_appearance_mode("Light")
customtkinter.set_default_color_theme("green")
pathExcel = ""
drop = Any
#class
class PatientModel:
    stt = 0
    prescription_id = ""
    code = ""
    codeATC = ""
    o_name = ""
    p_name =""
    date = ""
    morning = ""
    noon = ""
    afternoon = ""
    evening = ""
    quantity_num = 0
    solan_ngay = 0
    num_per_time = ""
    unit_used = ""
    note = ""
    mapping = 0

    def to_dict(self):
        return {
            "STT": self.stt,
            "Mã toa thuốc": self.prescription_id,
            "Mã thuốc": self.code,
            "Mã atc": self.codeATC,
            "Tên thuốc gốc": self.o_name,
            "Tên độc quyền": self.p_name,
            "Ngày": self.date,
            "Sáng": self.morning,
            "Trưa": self.noon,
            "Chiều": self.afternoon,
            "Tối": self.evening,
            "Số lượng kê": self.quantity_num,
            "Số lần dùng / ngày": self.solan_ngay,
            "Lượng dùng mỗi lần": self.num_per_time,
            "Đơn vị dùng": self.unit_used,
            "Ghi chú bác sĩ": self.note,
            "IDMapping": self.mapping
        }

    def ExportModel(self):
        headers = [
            ("STT", self.stt),
            ("Mã toa thuốc", self.prescription_id),
            ("Mã thuốc", self.code),
            ("Mã atc", self.codeATC),
            ("Tên thuốc gốc", self.o_name),
            ("Tên độc quyền", self.p_name),
            ("Ngày", self.date),
            ("Sáng", self.morning),
            ("Trưa", self.noon),
            ("Chiều", self.afternoon),
            ("Tối", self.evening),
            ("Số lượng kê", self.quantity_num),
            ("Số lần dùng / ngày", self.solan_ngay),
            ("Lượng dùng mỗi lần", self.num_per_time),
            ("Đơn vị dùng", self.unit_used),
            ("Ghi chú bác sĩ", self.note),
            ("IDMapping", self.mapping)
        ]

        # Tính toán độ rộng cột
        max_key_length = max(len(header[0]) for header in headers) + 2
        max_value_length = max(len(str(header[1])) for header in headers) + 2

        s = ""
        s += "+" + "-" * max_key_length + "+" + "-" * max_value_length + "+\n"

        for header, value in headers:
            key_column = f" {header:<{max_key_length-2}} "
            value_column = f" {value:<{max_value_length-2}} "
            s += f"|{key_column}|{value_column}|\n"

        s += "+" + "-" * max_key_length + "+" + "-" * max_value_length + "+\n"
        return s
#function
def get_last_value_of_column(csv_filename, column_index):
    last_value = None
    with open(csv_filename, mode='r', encoding='utf-8-sig') as file:
        reader = csv.reader(file)
        for row in reader:
            if len(row) > column_index:
                last_value = row[column_index]
    return last_value

def get_unit(key):
    df = pd.read_csv("resource\enum_unit_usage.csv")
    # Tìm giá trị dựa trên khóa
    value = df.loc[df['key'] == key, 'value']
    if not value.empty:
        return value.iloc[0]
    else:
        return None

def check_valid_used(key):
    if key == "" or key == "0":
        return "0"
    else:
        return key

def start_script_thread(listmodels, directory):
    global script_thread  # Khai báo biến toàn cục
    terminal_window, terminal_text = open_terminal_window()
    app.withdraw()
    script_thread = threading.Thread(target=run_script, args=(listmodels, directory, terminal_text))
    script_thread.start()

def open_terminal_window():
    terminal_window = tk.Toplevel(app)
    terminal_window.title("Màn hình chạy dữ liệu")
    terminal_window.attributes("-fullscreen", True)
    icon_path = os.path.abspath("resource\crawlLogo.ico")
    terminal_window.iconbitmap(icon_path)
    terminal_text = tk.Text(terminal_window, bg="black", fg="green", insertbackground="green")
    terminal_text.pack(expand=True, fill='both')

    def on_closing():
        terminal_window.destroy()
        app.deiconify()

    terminal_window.protocol("WM_DELETE_WINDOW", on_closing)
    return terminal_window, terminal_text

def handle_file_excel():
    #khai báo biến

    #xét lỗi và cho chọn file excel
    valid = validate_input()
    if valid == False:
        messagebox.showerror(title="Lỗi dữ liệu", message="Chưa chọn file excel dữ liệu")
        return
    
    directory = filedialog.askdirectory(title="Chọn nơi lưu trữ!")
    if directory:
        directory = directory.replace('/', '\\')
    else:
        return
    
    existing_excel_files = [f for f in os.listdir(directory) if f.endswith('.xlsx') or f.endswith('.xls')]
    num_existing_files = len(existing_excel_files)

    listmodel = []
    # Đọc file excel
    if pathExcel.endswith('.xlsx'):
        excel_file = pd.ExcelFile(pathExcel, engine='openpyxl')
    elif pathExcel.endswith('.xls'):
        excel_file = pd.ExcelFile(pathExcel, engine='xlrd')
    else:
        raise ValueError("Unsupported file format")

    # Duyệt qua các sheet của file excel đầu vào và bỏ qua số sheet tương ứng với số file Excel đã có trong thư mục
    for idx, sheet_name in enumerate(excel_file.sheet_names):
        if idx < num_existing_files:
            continue  # Bỏ qua sheet này
        try:
            # Cố gắng chuyển đổi tên sheet thành định dạng ngày
            sheet_name_trip = sheet_name.strip()
            date = datetime.datetime.strptime(sheet_name_trip, "%d-%m-%Y").date()
        except ValueError:
            # Nếu không thể chuyển đổi, in ra thông báo lỗi và bỏ qua sheet này
            print(f"Tên sheet '{sheet_name}' không đúng định dạng ngày. Sheet này sẽ bị bỏ qua.")
            continue

        df = excel_file.parse(sheet_name)
        patient_codes = df.iloc[2:, 1]  # Đọc cột mã bệnh nhân từ dòng thứ 3
        tickets_id = df.iloc[2:, 6]  # Đọc cột trạng thái khám từ dòng thứ 3

        # Chỉ lấy những dòng là mã bệnh nhân
        valid_patient_codes = []
        valid_patient_tickets =[]
        for a, b in zip(patient_codes, tickets_id):
            code_str = str(a)  # Convert to string
            ticketstr = str(b)  # Convert to string
            if code_str[0].isdigit():
                if ticketstr != "nan":
                    valid_patient_codes.append(code_str)
                    valid_patient_tickets.append(ticketstr)                   
        
        # Nếu không có danh sách mã bệnh nhân hợp lệ thì next
        if not valid_patient_codes:
            continue

        # Tạo model
        model = {
            'date': date,
            'patient_codes': valid_patient_codes,
            'tickets_id': valid_patient_tickets
        }
        listmodel.append(model)

    start_script_thread(listmodel, directory)

def run_script(listmodels,directory,terminal_text):
    global pathExcel

    # Hàm để hiển thị thông tin lên terminal
    def log_terminal(message):
        terminal_text.insert(tk.END,message + "\n")
        terminal_text.see(tk.END)  # Scroll to the end
        terminal_text.update_idletasks()

    #duyệt danh sách model
    for model in listmodels:
        dateKham = model['date']
        dateKham = dateKham.strftime("%Y-%m-%d")
        patient_codes_list = model['patient_codes']
        ticket_list = model['tickets_id']
        patient_result_list = []
        #csv file
        csv_filename = os.path.join(directory, f"Invoice_Data_{dateKham}.csv")
        csv_filename_detail = os.path.join(directory, f"InvoiceDetail_Data_{dateKham}.csv")
        # CSV file header
        csv_header = ["STT","Mã toa thuốc" ,"Mã thuốc", "Mã atc", "Tên thuốc gốc", "Tên độc quyền", "Ngày", "Sáng", "Trưa", "Chiều", "Tối", "Số lượng kê", "Số lần dùng / ngày", "Lượng dùng mỗi lần", "Đơn vị dùng", "Ghi chú bác sĩ","IDMapping"]
        csv_header_detail = ["STT","Mã toa thuốc" ,"Mã thuốc", "Mã atc", "Tên thuốc gốc", "Tên độc quyền", "Ngày", "Sáng", "Trưa", "Chiều", "Tối", "Số lượng kê", "Số lần dùng / ngày", "Lượng dùng mỗi lần", "Đơn vị dùng", "Ghi chú bác sĩ","IDMapping"]
         # Kiểm tra nếu file CSV đã tồn tại và đếm số dòng
        start_stt = 0
        start_code = ""
        idmapcurrent = 0
        if os.path.exists(csv_filename):
            with open(csv_filename, mode='r', encoding='utf-8-sig') as file:
                reader = csv.reader(file)
                next(reader) #bỏ qua dòng 1
                for row in reader:
                    if len(row) > 1:
                        start_stt+=1
                        start_code = row[16]
        if start_code != "":
            idmapcurrent = int(start_code)
        demstt = start_stt
        breakSTT = 1
        for pt,pl in zip(patient_codes_list[idmapcurrent:],prescription_list[idmapcurrent:]):
            success = False
            errorDrug = 0
            errorLogWeb = 0 
            while not success:
                try:
                    if breakSTT == 1 and errorDrug == 0:
                        #chrome settings
                        chrome_options = webdriver.ChromeOptions()
                        chrome_options.add_experimental_option("prefs", {
                            "download.default_directory": directory,
                            "profile.default_content_setting_values.automatic_downloads": 1,
                            "download.prompt_for_download": False,
                            "profile.default_content_settings.popups": 0,
                            "safebrowsing.enabled": "false",
                            "safebrowsing.disable_download_protection": True
                        })
                        ascii_nd2 = """
........................Đang kết nối lại website..............................
                        """
                        log_terminal(ascii_nd2)
                        prefs = {"credentials_enable_service": False,
                            "profile.password_manager_enabled": False}
                        chrome_options.add_experimental_option("prefs", prefs)
                        chrome_options.add_argument("--headless")  # Chạy trình duyệt trong chế độ headless
                        chrome_options.add_argument("--disable-gpu")  # Tăng tốc độ trên các hệ điều hành không có GPU
                        chrome_options.add_argument("--window-size=1920x1080")  # Thiết lập kích thước cửa sổ mặc định
                        driver = webdriver.Chrome(options=chrome_options)
                        url = "http://192.168.0.77/dist/#!/login"
                        try:
                            driver.get(url)
                        except Exception as e:
                            errorLogWeb += 1
                            log_terminal(f"Lỗi khi truy cập URL {url}: {e}")
                            if errorLogWeb >= 10:
                                log_terminal("Lỗi server, dừng chương trình.")
                                return
                            time.sleep(3)  # Chờ một chút trước khi thử lại
                            continue

                        errorLogWeb = 0  # Reset biến đếm lỗi khi truy cập URL thành công
                        time.sleep(3)
                        log_terminal(".........................Duyệt WEBSITE thành công.................................")
                        # login
                        form = driver.find_element(By.NAME, "LoginForm")
                        username = driver.find_element(By.NAME, 'username')
                        password = driver.find_element(By.NAME, 'password')
                        username.send_keys('quyen.ngoq')
                        password.send_keys('74777477')
                        # ấn nút login
                        form.submit()
                        time.sleep(2)
                        # Chọn
                        liElement = driver.find_element(By.XPATH,
                                                        "//li[contains(@class, 'nav-item pl-2 pr-3 d-lg-none d-xl-block') and contains(@permission, "
                                                        "'patient_info_payment')]")
                        liElement.click()
                        time.sleep(1)
                    try:
                        passForm = driver.find_element(By.XPATH, "//button[contains(text(), ' Bỏ qua')]")
                        passForm.click()
                        time.sleep(0.1)
                    except:
                        print("Frist access")
                        time.sleep(0.1)
                    cod = str(pt)
                    patientcode = driver.find_element(By.NAME, 'patient_code_qr')
                    patientcode.clear()
                    patientcode.send_keys(cod)
                    time.sleep(0.1)
                    btnFind = driver.find_element(By.XPATH, "//button[contains(text(), 'Tìm')]")
                    btnFind.click()
                    time.sleep(0.1)
                    btnThuoc = driver.find_element(By.XPATH, "//a[contains(text(), 'Thuốc Ngoại Trú ')]")
                    btnThuoc.click()
                    time.sleep(0.1)
                    btnPrescription = driver.find_element(By.XPATH, f"//a[contains(text(), '{str(pl)}')]")
                    btnPrescription.click()
                    time.sleep(0.5)
                    #tìm mã toa phù hợp
                    for request in reversed(driver.requests):
                        foundDrug = False
                        emptyDrug = False
                        if request.response:
                            if 'load_pkg_last_detail_presc' in request.url:
                                # format cho dữ liệu về dạng json
                                try:
                                    drug_data = json.loads(request.response.body)
                                    # đọc data
                                    drug_info = drug_data.get('data', {})
                                    if len(drug_info) > 0:
                                        if drug_info[0].get('prescription_id') == int(pl):
                                            foundDrug = True  
                                            idmapcurrent += 1
                                            for di in drug_info:
                                                patient_result = PatientModel()
                                                patient_result.prescription_id = int(pl)
                                                patient_result.code = di.get('code')
                                                if di.get('code_atc') == "0":
                                                    patient_result.codeATC = ""
                                                else:
                                                    patient_result.codeATC = di.get('code_atc')
                                                patient_result.o_name= di.get('original_names')
                                                patient_result.p_name= di.get('proprietary_name')
                                                patient_result.morning = check_valid_used(di.get('morning'))
                                                patient_result.noon = check_valid_used(di.get('noon'))
                                                patient_result.afternoon = check_valid_used(di.get('afternoon'))
                                                patient_result.evening = check_valid_used(di.get('evening'))
                                                patient_result.date = dateKham
                                                patient_result.quantity_num = int(di.get('quantity_num'))
                                                patient_result.solan_ngay = di.get('solan_ngay')
                                                patient_result.num_per_time = di.get('num_per_time')
                                                unit = di.get('enum_unit_import_sell')
                                                patient_result.unit_used = get_unit(unit)
                                                if di.get('note') != None:
                                                    patient_result.note = di.get('note')
                                                else:
                                                    patient_result.note = ""
                                                demstt+=1
                                                patient_result.stt = demstt
                                                patient_result.mapping = idmapcurrent
                                                patient_result_list.append(patient_result)
                                                log_terminal(patient_result.ExportModel())  
                                            break
                                    else:
                                        idmapcurrent += 1
                                        emptyDrug = True
                                        break         
                                except Exception as e:
                                    print(e)
                                    continue    
                            else:
                                continue
                        
                        #Xét trường hợp rỗng
                        if emptyDrug == True:
                            errorDrug = 0
                            break
                        else:
                            #xét trường hợp tìm thấy hay không
                            if foundDrug == True:
                                errorDrug = 0
                                break
                            else:
                                errorDrug +=1
                                break
                    if emptyDrug == True:
                        success = True
                        log_terminal(f"Toa thuốc không kê thuốc!...........")
                    else:
                        if errorDrug > 0:
                            log_terminal(f"Dữ liệu toa thuốc trống, đang cố thử lại...")
                            if errorDrug >= 10:
                                log_terminal(f"Quá nhiều lỗi dữ liệu trống, bỏ qua kết thúc quá trình cào tại mã bệnh nhân {cod}")
                                app.destroy()
                            continue  # Thử lại mã bệnh nhân hiện tại
                        else:
                            success = True
                            if breakSTT % 50 == 0:
                                # Ghi dữ liệu tạm thời vào file CSV
                                with open(csv_filename, mode='a', newline='', encoding='utf-8') as file:
                                    writer = csv.DictWriter(file, fieldnames=csv_header)
                                    if file.tell() == 0:  # Nếu file rỗng, ghi header
                                        writer.writeheader()
                                    writer.writerows([pt.to_dict() for pt in patient_result_list])
                                breakSTT = 1
                                driver.quit()
                                patient_result_list.clear()
                                log_terminal(".........................Ghi tạm vào file CSV........................................")
                            else:
                                breakSTT = breakSTT + 1
                except:
                    success = False
        # Ghi dữ liệu còn lại vào file CSV nếu có
        if patient_result_list:
            with open(csv_filename, mode='a', newline='', encoding='utf-8') as file:
                writer = csv.DictWriter(file, fieldnames=csv_header)
                if file.tell() == 0:
                    writer.writeheader()
                writer.writerows([pt.to_dict() for pt in patient_result_list])
        # Thoát trình duyệt 
        driver.quit()
        # Chuyển dữ liệu từ file CSV sang file Excel
        try:
            df = pd.read_csv(csv_filename, encoding='utf-8-sig')
            output_filename = os.path.join(directory, f"Prescription_Data_{dateKham}.xlsx")
            with pd.ExcelWriter(output_filename, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name='Sheet1')

            # Định dạng file Excel
            wb = load_workbook(output_filename)
            ws = wb.active

            # Định dạng header
            header_font = Font(bold=True)
            for cell in ws[1]:
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")

            # Tự động điều chỉnh độ rộng cột
            for col in ws.columns:
                max_length = 0
                column = col[0].column_letter  # Get the column name
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(cell.value)
                    except:
                        pass
                adjusted_width = (max_length + 2)
                ws.column_dimensions[column].width = adjusted_width
            
            # Định dạng cột Mã bệnh nhân là Text
            for row in ws.iter_rows(min_row=2, min_col=2, max_col=2):
                for cell in row:
                    cell.number_format = '@'
            
            wb.save(output_filename)
            os.remove(csv_filename)  # Xóa file CSV sau khi chuyển sang Excel
            log_terminal(f"Quá trình thu thập dữ liệu toa thuốc ngày {dateKham} thành công!")
        except Exception as e:
            messagebox.showerror(title="Lỗi", message=f"Lỗi quá trình tạo file excel không thành công! {e}")  
    
    # Đóng cửa sổ terminal và hiển thị lại cửa sổ chính
    app.deiconify()

def validate_input():
    if pathExcel == "":
        return False
    return True

def cancel_Excel_file():
    global pathExcel, drop
    if pathExcel == '':
        messagebox.showerror(title="Thất bại!", message="Chưa chọn file nào!")
    else:
        pathExcel = ''
        drop.configure(text="Ấn để chọn file Excel!")
        # Cập nhật fg_color của nút
        drop.configure(fg_color="#FF9292")
        # Bật enable nút
        drop.configure(state="enable")

def select_Excel_file():
    global pathExcel, drop
    # Hiển thị hộp thoại chọn file PDF
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        # Lấy đường dẫn file
        pathExcel = file_path
        # Lấy tên file từ đường dẫn
        file_name = os.path.basename(file_path)
        # Cập nhật text của nút với tên file
        drop.configure(text=file_name)
        # Cập nhật fg_color của nút
        drop.configure(fg_color="#5DBD60")
        # Tắt enable nút
        drop.configure(state="disabled")

def center_window(window, width=600, height=440):
    # Lấy kích thước màn hình
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    # Tính toán tọa độ x và y để cửa sổ ở giữa màn hình
    x = (screen_width / 2) - (width / 2)
    y = (screen_height / 2) - (height / 2)
    window.geometry(f'{width}x{height}+{int(x)}+{int(y)}')
#app
def run_secondary_interface(main_app):
    global run_button, app, drop
    app = customtkinter.CTkToplevel(main_app)
    app.title("Lấy dữ liệu hóa đơn")
    center_window(app)
    icon_path = os.path.abspath("resource\crawlLogo.ico")
    app.iconbitmap(icon_path)

    def on_closing():
        main_app.deiconify()  # Hiển thị lại cửa sổ chính khi đóng cửa sổ mới
        app.destroy()

    app.protocol("WM_DELETE_WINDOW", on_closing)

    imgBG = ImageTk.PhotoImage(Image.open("resource\\backgroundInvoice.jpg"))
    l1 = customtkinter.CTkLabel(master=app, image=imgBG)
    l1.pack()

    frame = customtkinter.CTkFrame(master=l1, width=320, height=300, corner_radius=15)
    frame.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

    l2 = customtkinter.CTkLabel(master=frame, text="Select configuration", font=('Century Gothic', 20))
    l2.place(x=70, y=45)

    drop = customtkinter.CTkButton(master=frame, command=select_Excel_file, width=300, height=50, corner_radius=15, fg_color="#FF9292", hover_color="#FFB6B6", text_color="#000000", text="Ấn để chọn file Excel!", font=('Tahoma', 13))
    drop.place(x=10, y=100)

    cancel_button = customtkinter.CTkButton(master=frame, text="Hủy bỏ", command=cancel_Excel_file, font=('Tahoma', 13), fg_color="#B3271C", hover_color="#FF3D4D")
    cancel_button.place(x=10, y=200)

    run_button = customtkinter.CTkButton(master=frame, text="Thực thi", command=handle_file_excel, font=('Tahoma', 13), fg_color="#005369", hover_color="#008097")
    run_button.place(x=170, y=200)

    app.mainloop()