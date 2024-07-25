#using
import os
import time
import tkinter as tk
import customtkinter
from PIL import ImageTk, Image
from babel.numbers import *
from tkinter import messagebox
import json
import threading
import sourceString as sour
import sourceModel as model
import requests
import math
import phpserialize
import psycopg2
#global
customtkinter.set_appearance_mode("Light")
customtkinter.set_default_color_theme("green")
txt_search = Any
run_button = Any
app = Any
urlAPI = "http://192.168.0.77/api/patients/find"
terminal_window = Any

def validate_textSearch():
    global txt_search
    if txt_search.get() != "":
        return True
    else:
        return False

def open_terminal_window():
    terminal_window = tk.Toplevel(app)
    terminal_window.title("Màn hình chạy dữ liệu")
    terminal_window.attributes("-fullscreen", True)
    icon_path = os.path.abspath("resource\crawlLogo.ico")
    terminal_window.iconbitmap(icon_path)
    terminal_text = tk.Text(terminal_window, bg="black", fg="green", insertbackground="green")
    terminal_text.pack(expand=True, fill='both')

    def on_closing():
        if script_thread.is_alive():
            terminal_window.destroy()
            app.deiconify()
        else:
            terminal_window.destroy()
            app.deiconify()

    terminal_window.protocol("WM_DELETE_WINDOW", on_closing)
    return terminal_window, terminal_text

def start_script_thread():
    global script_thread  # Khai báo biến toàn cục
    if validate_textSearch() == True:
        terminal_window, terminal_text = open_terminal_window()
        app.withdraw()
        script_thread = threading.Thread(target=run_Script, args=(terminal_text,))
        script_thread.start()
    else:
        messagebox.showerror(title="Lỗi thông tin", message="Vui lòng không bỏ trống dữ liệu tìm kiếm!")

def center_window(window, width=600, height=440):
    # Lấy kích thước màn hình
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    # Tính toán tọa độ x và y để cửa sổ ở giữa màn hình
    x = (screen_width / 2) - (width / 2)
    y = (screen_height / 2) - (height / 2)
    window.geometry(f'{width}x{height}+{int(x)}+{int(y)}')

def fetch_all_pages(base_url, headers, payload):
    all_data = []
    current_page = 1
    total_pages = None

    while total_pages is None or current_page <= total_pages:
        # Cập nhật tham số page cho mỗi yêu cầu
        payloads = payload.copy()
        payloads['page'] = current_page

        response = requests.get(base_url, headers=headers, json=payloads)

        if response.status_code == 200:
            dataf = response.json()
            data = dataf['data'] #đây là dữ liệu tổng gồm có list dữ liệu và thông tin paging
            # Thêm dữ liệu từ trang hiện tại vào kết quả
            all_data.extend(data['data_list']) # chỗ này là lấy dữ liệu ra
            
            # Cập nhật thông tin phân trang
            paging = data['paging'] # lấy thông tin page
            current_page = paging['current_page']
            row_per_page = paging['row_per_page']
            total_record = paging['total_record']
            
            # Tính tổng số trang nếu chưa có
            if total_pages is None:
                total_pages = math.ceil(total_record / row_per_page)

            print(f"Đã lấy trang {current_page}/{total_pages}")
            
            # Tăng số trang cho lần lặp tiếp theo
            current_page += 1
        else:
            print(f"Lỗi khi lấy trang {current_page}: {response.status_code}")
            break

    return all_data

def run_Script(terminal_text):
    global app, txt_search, terminal_window
    # Hàm để hiển thị thông tin lên terminal
    def log_terminal(message):
        terminal_text.insert(tk.END,message + "\n")
        terminal_text.see(tk.END)  # Scroll to the end
        terminal_text.update_idletasks()
    log_terminal(".........................Khởi động chương trình..........................................")
    log_terminal(".........................Khởi động Chrome................................................")
    #chrome settings
    bTry = False
    errorChrome = 1
    while bTry == False:
        bTry = sour._initSelenium_()
        if bTry == False:
            if errorChrome >= 10:
                log_terminal("..............Khởi động Chrome thất bại quá nhiều! Tắt chương trình....................")
                app.destroy() 
            else:
                errorChrome += 1
                log_terminal(".........................Khởi động Chrome thất bại! Đang thử lại.......................")
        
    time.sleep(3)
    log_terminal(".........................Duyệt WEBSITE thành công........................................")
    # login
    sour._login_("quyen.ngoq", "74777477")
    # ấn nút login
    time.sleep(0.5)
    log_terminal(".........................Sao chép userkey thành công.....................................")
    headers = {
                    "Appkey": sour.Appkey,
                    "Userkey": sour.secretKey,
                    "Authorization": sour.secretKey,
                    "Content-Type": sour.contentType
                }
    payload = {
        "page": 1,
        "patient_code": txt_search.get(),
    }
    all_data_fetch = fetch_all_pages(urlAPI, headers, payload)
    conn_params = sour.ConnectStr
    conn = psycopg2.connect(**conn_params)
    cur = conn.cursor()
    for item in all_data_fetch:
        mod = model.PatientModel()
        mod.person_id = item.get('person_id')
        mod.patient_code = item.get('patient_code')
        mod.patient_code_2 = item.get('patient_code2')
        mod.vaccination_code = item.get('vaccination_code')
        mod.name = item.get('name')
        mod.backup_name = item.get('backup_name')
        mod.gender = item.get('gender')
        mod.date_of_birth = item.get('date_of_birth')
        mod.phone_number = item.get('phone_number')
        mod.full_address = item.get('full_address')
        mod.career_en_name = item.get('career_en_name')
        mod.career_vi_name = item.get('career_vi_name')
        mod.ethnic_vi_name = item.get('ethnic_vi_name')
        mod.vi_nationality = item.get('vi_nationality')
        mod.en_nationality = item.get('en_nationality')
        mod.blood_group = item.get('blood_group')
        mod.blood_result_time = item.get('blood_result_time')
        mod.blood_rh = item.get('blood_rh')
        mod.qr_code_bhyt = item.get('qr_code_bhyt')
        mod.qr_code_cccd_chip = item.get('qr_code_cccd_chip')
        mod.created_date = item.get('created_date')
        mod.last_exam = item.get('last_exam')
        serialized_data = item.get('relative_name')
        unserialized_data = phpserialize.loads(serialized_data.encode(), decode_strings=True)
        mod.father_name = unserialized_data['father_name']
        mod.father_phone = unserialized_data['father_phone']
        mod.mother_name = unserialized_data['mother_name']
        mod.mother_phone = unserialized_data['mother_phone']
        try:
            cur.execute(
                "CALL public.insert_patient(%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)",
                (
                    mod.person_id,
                    mod.patient_code,
                    mod.patient_code_2,
                    mod.vaccination_code,
                    mod.name,
                    mod.backup_name,
                    mod.gender,
                    mod.date_of_birth,
                    mod.phone_number,
                    mod.full_address,
                    mod.career_vi_name,
                    mod.career_en_name,
                    mod.ethnic_vi_name,
                    mod.vi_nationality,
                    mod.en_nationality,
                    mod.blood_group,
                    mod.blood_rh,
                    mod.blood_result_time,
                    mod.qr_code_bhyt,
                    mod.qr_code_cccd_chip,
                    mod.created_date,
                    mod.last_exam,
                    mod.father_name,
                    mod.father_phone,
                    mod.mother_name,
                    mod.mother_phone
                )
            )
            conn.commit()
            log_terminal("Chèn thành công bệnh nhân : "+ mod.patient_code)
        except (Exception, psycopg2.Error) as error:
            print("Lỗi khi chèn dữ liệu vào PostgreSQL:", error)

    if conn:
        cur.close()
        conn.close()
        log_terminal("Hoàn thành cào dữ liệu bệnh nhân! Kết nối PostgreSQL đã đóng!.....")

    terminal_window.destroy()
    app.deiconify()

#app
def run_secondary_interface(main_app):
    global run_button, txt_search, app
    app = customtkinter.CTkToplevel(main_app)
    app.title("Lấy báo cáo sổ khám bệnh")
    center_window(app)
    icon_path = os.path.abspath("resource\crawlLogo.ico")
    app.iconbitmap(icon_path)

    def on_closing():
        main_app.deiconify()  # Hiển thị lại cửa sổ chính khi đóng cửa sổ mới
        app.destroy()

    app.protocol("WM_DELETE_WINDOW", on_closing)

    imgBG = ImageTk.PhotoImage(Image.open("resource\BG2.jpg"))
    l1 = customtkinter.CTkLabel(master=app, image=imgBG)
    l1.pack()

    frame = customtkinter.CTkFrame(master=l1, width=320, height=250, corner_radius=15)
    frame.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

    l2 = customtkinter.CTkLabel(master=frame, text="Input Patient Code Search", font=('Century Gothic', 20))
    l2.place(x=30, y=45)

    txt_search = customtkinter.CTkEntry(master=frame, placeholder_text="VD: 794...", font=('Century Gothic', 13), width=280)
    txt_search.place(x=20, y=100)

    run_button = customtkinter.CTkButton(master=frame, command=start_script_thread,text="Thực thi", font=('Tahoma', 13), fg_color="#005369", hover_color="#008097")
    run_button.place(x=160, y=200)

    app.mainloop()