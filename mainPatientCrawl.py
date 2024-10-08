#using
import os
import time
import tkinter as tk
import customtkinter
from PIL import ImageTk, Image
from babel.numbers import *
from tkinter import messagebox
import json
import sourceString as sour
import sourceModel as model
import requests
import math
import phpserialize
import psycopg2
import threading
import sys
import mysql.connector as connector
#global
customtkinter.set_appearance_mode("Light")
customtkinter.set_default_color_theme("green")
txt_search = Any
run_button = Any
run2_button = Any
app = Any
urlAPIfind = "http://192.168.0.77/api/patients/find"
urlAPIload = "http://192.168.0.77/api/patients/load_patientById/"
terminal_window = Any

def validate_textSearch():
    global txt_search
    if txt_search.get() != "":
        return True
    else:
        return False

def open_terminal_window():
    global terminal_window
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

def sanitize_date(date_string):
    if date_string == "0000-00-00 00:00:00" or not date_string:
        return None
    return date_string

def start_script_thread(selectMode):
    global script_thread  # Khai báo biến toàn cục
    if selectMode == 1:
        if validate_textSearch() == True:
            terminal_window, terminal_text = open_terminal_window()
            app.withdraw()
            script_thread = threading.Thread(target=run_Script, args=(terminal_text,selectMode))
            script_thread.start()
        else:
            messagebox.showerror(title="Lỗi thông tin", message="Vui lòng không bỏ trống dữ liệu tìm kiếm!")
    elif selectMode == 0:
        terminal_window, terminal_text = open_terminal_window()
        app.withdraw()
        script_thread = threading.Thread(target=run_Script, args=(terminal_text,selectMode))
        script_thread.start()
    else:
        terminal_window, terminal_text = open_terminal_window()
        app.withdraw()
        script_thread = threading.Thread(target=run_Script, args=(terminal_text,selectMode))
        script_thread.start()

def center_window(window, width=600, height=440):
    # Lấy kích thước màn hình
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    # Tính toán tọa độ x và y để cửa sổ ở giữa màn hình
    x = (screen_width / 2) - (width / 2)
    y = (screen_height / 2) - (height / 2)
    window.geometry(f'{width}x{height}+{int(x)}+{int(y)}')

def fetch_one_person(base_url, headers):
    # Tên file để lưu trạng thái
    state_file = 'search_state.json'
    
    # Đọc trạng thái từ file nếu tồn tại
    if os.path.exists(state_file):
        with open(state_file, 'r') as f:
            search_state = json.load(f)
    else:
        search_state = {}
    
    # Lấy trang hiện tại từ trạng thái đã lưu hoặc bắt đầu từ bệnh nhân số 1
    patientIdCurrent = search_state.get("patient_id", 1)
    
    headersFix = headers
    demloi = 1
    
    while True:
        full_url = base_url + str(patientIdCurrent)
        response = requests.get(full_url, headers=headersFix)

        if response.status_code == 200:
            try:
                dataf = response.json()
                datal = dataf['data']
                data = datal['patient']

                if not data:
                    print(f"Thông tin bệnh nhân bị rỗng!...")
                else:
                    add_onedata_to_database(data)
                    demloi = 1
                
                # Lưu trạng thái hiện tại
                patientIdCurrent += 1
                search_state['patient_id'] = patientIdCurrent #1
                with open(state_file, 'w') as f:
                    json.dump(search_state, f)
                
            except:
                print(f"Lỗi khi lấy bệnh nhân {patientIdCurrent} với mã lỗi: {response.status_code}")
                demloi +=1
        else:
            if response.status_code == 10000:
                print("Lỗi out session đang cố kết nối lại...")
                bTry = False
                errorChrome = 1
                while bTry == False:
                    bTry = sour._initSelenium_()
                    if bTry == False:
                        if errorChrome >= 10:
                            app.destroy() 
                        else:
                            errorChrome += 1
                time.sleep(3)
                # login
                sour._login_("quyen.ngoq", "74777477")
                # ấn nút login
                time.sleep(0.5)
                headersFix['Userkey'] = sour.secretKey
                headersFix['Authorization'] = sour.secretKey
                print("Kết nối thành công...")  
                demloi = 1    
            else:
                print(f"Lỗi khi lấy bệnh nhân {patientIdCurrent} với mã lỗi: {response.status_code}")
                demloi +=1

        if demloi == 10:
            print('Lỗi quá nhiều lần! Đóng quá trình lấy trang!')
            break

    print(f"Hoàn tất lấy dữ liệu thành công!...")

def fetch_one_person_old(base_url, headers):
    # Tên file để lưu trạng thái
    state_file = 'search_state.json'
    
    # Đọc trạng thái từ file nếu tồn tại
    if os.path.exists(state_file):
        with open(state_file, 'r') as f:
            search_state = json.load(f)
    else:
        search_state = {}
    
    # Lấy trang hiện tại từ trạng thái đã lưu hoặc bắt đầu từ bệnh nhân số 1
    patientIdCurrent = search_state.get("patient_id", 1)
    listpatientOld = loadData(patientIdCurrent)
    headersFix = headers
    demloi = 1
    
    for item in listpatientOld:
        patientIdOld = item[0]
        full_url = base_url + str(patientIdOld)
        response = requests.get(full_url, headers=headersFix)

        if response.status_code == 200:
            try:
                dataf = response.json()
                datal = dataf['data']
                data = datal['patient']

                if not data:
                    print(f"Thông tin bệnh nhân bị rỗng!...")
                else:
                    add_onedata_to_database(data)
                    demloi = 1
                
                search_state['patient_id'] = patientIdOld
                with open(state_file, 'w') as f:
                    json.dump(search_state, f)
                
            except:
                print(f"Lỗi khi lấy bệnh nhân {patientIdOld} với mã lỗi: {response.status_code}")
                demloi +=1
        else:
            if response.status_code == 10000:
                print("Lỗi out session đang cố kết nối lại...")
                bTry = False
                errorChrome = 1
                while bTry == False:
                    bTry = sour._initSelenium_()
                    if bTry == False:
                        if errorChrome >= 10:
                            app.destroy() 
                        else:
                            errorChrome += 1
                time.sleep(3)
                # login
                sour._login_("quyen.ngoq", "74777477")
                # ấn nút login
                time.sleep(0.5)
                headersFix['Userkey'] = sour.secretKey
                headersFix['Authorization'] = sour.secretKey
                print("Kết nối thành công...")  
                demloi = 1    
            else:
                print(f"Lỗi khi lấy bệnh nhân {patientIdCurrent} với mã lỗi: {response.status_code}")
                demloi +=1

        if demloi == 10:
            print('Lỗi quá nhiều lần! Đóng quá trình lấy trang!')
            break

    print(f"Hoàn tất lấy dữ liệu thành công!...")

def fetch_all_pages(base_url, headers, payload):
    # Tên file để lưu trạng thái
    state_file = 'search_state.json'
    
    # Đọc trạng thái từ file nếu tồn tại
    if os.path.exists(state_file):
        with open(state_file, 'r') as f:
            search_state = json.load(f)
    else:
        search_state = {}
    
    # Lấy textsearch từ payload
    textsearch = payload.get('patient_code', '')
    
    # Lấy trang hiện tại từ trạng thái đã lưu hoặc bắt đầu từ trang 1
    current_page = search_state.get(textsearch, 1)
    
    headersFix = headers
    demloi = 1
    
    while True:
        payloads = payload.copy()
        payloads['page'] = current_page

        response = requests.get(base_url, headers=headersFix, json=payloads)

        if response.status_code == 200:
            try:
                dataf = response.json()
                data = dataf['data']
                data_list = data['data_list']
                
                if not data_list:
                    print(f"Đã đạt đến cuối danh sách ở trang {current_page - 1}")
                    # Xóa trạng thái khi hoàn thành
                    if textsearch in search_state:
                        del search_state[textsearch]
                    break
                
                add_data_to_database(data_list)
                
                paging = data['paging']
                total_record = paging['total_record']
                row_per_page = paging['row_per_page']
                
                print(f"Đã lấy trang {current_page}. Tổng số bản ghi hiện tại: {total_record}")
                
                current_page += 1
                demloi = 1
                
                # Lưu trạng thái hiện tại
                search_state[textsearch] = current_page
                with open(state_file, 'w') as f:
                    json.dump(search_state, f)
                
            except:
                print(f"Lỗi khi lấy trang {current_page}: {response.status_code}")
                demloi +=1
        else:
            if response.status_code == 10000:
                print("Lỗi out session đang cố kết nối lại...")
                bTry = False
                errorChrome = 1
                while bTry == False:
                    bTry = sour._initSelenium_()
                    if bTry == False:
                        if errorChrome >= 10:
                            app.destroy() 
                        else:
                            errorChrome += 1
                time.sleep(3)
                # login
                sour._login_("quyen.ngoq", "74777477")
                # ấn nút login
                time.sleep(0.5)
                headersFix['Userkey'] = sour.secretKey
                headersFix['Authorization'] = sour.secretKey
                print("Kết nối thành công...")  
                demloi = 1    
            else:
                print(f"Lỗi khi lấy trang {current_page}: {response.status_code}")
                demloi +=1

        if demloi == 10:
            print('Lỗi quá nhiều lần! Đóng quá trình lấy trang!')
            break

    print(f"Hoàn tất lấy dữ liệu thành công!...")

def add_data_to_database(datalist):
    conn_params = sour.ConnectStr
    conn = psycopg2.connect(**conn_params)
    cur = conn.cursor()
    for item in datalist:
        mod = model.PatientModel()
        mod.person_id = item.get('person_id')
        mod.patient_code = item.get('patient_code')
        mod.patient_code_2 = item.get('patient_code2')
        mod.vaccination_code = item.get('vaccination_code')
        mod.name = item.get('name')
        mod.backup_name = item.get('backup_name')
        mod.gender = item.get('gender')
        mod.date_of_birth = item.get('date_of_birth')
        if mod.date_of_birth == "0000-00-00":
            mod.date_of_birth = None
        mod.phone_number = item.get('phone_number')
        mod.full_address = item.get('full_address')
        mod.career_en_name = item.get('career_en_name')
        mod.career_vi_name = item.get('career_vi_name')
        mod.ethnic_vi_name = item.get('ethnic_vi_name')
        mod.vi_nationality = item.get('vi_nationality')
        mod.en_nationality = item.get('en_nationality')
        mod.blood_group = item.get('blood_group')
        mod.blood_result_time = sanitize_date(item.get('blood_result_time'))
        mod.blood_rh = item.get('blood_rh')
        mod.qr_code_bhyt = item.get('qr_code_bhyt')
        mod.qr_code_cccd_chip = item.get('qr_code_cccd_chip')
        mod.created_date = sanitize_date(item.get('created_date'))
        mod.last_exam = sanitize_date(item.get('last_exam'))
        if item.get('relative_name') == '' or not item.get('relative_name'):
            mod.father_name = ""
            mod.father_phone = ""
            mod.mother_name = ""
            mod.mother_phone = ""
        else:
            serialized_data = item.get('relative_name')
            try:
                unserialized_data =  phpserialize.loads(serialized_data.encode('utf-8'), decode_strings=True)
                mod.father_name = unserialized_data['father_name']
                mod.father_phone = unserialized_data['father_phone']
                mod.mother_name = unserialized_data['mother_name']
                mod.mother_phone = unserialized_data['mother_phone']
            except Exception as e:
                mod.father_name = ""
                mod.father_phone = ""
                mod.mother_name = ""
                mod.mother_phone = "" 
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
            print("Chèn thành công bệnh nhân : "+ mod.patient_code)
        except (Exception, psycopg2.Error) as error:
            print("Lỗi khi chèn dữ liệu vào PostgreSQL:", error)

    if conn:
        cur.close()
        conn.close()

def add_onedata_to_database(item):
    conn_params = sour.ConnectStr
    conn = psycopg2.connect(**conn_params)
    cur = conn.cursor()

    mod = model.PatientModel()
    mod.person_id = item.get('person_id')
    mod.patient_code = item.get('patient_code')
    mod.patient_code_2 = item.get('patient_code2')
    mod.vaccination_code = item.get('vaccination_code')
    mod.name = item.get('name')
    mod.backup_name = item.get('backup_name')
    mod.gender = item.get('gender')
    mod.date_of_birth = item.get('date_of_birth')
    if mod.date_of_birth == "0000-00-00":
        mod.date_of_birth = None
    mod.phone_number = item.get('phone_number')
    mod.full_address = item.get('full_address')
    mod.career_en_name = item.get('career_en_name')
    mod.career_vi_name = item.get('career_vi_name')
    mod.ethnic_vi_name = item.get('ethnic_vi_name')
    mod.vi_nationality = item.get('vi_nationality')
    mod.en_nationality = item.get('vi_nationality')
    mod.blood_group = item.get('blood_group')
    mod.blood_result_time = sanitize_date(item.get('blood_result_time'))
    mod.blood_rh = item.get('blood_rh')
    mod.qr_code_bhyt = item.get('qr_code_bhyt')
    mod.qr_code_cccd_chip = item.get('qr_code_cccd_chip')
    mod.created_date = sanitize_date(item.get('created_date'))
    mod.last_exam = sanitize_date(item.get('last_exam'))
    if item.get('relative_name') == '' or not item.get('relative_name'):
        mod.father_name = ""
        mod.father_phone = ""
        mod.mother_name = ""
        mod.mother_phone = ""
    else:
        serialized_data = item.get('relative_name')
        try:
            unserialized_data =  phpserialize.loads(serialized_data.encode('utf-8'), decode_strings=True)
            mod.father_name = unserialized_data['father_name']
            mod.father_phone = unserialized_data['father_phone']
            mod.mother_name = unserialized_data['mother_name']
            mod.mother_phone = unserialized_data['mother_phone']
        except Exception as e:
            mod.father_name = ""
            mod.father_phone = ""
            mod.mother_name = ""
            mod.mother_phone = "" 
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
        print("Chèn thành công bệnh nhân : "+ mod.patient_code)
    except (Exception, psycopg2.Error) as error:
        print("Lỗi khi chèn dữ liệu vào PostgreSQL:", error)

    if conn:
        cur.close()
        conn.close()

def run_Script(terminal_text, selectMode):
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
    if selectMode == 1:
        payload = {
            "page": 1,
            "patient_code": txt_search.get(),
        }
    log_terminal(".........................Đang tiến hành thu thập! Vui lòng chờ...........................")
    # Biến global để kiểm soát animation
    loading = True
    def loading_animation():
        chars = "/—\\|"
        i = 0
        while loading:
            log_terminal('\r' + 'Vui lòng đợi quá trình tải dữ liệu đang được diễn ra... ' + chars[i % len(chars)])
            time.sleep(0.1)
            i += 1
    # Thread cho animation
    loading_thread = threading.Thread(target=loading_animation)
    loading_thread.daemon = True  # Đảm bảo thread sẽ kết thúc khi chương trình chính kết thúc
    # Bắt đầu animation
    loading_thread.start()
    try:
        if selectMode == 1:
            fetch_all_pages(urlAPIfind, headers, payload)
        elif selectMode == 0:
            fetch_one_person(urlAPIload, headers)
        else:
            fetch_one_person_old(urlAPIload, headers)
    finally:
        loading = False
        loading_thread.join()
    messagebox.showinfo(title="Thành công!",message="Hoàn thành cào dữ liệu bệnh nhân! Kết nối PostgreSQL đã đóng!.....")
    sour._destroySelenium_()
    terminal_window.destroy()
    app.deiconify()

def loadData(pid):
    mydb = connector.connect(
            host="192.168.0.127",
            port="3306",
            user="root",
            password="So17052001",
            database="nhidong2_full")
    mycursor = mydb.cursor()
    query_string = f"select person_id from patient p where person_id >= {pid}"
    listPatients = mycursor.execute(query_string, multi=True)
    all_results = []
    for result in listPatients:
        if result.with_rows:
            all_results = result.fetchall()
            break
    return all_results

#app
def run_secondary_interface(main_app):
    global run_button, txt_search, app, run2_button
    app = customtkinter.CTkToplevel(main_app)
    app.title("Lấy dữ liệu bệnh nhân")
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

    run2_button = customtkinter.CTkButton(master=frame, command=lambda: start_script_thread(0),text="Chạy theo ID", font=('Tahoma', 13), fg_color="#005369", hover_color="#008097")
    run2_button.place(x=20, y=150)

    run_button = customtkinter.CTkButton(master=frame, command=lambda: start_script_thread(1),text="Thực thi", font=('Tahoma', 13), fg_color="#005369", hover_color="#008097")
    run_button.place(x=160, y=150)

    run3_button = customtkinter.CTkButton(master=frame, command=lambda: start_script_thread(2),text="Dữ liệu cũ", font=('Tahoma', 13), fg_color="#005369", hover_color="#008097")
    run3_button.place(x=100, y=200)

    app.mainloop()