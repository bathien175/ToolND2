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
import psycopg2
import threading
#global
customtkinter.set_appearance_mode("Light")
customtkinter.set_default_color_theme("green")
txt_search = Any
run_button = Any
app = Any
l4 = Any
l6 = Any
terminal_window = Any
urlNgoaiTru = "http://192.168.0.77/api/patients/getInvoices_outpatient/"
urlNoiTru = "http://192.168.0.77/api/patients/getLast_department/"
urlThuoc = "http://192.168.0.77/api/patients/getListPrescriptionOutpatient/"

def update_file_json(l4_value, l6_value):
    config = load_config()
    config['page_value'] = l4_value
    config['record_value'] = l6_value
    save_config(config)

def update_labels():
    config = load_config()
    l4_value = config['page_value']
    l6_value = config['record_value']
    # Cập nhật text cho l4 và l6
    l4.configure(text=l4_value)
    l6.configure(text=l6_value)

def load_config():
    if os.path.exists('KCB_info_step.json'):
        with open('KCB_info_step.json', 'r') as f:
            return json.load(f)
    return {"page_value": "1", "record_value": "0"}

def save_config(config):
    with open('KCB_info_step.json', 'w') as f:
        json.dump(config, f)

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

def start_script_thread():
    global script_thread  # Khai báo biến toàn cục
    terminal_window, terminal_text = open_terminal_window()
    app.withdraw()
    script_thread = threading.Thread(target=run_Script, args=(terminal_text,))
    script_thread.start()

def center_window(window, width=600, height=440):
    # Lấy kích thước màn hình
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    # Tính toán tọa độ x và y để cửa sổ ở giữa màn hình
    x = (screen_width / 2) - (width / 2)
    y = (screen_height / 2) - (height / 2)
    window.geometry(f'{width}x{height}+{int(x)}+{int(y)}')

def fetch_data_from_api(header):
    global urlNgoaiTru, urlNoiTru, urlThuoc
    all_data = []
    config = load_config()
    page_value = config['page_value']
    record_value = config['record_value']
    conn_params = sour.ConnectStr
    conn = psycopg2.connect(**conn_params)
    cur = conn.cursor()
    while(True):
        try:
            #page_offset = page_value - 1
            #start_index = page_offset * 20 - 1
            queryStr = f"SELECT * FROM patient order by stt asc OFFSET {record_value} LIMIT 20;"
            cur.execute(queryStr)
            listdata = cur.fetchall()
            if len(listdata) > 0:
                for item in listdata:
                    patient_id = item[1]
                    fullUrl = urlXaPhuong + str(code_district)
                    response = requests.get(fullUrl, headers=header)
                    if response.status_code == 200:
                        dataf = response.json()
                        all_data.extend(dataf['data'])
                    else:
                        print(f"Lỗi khi lấy trang dữ liệu...")
            

        except:
            print("Lỗi xảy ra trong quá trình truy cập CSDL...")
        finally:
            if conn:
                cur.close()
                conn.close()

    return all_data

def run_Script(terminal_text):
    global app, txt_search, terminal_window, int_selection
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
        all_data_fetch = fetch_data_from_api(int_selection, headers)
    finally:
        loading = False
        loading_thread.join()
    log_terminal(".........................Đã hoàn tất thu thập danh sách..................................")
    conn_params = sour.ConnectStr
    conn = psycopg2.connect(**conn_params)
    cur = conn.cursor()
    for item in all_data_fetch:
        match int_selection.get():
            case "Nghề nghiệp":
                mod = model.CareerModel()
                mod.career_id = item.get('career_id')
                mod.code = item.get('code')
                mod.disable = item.get('disable')
                mod.en_name = item.get('en_name')
                mod.vi_name = item.get('vi_name')
                mod.ma_nghenghiep_bhyt = item.get('ma_nghenghiep_bhyt')
                try:
                    cur.execute(
                        "CALL public.insert_career(%s, %s, %s, %s, %s, %s)",
                        (
                            mod.career_id,
                            mod.code,
                            mod.disable,
                            mod.en_name,
                            mod.vi_name,
                            mod.ma_nghenghiep_bhyt
                        )
                    )
                    conn.commit()
                    log_terminal("Chèn thành công nghề nghiệp : "+ mod.vi_name)
                except (Exception, psycopg2.Error) as error:
                    log_terminal("Lỗi khi chèn dữ liệu vào PostgreSQL:", error)
            case "Quốc Tịch":
                mod = model.NationalityModel()
                mod.nationality_id = item.get('nationality_id')
                mod.disable = item.get('disable')
                mod.en_name = item.get('en_name')
                mod.vi_name = item.get('vi_name')
                mod.ma_quoc_tich_bhyt = item.get('ma_quoc_tich_bhyt')
                try:
                    cur.execute(
                        "CALL public.insert_nationality(%s, %s, %s, %s, %s)",
                        (
                            mod.nationality_id,
                            mod.disable,
                            mod.en_name,
                            mod.vi_name,
                            mod.ma_quoc_tich_bhyt
                        )
                    )
                    conn.commit()
                    log_terminal("Chèn thành công quốc tịch : "+ mod.vi_name)
                except (Exception, psycopg2.Error) as error:
                    log_terminal("Lỗi khi chèn dữ liệu vào PostgreSQL:", error)
            case "Quốc Gia":
                mod = model.CountryModel()
                mod.country_id = item.get('country_id')
                mod.disable = item.get('disable')
                mod.en_name = item.get('en_name')
                mod.vi_name = item.get('vi_name')
                try:
                    cur.execute(
                        "CALL public.insert_country(%s, %s, %s, %s)",
                        (
                            mod.country_id,
                            mod.disable,
                            mod.en_name,
                            mod.vi_name
                        )
                    )
                    conn.commit()
                    log_terminal("Chèn thành công quốc gia : "+ mod.vi_name)
                except (Exception, psycopg2.Error) as error:
                    log_terminal("Lỗi khi chèn dữ liệu vào PostgreSQL:", error)
            case "Tỉnh Thành":
                mod = model.ProvinceModel()
                mod.province_id = item.get('province_id')
                mod.code = item.get('code')
                mod.disabled = item.get('disabled')
                mod.en_name = item.get('en_name')
                mod.vi_name = item.get('vi_name')
                mod.ma_tinh = item.get('ma_tinh')
                try:
                    cur.execute(
                        "CALL public.insert_province(%s, %s, %s, %s, %s, %s)",
                        (
                            mod.province_id,
                            mod.code,
                            mod.disabled,
                            mod.en_name,
                            mod.vi_name,
                            mod.ma_tinh
                        )
                    )
                    conn.commit()
                    log_terminal("Chèn thành công tỉnh thành : "+ mod.vi_name)
                except (Exception, psycopg2.Error) as error:
                    log_terminal("Lỗi khi chèn dữ liệu vào PostgreSQL:", error)
            case "Quận Huyện":
                if len(item) != 3:
                    mod = model.DistrictModel()
                    mod.district_id = item.get('district_id')
                    mod.province_id = item.get('province_id')
                    mod.ma_quan = item.get('ma_quan')
                    mod.code = item.get('code')
                    mod.disabled = item.get('disabled')
                    mod.en_name = item.get('en_name')
                    mod.vi_name = item.get('vi_name')
                    mod.ma_quan_bhyt = item.get('ma_quan_bhyt')
                    try:
                        cur.execute(
                            "CALL public.insert_district(%s, %s, %s, %s, %s, %s, %s, %s)",
                            (
                                mod.district_id,
                                mod.province_id,
                                mod.ma_quan,
                                mod.code,
                                mod.disabled,
                                mod.en_name,
                                mod.vi_name,
                                mod.ma_quan_bhyt
                            )
                        )
                        conn.commit()
                        log_terminal("Chèn thành công quận huyện : "+ mod.vi_name)
                    except (Exception, psycopg2.Error) as error:
                        log_terminal("Lỗi khi chèn dữ liệu vào PostgreSQL:", error)
            case "Xã Phường":
                mod = model.WardModel()
                mod.ward_id = item.get('ward_id')
                mod.district_id = item.get('district_id')
                mod.disabled = item.get('disabled')
                mod.en_name = item.get('en_name')
                mod.vi_name = item.get('vi_name')
                mod.ma_phuong = item.get('ma_phuong')
                mod.ma_phuong_bhyt = item.get('ma_phuong_bhyt')
                mod.auto_suggest_code = item.get('auto_suggest_code')
                try:
                    cur.execute(
                        "CALL public.insert_ward(%s, %s, %s, %s, %s, %s, %s, %s)",
                        (
                            mod.ward_id,
                            mod.district_id,
                            mod.disabled,
                            mod.en_name,
                            mod.vi_name,
                            mod.ma_phuong,
                            mod.ma_phuong_bhyt,
                            mod.auto_suggest_code
                        )
                    )
                    conn.commit()
                    log_terminal("Chèn thành công xã phường : "+ mod.vi_name)
                except (Exception, psycopg2.Error) as error:
                    log_terminal("Lỗi khi chèn dữ liệu vào PostgreSQL:", error)

    if conn:
        cur.close()
        conn.close()
        messagebox.showinfo(title="Thành công!",message="Hoàn thành cào dữ liệu bệnh nhân! Kết nối PostgreSQL đã đóng!.....")
    
    sour._destroySelenium_()
    terminal_window.destroy()
    app.deiconify()

#app
def run_secondary_interface(main_app):
    global run_button, txt_search, app, l4, l6
    app = customtkinter.CTkToplevel(main_app)
    app.title("Lấy dữ liệu hành chính")
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

    l2 = customtkinter.CTkLabel(master=frame, text="Thông tin cào khám chữa bệnh", font=('Tahoma', 20))
    l2.place(x=40, y=45)
    
    l3 = customtkinter.CTkLabel(master=frame, text="Thứ tự trang: ", font=('Tahoma', 14))
    l3.place(x=20, y=70)
    l4 = customtkinter.CTkLabel(master=frame, text="1", font=('Tahoma', 14))
    l4.place(x=70, y=70)

    l5 = customtkinter.CTkLabel(master=frame, text="Số record: ", font=('Tahoma', 14))
    l5.place(x=20, y=100)
    l6 = customtkinter.CTkLabel(master=frame, text="0", font=('Tahoma', 14))
    l6.place(x=70, y=100)

    update_labels()

    run_button = customtkinter.CTkButton(master=frame, command=start_script_thread,text="Thực thi", font=('Tahoma', 13), fg_color="#005369", hover_color="#008097")
    run_button.place(x=160, y=200)

    app.mainloop()