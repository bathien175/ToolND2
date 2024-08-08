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
urlNoiTru = "http://192.168.0.77/api/patients/getListInpatientOfPatient/"
urlThuoc = "http://192.168.0.77/api/patients/getListPrescriptionOutpatient/"

def convert_to_timestamp(value):
    if isinstance(value, str):
        try:
            # Thử chuyển đổi chuỗi thành timestamp
            return datetime.strptime(value, "%Y-%m-%d")
        except ValueError:  
            # Nếu không thành công, có thể là đã là timestamp hoặc định dạng khác
            pass
    return value

def update_file_json(l4_value, l6_value):
    config = load_config()
    config['page_value'] = l4_value
    config['record_value'] = l6_value
    save_config(config)

def sanitize_date(date_string):
    if date_string == "0000-00-00 00:00:00" or not date_string or date_string.startswith("0000-00-00"):
        return None
    return date_string

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

def add_ngoaitru_to_database(datalist, patientId):
    conn_params = sour.ConnectStr
    conn = psycopg2.connect(**conn_params)
    cur = conn.cursor()
    if len(datalist) > 0:
        for d in datalist:
            try:
                insert_query = """
                INSERT INTO InvoiceOutPatient (
                    amount, bhyt, cashier_name, created, created_time, deleted, 
                    enum_examination_type, enum_invoice_type, enum_patient_type, 
                    enum_payment_type, icd10, ins_paid_price, invoice_code, 
                    new_invoice_code, note, patient_price, pay_payment_account_document_id, 
                    pay_payment_account_id, pay_receipt_id, pay_return, refund_invoice_code, 
                    returned, so_hd, so_luu_tru, ticket_id, total_price, 
                    total_returned_amount, treatment_id, patient_id
                ) VALUES (
                    %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
                    %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s
                )
                """
                # Chuẩn bị dữ liệu để chèn
                createtxt = any
                enum_exam = 0
                totalReturn = Any
                try:
                    createtxt = sanitize_date(d["created"])
                except:
                    createtxt = sanitize_date(d["created_date"])

                try:
                    enum_exam = sanitize_date(d["enum_examination_type"])
                except:
                    enum_exam = None

                try:
                    totalReturn = sanitize_date(d["total_returned_amount"])
                except:
                    totalReturn = None
                data = (
                    d["amount"],  # amount
                    d["bhyt"],  # bhyt
                    d["cashier_name"],  # cashier_name
                    createtxt,  # created
                    sanitize_date(d["created_time"]),  # created_time
                    d["deleted"],  # deleted
                    enum_exam,  # enum_examination_type
                    d["enum_invoice_type"],  # enum_invoice_type
                    d["enum_patient_type"],  # enum_patient_type
                    d["enum_payment_type"],  # enum_payment_type
                    d["icd10"],  # icd10
                    d["ins_paid_price"],  # ins_paid_price
                    d["invoice_code"],  # invoice_code
                    d["new_invoice_code"],  # new_invoice_code
                    d["note"],  # note
                    d["patient_price"],  # patient_price
                    d["pay_payment_account_document_id"],  # pay_payment_account_document_id
                    d["pay_payment_account_id"],  # pay_payment_account_id
                    d["pay_receipt_id"],  # pay_receipt_id
                    d["pay_return"],  # pay_return
                    d["refund_invoice_code"],  # refund_invoice_code
                    d["returned"],  # returned
                    d["so_hd"],  # so_hd
                    d["so_luu_tru"],  # so_luu_tru
                    d["ticket_id"],  # ticket_id
                    d["total_price"],  # total_price
                    totalReturn,  # total_returned_amount
                    d["treatment_id"],  # treatment_id,
                    patientId
                )

                # Thực thi câu lệnh INSERT
                cur.execute(insert_query, data)

                # Commit thay đổi
                conn.commit()

                print("Dữ liệu hóa đơn ngoại trú đã được chèn thành công!")
            except Exception as e:
                print("Lỗi ngoại trú "+ e)
    # Đóng cursor và kết nối
    cur.close()
    conn.close()

def add_thuoc_to_database(datalist):
    conn_params = sour.ConnectStr
    conn = psycopg2.connect(**conn_params)
    cur = conn.cursor()
    if len(datalist) > 0:
        for d in datalist:
            try:
                # Chuẩn bị câu lệnh INSERT
                insert_query = """
                INSERT INTO Prescription (
                    backup_medical_record_id, canh_bao_hoat_chat, canh_bao_hoat_chat_dac_biet, 
                    canh_bao_qua_lieu, canh_bao_thang_tuoi_chi_dinh, check_in_out_hospital_record_id, 
                    chi_tiet_icd_canh_bao, clinical_mani, code, create_date, created_time, da_cap, 
                    date, date_expired, date_issued, day_num, diagnosis, doctor_id, doctor_name, 
                    donthuocquocgia_checksum, donthuocquocgia_error, donthuocquocgia_modified_person_id, 
                    donthuocquocgia_modified_time, dst_invoice_id, e_invoice_id, e_invoice_status, 
                    enum_customer_type, fourth_icd10, fourth_icd10_code, ghi_chu_toa_h, giam_sat_note, 
                    giam_sat_time, giam_sat_user_id, icd10, icd10_note, in_toa_thuoc, input_date, 
                    insurance_code, insurance_price, invoice_code, is_confirm_app, is_service, 
                    khth_duyet, khth_tg_duyet, khth_use_id, _locked, medical_record_id, note, 
                    patient_height, patient_id, patient_temp, patient_weight, _percent, prescription_id, 
                    prescription_type, primary_icd10, primary_icd10_code, processed, queue_code, 
                    re_examine, reason_decline, reexam_date, second_icd10, second_icd10_code, 
                    shift_id, sign_date, signature_url, status, third_icd10, third_icd10_code, 
                    ticket_id, time_decline, total_price, trang_thai_giam_sat_toa_id, tuong_tac_thuoc, 
                    updated_at, user_id_decline, zone_code
                ) VALUES (
                    %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
                    %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
                    %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
                    %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s
                )
                """

                # Chuẩn bị dữ liệu để chèn
                data = (
                    d["backup_medical_record_id"],  # backup_medical_record_id
                    d["canh_bao_hoat_chat"],  # canh_bao_hoat_chat
                    d["canh_bao_hoat_chat_dac_biet"],  # canh_bao_hoat_chat_dac_biet
                    d["canh_bao_qua_lieu"],  # canh_bao_qua_lieu
                    d["canh_bao_thang_tuoi_chi_dinh"],  # canh_bao_thang_tuoi_chi_dinh
                    d["check_in_out_hospital_record_id"],  # check_in_out_hospital_record_id
                    d["chi_tiet_icd_canh_bao"],  # chi_tiet_icd_canh_bao
                    d["clinical_mani"],  # clinical_mani
                    d["code"],  # code
                    sanitize_date(d["create_date"]),  # create_date
                    sanitize_date(d["created_time"]),  # created_time
                    d["da_cap"],  # da_cap
                    d["date"],  # date
                    d["date_expired"],  # date_expired
                    d["date_issued"],  # date_issued
                    d["day_num"],  # day_num
                    d["diagnosis"],  # diagnosis
                    d["doctor_id"],  # doctor_id
                    d["doctor_name"],  # doctor_name
                    d["donthuocquocgia_checksum"],  # donthuocquocgia_checksum
                    d["donthuocquocgia_error"],  # donthuocquocgia_error
                    d["donthuocquocgia_modified_person_id"],  # donthuocquocgia_modified_person_id
                    sanitize_date(d["donthuocquocgia_modified_time"]),  # donthuocquocgia_modified_time
                    d["dst_invoice_id"],  # dst_invoice_id
                    d["e_invoice_id"],  # e_invoice_id
                    d["e_invoice_status"],  # e_invoice_status
                    d["enum_customer_type"],  # enum_customer_type
                    d["fourth_icd10"],  # fourth_icd10
                    d["fourth_icd10_code"],  # fourth_icd10_code
                    d["ghi_chu_toa_h"],  # ghi_chu_toa_h
                    d["giam_sat_note"],  # giam_sat_note
                    sanitize_date(d["giam_sat_time"]),  # giam_sat_time
                    d["giam_sat_user_id"],  # giam_sat_user_id
                    d["icd10"],  # icd10
                    d["icd10_note"],  # icd10_note
                    d["in_toa_thuoc"],  # in_toa_thuoc
                    sanitize_date(d["input_date"]),  # input_date
                    d["insurance_code"],  # insurance_code
                    d["insurance_price"],  # insurance_price
                    d["invoice_code"],  # invoice_code
                    d["is_confirm_app"],  # is_confirm_app
                    d["is_service"],  # is_service
                    d["khth_duyet"],  # khth_duyet
                    sanitize_date(d["khth_tg_duyet"]),  # khth_tg_duyet
                    d["khth_use_id"],  # khth_use_id
                    d["locked"],  # _locked
                    d["medical_record_id"],  # medical_record_id
                    d["note"],  # note
                    d["patient_height"],  # patient_height
                    d["patient_id"],  # patient_id
                    d["patient_temp"],  # patient_temp
                    d["patient_weight"],  # patient_weight
                    d["percent"],  # _percent
                    d["prescription_id"],  # prescription_id
                    d["prescription_type"],  # prescription_type
                    d["primary_icd10"],  # primary_icd10
                    d["primary_icd10_code"],  # primary_icd10_code
                    d["processed"],  # processed
                    d["queue_code"],  # queue_code
                    d["re_examine"],  # re_examine
                    d["reason_decline"],  # reason_decline
                    sanitize_date(d["reexam_date"]),  # reexam_date
                    d["second_icd10"],  # second_icd10
                    d["second_icd10_code"],  # second_icd10_code
                    d["shift_id"],  # shift_id
                    sanitize_date(d["sign_date"]),  # sign_date
                    d["signature_url"],  # signature_url
                    d["status"],  # status
                    d["third_icd10"],  # third_icd10
                    d["third_icd10_code"],  # third_icd10_code
                    d["ticket_id"],  # ticket_id
                    sanitize_date(d["time_decline"]),  # time_decline
                    d["total_price"],  # total_price
                    d["trang_thai_giam_sat_toa_id"],  # trang_thai_giam_sat_toa_id
                    d["tuong_tac_thuoc"],  # tuong_tac_thuoc
                    sanitize_date(d["updated_at"]),  # updated_at
                    d["user_id_decline"],  # user_id_decline
                    d["zone_code"]  # zone_code
                )

                # Thực thi câu lệnh INSERT
                cur.execute(insert_query, data)

                # Commit thay đổi
                conn.commit()

                print("Dữ liệu thuốc đã được chèn thành công!")
            except Exception as e:
                print("Lỗi thuốc "+ e)
    # Đóng cursor và kết nối
    cur.close()
    conn.close()
def add_noitru_to_database(datalist):
    conn_params = sour.ConnectStr
    conn = psycopg2.connect(**conn_params)
    cur = conn.cursor()
    if len(datalist) > 0:
        for d in datalist:
            try:
                # Chuẩn bị câu lệnh INSERT
                insert_query = """
                INSERT INTO InvoiceInPatient (
                    allow_re_print_hos_con, amount, bhbl_url, bhyt, cashier_name, check_in_date,
                    check_in_out_hospital_record_id, check_out_date, closed, created, created_time,
                    deleted, discount_amount, end_department, end_department_id, enum_account_type,
                    enum_examination_type, enum_invoice_type, enum_patient_type, enum_payment_type,
                    first_department, first_department_id, icd10, icd10_id, ins_paid_price,
                    ins_percentage, introduction_phone, introduction_text, invoice_code,
                    new_invoice_code, note, patient_id, patient_price, pay_payment_account_document_id,
                    pay_payment_account_id, pay_receipt_id, pay_return, phai_thu, phai_tra,
                    print_notice_payment, real_ins_payment, real_ser_payment, refund_invoice_code,
                    returned, so_hd, so_luu_tru, surgery_type, ticket_id, total_foresee,
                    total_ins_paid_price, total_ins_price, total_price, total_real_payment,
                    total_ser_price, treatment_id
                ) VALUES (
                    %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
                    %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
                    %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s
                )
                """

                # Chuẩn bị dữ liệu để chèn
                data = (
                    d["allow_re_print_hos_con"],  # allow_re_print_hos_con
                    d["amount"],  # amount
                    d["bhbl_url"],  # bhbl_url
                    d["bhyt"],  # bhyt
                    d["cashier_name"],  # cashier_name
                    sanitize_date(d["check_in_date"]),  # check_in_date
                    d["check_in_out_hospital_record_id"],  # check_in_out_hospital_record_id
                    sanitize_date(d["check_out_date"]),  # check_out_date
                    d["closed"],  # closed
                    sanitize_date(d["created"]),  # created
                    sanitize_date(d["created_time"]),  # created_time
                    d["deleted"],  # deleted
                    d["discount_amount"],  # discount_amount
                    d["end_department"],  # end_department
                    d["end_department_id"],  # end_department_id
                    d["enum_account_type"],  # enum_account_type
                    d["enum_examination_type"],  # enum_examination_type
                    d["enum_invoice_type"],  # enum_invoice_type
                    d["enum_patient_type"],  # enum_patient_type
                    d["enum_payment_type"],  # enum_payment_type
                    d["first_department"],  # first_department
                    d["first_department_id"],  # first_department_id
                    d["icd10"],  # icd10
                    d["icd10_id"],  # icd10_id
                    d["ins_paid_price"],  # ins_paid_price
                    d["ins_percentage"],  # ins_percentage
                    d["introduction_phone"],  # introduction_phone
                    d["introduction_text"],  # introduction_text
                    d["invoice_code"],  # invoice_code
                    d["new_invoice_code"],  # new_invoice_code
                    d["note"],  # note
                    d["patient_id"],  # patient_id
                    d["patient_price"],  # patient_price
                    d["pay_payment_account_document_id"],  # pay_payment_account_document_id
                    d["pay_payment_account_id"],  # pay_payment_account_id
                    d["pay_receipt_id"],  # pay_receipt_id
                    d["pay_return"],  # pay_return
                    d["phai_thu"],  # phai_thu
                    d["phai_tra"],  # phai_tra
                    d["print_notice_payment"],  # print_notice_payment
                    d["real_ins_payment"],  # real_ins_payment
                    d["real_ser_payment"],  # real_ser_payment
                    d["refund_invoice_code"],  # refund_invoice_code
                    d["returned"],  # returned
                    d["so_hd"],  # so_hd
                    d["so_luu_tru"],  # so_luu_tru
                    d["surgery_type"],  # surgery_type
                    d["ticket_id"],  # ticket_id
                    d["total_foresee"],  # total_foresee
                    d["total_ins_paid_price"],  # total_ins_paid_price
                    d["total_ins_price"],  # total_ins_price
                    d["total_price"],  # total_price
                    d["total_real_payment"],  # total_real_payment
                    d["total_ser_price"],  # total_ser_price
                    d["treatment_id"]  # treatment_id
                )

                # Thực thi câu lệnh INSERT
                cur.execute(insert_query, data)

                # Commit thay đổi
                conn.commit()

                print("Dữ liệu hóa đơn nội trú đã được chèn thành công!")
            except Exception as e:
                print("Lỗi nội trú "+ e)
    # Đóng cursor và kết nối
    cur.close()
    conn.close()

def fetch_data_from_api(header):
    global urlNgoaiTru, urlNoiTru, urlThuoc
    config = load_config()
    page_value = config['page_value']
    record_value = config['record_value']
    while(True):
        try:
            conn_params = sour.ConnectStr
            conn = psycopg2.connect(**conn_params)
            cur = conn.cursor()
            queryStr = f"SELECT * FROM patient order by stt asc OFFSET {record_value} LIMIT 20;"
            cur.execute(queryStr)
            listdata = cur.fetchall()
            if len(listdata) > 0:
                for item in listdata:
                    patient_id = item[1]
                    payload = {
                                "patient_id" : patient_id
                            }
                    fullUrlThuoc = urlThuoc + str(patient_id)
                    fullUrlNoiTru = urlNoiTru + str(patient_id)
                    responseNgoaiTru = requests.get(urlNgoaiTru, headers=header, json=payload)
                    if responseNgoaiTru.status_code == 200:
                        dataNG = responseNgoaiTru.json()
                        try:
                            data = dataNG['data']
                            if len(data) > 0:
                                add_ngoaitru_to_database(data['invoices'], patient_id)
                        except Exception as e:
                            print(f"Lỗi khi thêm dữ liệu vào database...")
                    else:
                        print(f"Lỗi khi lấy dữ liệu hóa đơn ngoại trú... "+ str(e))       

                    responseThuoc = requests.get(fullUrlThuoc, headers=header)
                    if responseThuoc.status_code == 200:
                        datat = responseThuoc.json()
                        try:
                            add_thuoc_to_database(datat['data'])
                        except Exception as e:
                            print(f"Lỗi khi thêm dữ liệu vào database... "  + str(e))
                    else:
                        print(f"Lỗi khi lấy dữ liệu thuốc...")    
                    
                    responseNoiTru = requests.get(fullUrlNoiTru, headers=header)
                    if responseNoiTru.status_code == 200:
                        datan = responseNoiTru.json()
                        try:
                            add_noitru_to_database(datan['data'])
                        except Exception as e:
                            print(f"Lỗi khi thêm dữ liệu vào database... "+ str(e))
                    else:
                        print(f"Lỗi khi lấy dữ liệu nội trú...")    
                p = page_value + 1
                pSub = p - 1
                rc = pSub * 20 - 1
                update_file_json(l4_value=p, l6_value=rc)
                page_value = p
                record_value = rc
        except Exception as e:
             print("Lỗi xảy ra trong quá trình truy cập CSDL... : "+ str(e))
        finally:
            if conn:
                cur.close()
                conn.close()

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
        fetch_data_from_api(headers)
        messagebox.showinfo(title="Thành công!",message="Hoàn thành cào dữ liệu bệnh nhân! Kết nối PostgreSQL đã đóng!.....")
    finally:
        loading = False
        loading_thread.join()
    
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
    l4.place(x=150, y=70)

    l5 = customtkinter.CTkLabel(master=frame, text="Số record: ", font=('Tahoma', 14))
    l5.place(x=20, y=100)
    l6 = customtkinter.CTkLabel(master=frame, text="0", font=('Tahoma', 14))
    l6.place(x=150, y=100)

    update_labels()

    run_button = customtkinter.CTkButton(master=frame, command=start_script_thread,text="Thực thi", font=('Tahoma', 13), fg_color="#005369", hover_color="#008097")
    run_button.place(x=160, y=200)

    app.mainloop()