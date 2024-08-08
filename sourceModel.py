def CalculateModel(headers):
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

class CareerModel:
    career_id = 0
    code = ""
    disable = 0
    en_name = ""
    vi_name = ""
    ma_nghenghiep_bhyt = ""

class CountryModel:
    country_id = 0
    disable = 0
    en_name = ""
    vi_name = ""
    
class NationalityModel:
    nationality_id = 0
    disable = 0
    en_name = ""
    vi_name = ""
    ma_quoc_tich_bhyt = ""

class ProvinceModel:
    province_id = 0
    code = ""
    disabled = 0
    en_name = ""
    vi_name = ""
    ma_tinh = ""

class DistrictModel:
    district_id = 0
    province_id = 0
    ma_quan = ""
    code = ""
    disabled = 0
    en_name = ""
    vi_name = ""
    ma_quan_bhyt = ""

class WardModel:
    ward_id = 0
    district_id = 0
    disabled = 0
    en_name = ""
    vi_name = ""
    ma_phuong = ""
    ma_phuong_bhyt = ""
    auto_suggest_code = ""

class InvoiceModel:
    amount = 0
    bhyt = 0
    cashier_name = ""
    created = None
    created_time = None
    deleted = 0
    enum_examination_type = 0
    enum_invoice_type = 0
    enum_patient_type = 0
    enum_payment_type = 0
    icd10 = ""
    ins_paid_price = 0
    invoice_code = ""
    new_invoice_code = ""
    note = ""
    patient_price = 0
    pay_payment_account_document_id = 0
    pay_payment_account_id = 0
    pay_receipt_id = 242104738
    pay_return = 0
    refund_invoice_code = ""
    returned = 0
    so_hd = ""
    so_luu_tru = ""
    ticket_id = 0
    total_price = 0
    total_returned_amount = 0
    treatment_id = 0

class InvoiceDetailModel:
    amount= 0
    audi_last_modified_time= None
    audi_last_modified_user_id= 0
    bank_transaction_log_id= 0
    bhbl_url= ""
    branch_id= 0
    cashier_id= 0
    cashier_name= ""
    change_note= ""
    code= ""
    company_id= 0
    company_name= ""
    counter_id= 0
    created_time= None
    da_thu= ""
    deleted= 0
    discount_amount= 0
    discount_enum_unit= 0
    doctor_name= ""
    e_invoice_id= 0
    e_invoice_status= 0
    enum_invoice_type= 0
    enum_item_type= 0
    enum_payment_type= 0
    enum_re_exam_type= 0
    hd_date= ""
    hd_info= ""
    icd10= "0"
    ins_paid_price= 0
    ins_send_status= 0
    ins_transaction_code= None
    insurance_name= ""
    invoice_code= ""
    is_inventory= 0
    is_ngoai_quy_ds= 0
    item_type= ""
    new_invoice_code= ""
    new_pay_receipt_id= 0
    ngay_thu= None
    nguoi_thu= 0
    note= ""
    partner_invoice_code= None
    patient_id= 0
    patient_price= 0
    pay_payment_account_id= 0
    pay_payment_item_id= 0
    pay_receipt_id= 0
    percent_fee= 0
    phai_thu= 0
    phai_tra= 0
    quantity= 0
    reason_decline= ""
    refund_amount= 0
    refund_date= None
    refund_invoice_code= ""
    refund_pay_receipt_id= 0
    register_date= None
    retail_patient_address= ""
    retail_patient_name= ""
    returned= 0
    service_id= 0
    service_name= ""
    shift= 0
    so_hd= ""
    so_luu_tru= ""
    tam_ung= 0
    time_decline= None
    tochuc_quy= 0
    total= 0
    total_dichvu= 0
    total_nutrition= 0
    total_price= 0
    total_returned_amount= 0
    transfer_date= "0000-00-00"
    unit= ""
    unit_price= 0
    user_id_decline= 0
    zone= ""

class PrescriptionModel:
    backup_medical_record_id= 0
    canh_bao_hoat_chat= 0
    canh_bao_hoat_chat_dac_biet= 0
    canh_bao_qua_lieu= 0
    canh_bao_thang_tuoi_chi_dinh= 0
    check_in_out_hospital_record_id= 0
    chi_tiet_icd_canh_bao= None
    clinical_mani= ""
    code= ""
    create_date= None
    created_time= None
    da_cap= 0
    date= ""
    date_expired= None
    date_issued= None
    day_num= 0
    diagnosis= ""
    doctor_id= 0
    doctor_name= ""
    donthuocquocgia_checksum= ""
    donthuocquocgia_error= ""
    donthuocquocgia_modified_person_id= 0
    donthuocquocgia_modified_time= None
    dst_invoice_id= None
    e_invoice_id= None
    e_invoice_status= 0
    enum_customer_type= None
    fourth_icd10= None
    fourth_icd10_code= ""
    ghi_chu_toa_h= None
    giam_sat_note= None
    giam_sat_time= None
    giam_sat_user_id= 0
    icd10= ""
    icd10_note= None
    in_toa_thuoc= 0
    input_date= None
    insurance_code= None
    insurance_price= None
    invoice_code= None
    is_confirm_app= 0
    is_service= 0
    khth_duyet= 0
    khth_tg_duyet= None
    khth_use_id= 0
    locked= 0
    medical_record_id= 0
    note= ""
    patient_height= 0
    patient_id= 0
    patient_temp= 0
    patient_weight= 0
    percent= None
    prescription_id= 0
    prescription_type= 0
    primary_icd10= ""
    primary_icd10_code= ""
    processed= 0
    queue_code= ""
    re_examine= ""
    reason_decline= ""
    reexam_date= None
    second_icd10= None
    second_icd10_code= ""
    shift_id= 0
    sign_date= None
    signature_url= ""
    status= 0
    third_icd10= None
    third_icd10_code= ""
    ticket_id= 0
    time_decline= None
    total_price= 0
    trang_thai_giam_sat_toa_id= 0
    tuong_tac_thuoc= None
    updated_at= None
    user_id_decline= 0
    zone_code= ""

class PrescriptionDetailModel:
    afternoon= ""
    allow_auto_cal= 0
    bhbl_amount= 0
    bhbl_must_buy_full= 0
    bhbl_percent= 0
    bhxh_id= 0
    bhxh_pay_percent= 0
    bhyt_effect_date= ""
    bhyt_exp_effect_date= ""
    bhyt_ham_luong= ""
    bhyt_loai_thau= ""
    bhyt_loai_thuoc= ""
    bhyt_nha_thau= ""
    bhyt_nha_thau_bak= ""
    bhyt_nha_thau_code= None
    bhyt_nha_thau_id= None
    bhyt_pay_percent= 0
    bhyt_quyet_dinh= ""
    bhyt_so_dk_gpnk= ""
    bhyt_so_luong= 0
    bhyt_store= 0
    bk_enum_item_type= 0
    buoi_uong= ""
    bv_ap_thau= ""
    cancer_drug= 0
    canh_bao_thang_tuoi_chi_dinh= 0
    chi_dinh= ""
    chong_chi_dinh= ""
    co_han_dung= 0
    code= ""
    code_atc= ""
    code_insurance= ""
    confirm_sell_num= None
    country_name= ""
    created_at= None
    creator_id= 0
    da_cap= 0
    dang_bao_che= ""
    default_usage_id= 0
    disable= 0
    dong_goi= ""
    dosage= ""
    dosage_title= 0
    dose_quantity= 0
    dose_unit= 0
    drug_material_id= 0
    drug_original_name_id= 0
    dst_price= 0
    duong_dung_ax= ""
    enum_insurance_type= 0
    enum_item_type= 0
    enum_unit_import_sell= 0
    enum_usage= 0
    evening= ""
    gia_temp= 0
    goi_thau_bhyt= None
    ham_luong= ""
    hoat_chat_ax= ""
    im_price= 0
    include_children= 0
    insurance_company_id= 0
    insurance_drug_material_id= 0
    insurance_name= ""
    insurance_support= 0
    is_bhbl= 0
    is_bhyt= 0
    is_bhyt_in_surgery= 0
    is_control= 0
    is_deleted= 0
    is_duply_original_name= None
    is_inventory= 0
    is_max_one_day= 0
    is_max_one_day_by_weight= 0
    is_max_one_times= 0
    is_max_one_times_by_weight= 0
    is_min_one_day_by_weight= 0
    is_min_one_times_by_weight= 0
    is_special_dept= 0
    is_used_event= 0
    is_used_event_idm= 0
    khong_thanh_toan_rieng= 0
    khu_dieu_tri= 0
    latest_import_price= 0
    latest_import_price_vat= 0
    lieu_luong= ""
    loai_ke_toa= 0
    loai_thuan_hop= ""
    locked= 0
    ma_duong_dung_ax= ""
    ma_hoat_chat_ax= ""
    ma_thuoc_dqg= None
    made_in= 0
    manufacturer_id= 0
    max_one_day= 0
    max_one_day_by_weight= 0
    max_one_times= 0
    max_one_times_by_weight= 0
    max_usage= 0
    medicine_id= 0
    min_one_day_by_weight= 0
    min_one_times_by_weight= 0
    modifier_id= 0
    morning= ""
    ngay_dung_thuoc= 0
    ngay_hieu_luc_hop_dong= ""
    nhom_duoc_ly= 0
    nhom_thuoc= ""
    noon= ""
    note= ""
    num_of_day= 0
    num_of_time= 0
    num_per_time= ""
    order_by= 0
    original_names= ""
    paid= 0
    phan_nhom_bhyt= ""
    phan_nhom_thuoc_id= 0
    pharmacology_id= 0
    poison_type_id= 0
    prescription_id= 0
    prescription_item_id= 0
    price= 0
    price_bv= 0
    price_qd= 0
    proprietary_name= ""
    quantity_num= 0
    quantity_remain= 0
    quantity_title= 0
    quantity_use= 0
    renumeration_price= 0
    service_group_cost_code= 0
    so_dk_gpnk= ""
    so_luong_cho_nhap= 0
    so_luong_da_nhan= 0
    solan_ngay= 0
    status= 0
    stt_dmt= ""
    stt_tt= ""
    t_trantt= 0
    tac_dung= ""
    tac_dung_phu= ""
    ten_hang_sx= None
    ten_theo_thau= None
    ten_thuongmai= None
    thang_tuoi_chi_dinh= 0
    thoi_gian_bao_quan= 0
    thuoc_ra_le= 0
    time= 0
    unit_usage_id= 0
    unit_volume_id= 0
    updated_at= None
    usage_num= 0
    usage_title= 0
    volume_value= 0
    warning_note_doctor= ""

class ServiceDetailModel:
    bodyPartName= None
    body_part_id= None
    cancel= 0
    chi_tiet_chi_dinh= None
    code= ""
    create_date= None
    doctor_id= 0
    doctorname= ""
    done= 0
    enum_examination_type= 0
    enum_unit= 0
    examination_type_id= 0
    insurance_type= 0
    is_bhyt= 0
    item_id= 0
    item_type= ""
    lab_type= None
    lab_type_sub= None
    loai_mau_code= None
    loai_mau_id= 0
    nhomMauId= None
    nhom_loai_mau_code= None
    normal_price= 0
    note= ""
    paid= 0
    paid_ins= 0
    patient_id= 0
    printed= 0
    quantity= 0
    queue_code= ""
    s_name= ""
    service_id= 0
    status= ""
    tenLoaiMau= None
    ticket_id= 0
    ticket_item_id= 0
    tinh_trang_thuc_hien= 0
    tk_name= ""
    trai_phai= ""
    urlResultPacs= ""
    zone_code= ""

class InvoiceInPatientModel:
    allow_re_print_hos_con= 0
    amount= 0
    bhbl_url= ""
    bhyt= 0
    cashier_name= ""
    check_in_date= None
    check_in_out_hospital_record_id= 0
    check_out_date= None
    closed= 0
    created= None
    created_time= None
    deleted= 0
    discount_amount= 0
    end_department= ""
    end_department_id= 0
    enum_account_type= 0
    enum_examination_type= None
    enum_invoice_type= 0
    enum_patient_type= 0
    enum_payment_type= 0
    first_department= ""
    first_department_id= 0
    icd10= ""
    icd10_id= 0
    ins_paid_price= 0
    ins_percentage= 0
    introduction_phone= ""
    introduction_text= ""
    invoice_code= ""
    new_invoice_code= ""
    note= ""
    patient_id= 0
    patient_price= 0
    pay_payment_account_document_id= 0
    pay_payment_account_id= 0
    pay_receipt_id= 0
    pay_return= 0
    phai_thu= 0
    phai_tra= 0
    print_notice_payment= 0
    real_ins_payment= 0
    real_ser_payment= 0
    refund_invoice_code= ""
    returned= 0
    so_hd= ""
    so_luu_tru= ""
    surgery_type= 0
    ticket_id= None
    total_foresee= 0
    total_ins_paid_price= 0
    total_ins_price= 0
    total_price= 0
    total_real_payment= 0
    total_ser_price= 0
    treatment_id= 0

class InvoiceInPatientDetailModel:
    code= ""
    department_id= 0
    department_name= ""
    doctor_name= None
    enum_item_type= 0
    insurance_name= ""
    item_code= ""
    item_type= ""
    quantity= 0
    returned= 0
    service_name= ""
    total= 0
    total_price= 0
    unit= ""
    unit_price= 0

class PatientModel:
    stt = 0
    person_id = 0
    patient_code = ""
    patient_code_2 = None
    vaccination_code = None
    name = ""
    backup_name = None
    gender = ""
    date_of_birth = ""
    phone_number = None
    full_address = ""
    career_vi_name = None
    career_en_name = None
    ethnic_vi_name = ""
    vi_nationality = ""
    en_nationality = ""
    blood_group = None
    blood_rh = None
    blood_result_time = None
    qr_code_bhyt = ""
    qr_code_cccd_chip = ""
    created_date = ""
    last_exam = ""
    father_name = ""
    father_phone = ""
    mother_name = ""
    mother_phone = ""

    def to_dict(self):
        return {
            "STT": self.stt,
            "ID bệnh nhân": self.person_id,
            "Mã bệnh nhân": self.patient_code,
            "Mã bệnh nhân 2": self.patient_code_2,
            "Mã tiêm chủng": self.vaccination_code,
            "Tên bệnh nhân": self.name,
            "Tên khác": self.backup_name,
            "Giới tính": self.gender,
            "Ngày sinh": self.date_of_birth,
            "Số điện thoại": self.phone_number,
            "Địa chỉ": self.full_address,
            "Nghề nghiệp Tiếng Việt": self.career_vi_name,
            "Nghề nghiệp Tiếng Anh": self.career_en_name,
            "Dân tộc": self.ethnic_vi_name,
            "Quốc tịch Tiếng Việt": self.vi_nationality,
            "Quốc tịch Tiếng Anh": self.en_nationality,
            "Nhóm máu": self.blood_group,
            "RH máu": self.blood_rh,
            "Kết quả máu": self.blood_result_time,
            "Mã BHYT": self.qr_code_bhyt,
            "Mã CCCD": self.qr_code_cccd_chip,
            "Thời gian tạo": self.created_date,
            "Lần khám gần nhất": self.last_exam,
            "Họ tên cha": self.father_name,
            "SĐT cha": self.father_phone,
            "Họ tên mẹ": self.mother_name,
            "SĐT mẹ": self.mother_phone
        }
    
    def ExportModel(self):
        headers = [
            ("STT" , self.stt),
            ("ID bệnh nhân", self.person_id),
            ("Mã bệnh nhân", self.patient_code),
            ("Mã bệnh nhân 2", self.patient_code_2),
            ("Mã tiêm chủng", self.vaccination_code),
            ("Tên bệnh nhân", self.name),
            ("Tên khác", self.backup_name),
            ("Giới tính", self.gender),
            ("Ngày sinh", self.date_of_birth),
            ("Số điện thoại", self.phone_number),
            ("Địa chỉ", self.full_address),
            ("Nghề nghiệp Tiếng Việt", self.career_vi_name),
            ("Nghề nghiệp Tiếng Anh", self.career_en_name),
            ("Dân tộc", self.ethnic_vi_name),
            ("Quốc tịch Tiếng Việt", self.vi_nationality),
            ("Quốc tịch Tiếng Anh", self.en_nationality),
            ("Nhóm máu", self.blood_group),
            ("RH máu", self.blood_rh),
            ("Kết quả máu", self.blood_result_time),
            ("Mã BHYT", self.qr_code_bhyt),
            ("Mã CCCD", self.qr_code_cccd_chip),
            ("Thời gian tạo", self.created_date),
            ("Lần khám gần nhất", self.last_exam),
            ("Họ tên cha", self.father_name),
            ("SĐT cha", self.father_phone),
            ("Họ tên mẹ", self.mother_name),
            ("SĐT mẹ", self.mother_phone)
        ]
        s = CalculateModel(headers=headers)
        return s
    