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
    