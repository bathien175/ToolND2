urlNgoaiTru = "http://192.168.0.77/api/patients/getInvoices_outpatient/"
Ngoaitrupayload = {
    "patient_id": 173554757
}
urlThuocNgoaiTru = "http://192.168.0.77/api/patients/getListPrescriptionOutpatient/173554757"
urlNoiTru = "http://192.168.0.77/api/patients/getLast_department/"
NoitruPayload = {
    "patient_id": 173554757
}
ChitiethoadonUrl = "http://192.168.0.77/api/patients/getOutPatientInvoiceDetail/"
ChitiethoadonPayload = {
    "invoice_code" : "2T403181",
    "patient_id" : 173554757,
    "pay_receipt_id" : 242104738
}
chitietdichvuurl = "http://192.168.0.77/api/doctor_outpatient/load_doc_lst_srv"
chitietdichvupayload = {
    "cancel <>": 1,
    "employee_code": "QUN12",
    "is_bhyt": 0,
    "lst_not_cls_csul_type": "'emergency','radiology','bed','drug','material','consultation_package','tdcn_lgtl','consultation'",
    "name": "Ngô Quang Quyền",
    "patientCode": "79408230156818",
    "patient_id": 173554757,
    "time_in_day": "2023-08-16",
    "username": "quyen.ngoq"
}
chitietthuocurl = "http://192.168.0.77/api/doctor_pkg/load_pkg_last_detail_presc"
chitietthuocpayload = {
    "prescription_id": 188765153
}