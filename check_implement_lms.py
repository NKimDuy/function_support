from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service
import time
import requests
from datetime import datetime
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import json
import re
import pandas as pd

#------------------------------
# khớp dữ liệu lấy từ lsa
#------------------------------
def get_dictionary(txt):
      lines = txt.splitlines()
      subject = ""
      group = ""
      techer_id = ""
      teacher_name = ""
      for i, line in enumerate(lines):
            if line.startswith("[242]"):
                  match = re.search(r"\[242\]\s+(\w+)\s+-.*\((\w+)-(\w+)\)", line)
                  if match:
                        subject = match.group(1)     # FINA2343
                        techer_id = match.group(2)      # KT196
                        group = match.group(3)       # TN303
            elif line.startswith("Giảng viên"):
                  teacher_name = line.split(". ", 1)[-1].strip()
      return [group, subject, techer_id, teacher_name]


#------------------------------
# lấy dữ liệu từ file excel theo từng học kỳ
# row[0]: nhóm
# row[1]: mã môn học
# row[2]: mã giảng viên
# row[3]: tên giảng viên
# row[4]: tên khoa
#------------------------------
def get_detail(file):
      list_detail = {}
      wb = load_workbook(file)
      ws = wb.active

      for row in ws.iter_rows(values_only=True):
            key = row[0] + "-" + row[1]
            list_detail[key] = [row[2], row[3], row[4]]
      wb.close()
      return list_detail


#------------------------------
# vào trang web lsa và lấy dữ liệu giảng viên đã soạn bài về 
#------------------------------
def get_lsa(semester, url):
      value = "Đăng nhập bằng HCMCOU-SSO"
      semester = " ".join(["[LIVE] LMS TX", semester])
      general = "http://lsa.ou.edu.vn/vi/admin/mm/report/usersiteoverviews"
      xpath = f"//button[text()='{value}']"
      chrome_options = Options()
      #chrome_options.add_argument("--ignore-certificate-errors")
      #chrome_options.add_argument("--disable-features=StrictTransportSecurity")
      chrome_options.add_argument("--headless")
      #chrome_options.add_argument("--allow-insecure-localhost")  # Nếu là localhost
      #driver = webdriver.Chrome(options=chrome_options)
      driver = webdriver.Chrome()
      #driver.get("http://lsa.ou.edu.vn")
      driver.get(url)

      try:
            button_semester = WebDriverWait(driver, 15).until(
                  EC.element_to_be_clickable((By.XPATH, xpath))
            )
            button_semester.click()
      except:
            print("Không tìm thấy nút đăng nhập bằng HCMCOU-SSO")

      try:
            dropdown = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "form-usertype"))
            )
            select_type_user = Select(dropdown)
            select_type_user.select_by_visible_text("Cán bộ-Nhân viên / Giảng viên")
      except:
            print("Không tìm thấy nút Cán bộ nhân viên/Giảng viên")

      try:
            username = WebDriverWait(driver, 15).until(
                  EC.presence_of_element_located((By.ID, "form-username"))
            )
            username.send_keys("duy.nk")
      except:
            print("Không tìm thấy ô để nhập tài khoản")

      try:
            password = driver.find_element(By.ID, "form-password")
            password.send_keys("tonyTeo!998")
      except:
            print("Không tìm thấy ô để nhập mật khẩu")

      try:
            captcha = driver.find_element(By.ID, "form-captcha")
            captcha.send_keys("clcl")
      except:
            print("Không tìm thấy ô để nhập Capcha")

      try:
            button_login = WebDriverWait(driver, 15).until(
                  EC.element_to_be_clickable((By.XPATH, "//button[text()='Đăng nhập']"))
            )
            button_login.click()
      except:
            print("Không tìm thấy nút đăng nhập")

      try:
            has_found_button_allow = driver.find_elements(By.CSS_SELECTOR, ".btn.btn-success.btn-approve")
            if has_found_button_allow:
                  button_allow = WebDriverWait(driver, 15).until(
                  EC.presence_of_element_located((By.CSS_SELECTOR, ".btn btn-success btn-approve"))
                  )
                  button_allow.click()
      except:
            print("Không tìm thấy nút để nhấn đồng ý")

      try:
            dropdown_semester = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.ID, "moodlesiteid"))
            )

            select_type_semester = Select(dropdown_semester)
            select_type_semester.select_by_value("54")
      except:
            print("Không tìm thấy dropdownlist thể hiện học kỳ")

      try:
            driver.execute_script("arguments[0].style.display='block';", driver.find_element(By.ID, "menu_1_sub"))
            overview_link = WebDriverWait(driver, 20).until(
                  EC.element_to_be_clickable((By.CSS_SELECTOR, "a[href*='usersiteoverviews']"))
            )
            overview_link.click()
      except:
            print("Không tìm thấy nút report")

      try:
            table = WebDriverWait(driver, 40).until(
                  EC.presence_of_element_located((By.ID, "ourptlistcourse"))  # Thay "myTable" bằng ID thực tế
            )
      except:
            print("không tìm thấy bảng")

      try:
            rows = WebDriverWait(driver, 20).until(
                  EC.presence_of_all_elements_located((By.XPATH, ".//tr"))
            )
      except: 
            print("không tìm thấy các dòng")

      get_subject = []
      for row in rows:
            cells = row.find_elements(By.XPATH, ".//td")
            for cell in cells:
                  group, subject, teacher_id, teacher_name = get_dictionary(cell.text)
                  if group and subject:
                        get_subject.append(group + "-" + subject)
      return get_subject


# định dạng ngày có dạng yyyy-MM-DD
def get_subject_by_day(semester, from_day, to_day, file):
      url_link_unit = "https://api.ou.edu.vn/api/v1/hdmdp"
      url_list_subject_semester = "https://api.ou.edu.vn/api/v1/tkblopdp"
      headers = {
            "Authorization": "Bearer 52C4E470AF3AE6C56276FAE8666788291F7AEA1667FE67C9DF743FF49FD5C74B"
      }
      from_day = datetime.strptime(from_day, "%Y-%m-%d")
      to_day = datetime.strptime(to_day, "%Y-%m-%d")
      list_subject_in_range = {}

      get_list_unit = requests.get(url_link_unit, headers=headers)
      list_unit = get_list_unit.json()
      for unit in list_unit.get("data", []):
            params_list_subject_semester = {
                  "nhhk": semester,
                  "madp": unit["MaDP"]
            }
            get_list_subject_semester = requests.get(url_list_subject_semester, headers=headers, params=params_list_subject_semester)
            list_subject_semester = get_list_subject_semester.json()
            for lst in list_subject_semester.get("data", []):
                  if lst["TUNGAYTKB"] is not None:
                        if from_day <= datetime.strptime(lst["TUNGAYTKB"], "%Y-%m-%d") <= to_day:
                              key = lst["NhomTo"] + "-" + lst["MaMH"]
                              list_detail = get_detail(file)
                              if key not in list_subject_in_range:
                                    list_subject_in_range[key] =  [
                                          lst["NhomTo"], 
                                          lst["MaMH"],
                                          lst["TenMH"],
                                          lst["TUNGAYTKB"],
                                          lst["MaLop"],
                                          lst["TenLop"],
                                          lst["MaDP"],
                                          lst["TenDP"],
                                         list_detail[key][0], #mã giảng viên
                                         list_detail[key][1], #tên giảng viên
                                         list_detail[key][2] #khoa
                                    ]
                              else:
                                    list_subject_in_range[key][4] = ",".join([list_subject_in_range[key][4], lst["MaLop"]])     
      return list_subject_in_range


#------------------------------
# Tạo file báo cáo
#------------------------------
def create_file_report(data):
      # TODO: Định dạng cho các dòng trong file
      header_font = Font(name="Time New Roman", size=12, bold=True)
      header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
      header_fill = PatternFill(start_color="97FFFF", end_color="97FFFF", fill_type="solid")
      data_font = Font(name="Time New Roman", size=11)
      data_alignment = Alignment(horizontal="center", wrap_text=True, vertical="center")
      border_style = Border(
            left = Side(style="thin"),
            right = Side(style="thin"),
            top = Side(style="thin"),
            bottom = Side(style="thin")
      )

      title = [
            "STT",
            "Khoa phụ trách",
            "Mã địa phương",
            "Tên địa phương",
            "Mã môn học",
            "Tên môn học",
            "Mã nhóm",
            "Mã lớp",
            "Tên lớp",
            "Mã giảng viên",
            "Tên giảng viên",
            "Ngày bắt đầu",
            "Đã soạn LMS"
      ]

      wb = Workbook()
      sheet_detail = wb.active 
      sheet_detail.title = "Chi tiết" # sheet tên chi tiết
      # TODO: thiết lập giá trị và định dạng cho dòng tiêu đề
      for title_idx, row_title in enumerate(title, start = 1):
            sheet_detail.cell(row = 1, column = title_idx).value = row_title
            sheet_detail.cell(row = 1, column = title_idx).font = header_font
            sheet_detail.cell(row = 1, column = title_idx).alignment = header_alignment
            sheet_detail.cell(row = 1, column = title_idx).fill = header_fill
            sheet_detail.cell(row = 1, column = title_idx).border = border_style
      
      # TODO: thiết lập giá trị và định dạng cho các dòng còn lại
      for row_idx, row_data in enumerate(data, start = 2):
            has_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
            if row_data["has_lms"] != "x":
                  has_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
            sheet_detail.cell(row = row_idx, column = 1).value = row_idx - 1
            sheet_detail.cell(row = row_idx, column = 1).font = data_font
            sheet_detail.cell(row = row_idx, column = 1).alignment = data_alignment
            sheet_detail.cell(row = row_idx, column = 1).border = border_style
            sheet_detail.cell(row = row_idx, column = 1).fill = has_fill

            sheet_detail.cell(row = row_idx, column = 2).value = row_data["department"]
            sheet_detail.cell(row = row_idx, column = 2).font = data_font
            sheet_detail.cell(row = row_idx, column = 2).alignment = data_alignment
            sheet_detail.cell(row = row_idx, column = 2).border = border_style
            sheet_detail.cell(row = row_idx, column = 2).fill = has_fill

            sheet_detail.cell(row = row_idx, column = 3).value = row_data["id_unit"]
            sheet_detail.cell(row = row_idx, column = 3).font = data_font
            sheet_detail.cell(row = row_idx, column = 3).alignment = data_alignment
            sheet_detail.cell(row = row_idx, column = 3).border = border_style
            sheet_detail.cell(row = row_idx, column = 3).fill = has_fill

            sheet_detail.cell(row = row_idx, column = 4).value = row_data["name_unit"]
            sheet_detail.cell(row = row_idx, column = 4).font = data_font
            sheet_detail.cell(row = row_idx, column = 4).alignment = data_alignment
            sheet_detail.cell(row = row_idx, column = 4).border = border_style
            sheet_detail.cell(row = row_idx, column = 4).fill = has_fill

            sheet_detail.cell(row = row_idx, column = 5).value = row_data["id_subject"]
            sheet_detail.cell(row = row_idx, column = 5).font = data_font
            sheet_detail.cell(row = row_idx, column = 5).alignment = data_alignment
            sheet_detail.cell(row = row_idx, column = 5).border = border_style
            sheet_detail.cell(row = row_idx, column = 5).fill = has_fill

            sheet_detail.cell(row = row_idx, column = 6).value = row_data["name_subject"]
            sheet_detail.cell(row = row_idx, column = 6).font = data_font
            sheet_detail.cell(row = row_idx, column = 6).alignment = data_alignment
            sheet_detail.cell(row = row_idx, column = 6).border = border_style
            sheet_detail.cell(row = row_idx, column = 6).fill = has_fill

            sheet_detail.cell(row = row_idx, column = 7).value = row_data["group"]
            sheet_detail.cell(row = row_idx, column = 7).font = data_font
            sheet_detail.cell(row = row_idx, column = 7).alignment = data_alignment
            sheet_detail.cell(row = row_idx, column = 7).border = border_style
            sheet_detail.cell(row = row_idx, column = 7).fill = has_fill

            sheet_detail.cell(row = row_idx, column = 8).value = row_data["id_class"]
            sheet_detail.cell(row = row_idx, column = 8).font = data_font
            sheet_detail.cell(row = row_idx, column = 8).alignment = data_alignment
            sheet_detail.cell(row = row_idx, column = 8).border = border_style
            sheet_detail.cell(row = row_idx, column = 8).fill = has_fill

            sheet_detail.cell(row = row_idx, column = 9).value = row_data["name_class"]
            sheet_detail.cell(row = row_idx, column = 9).font = data_font
            sheet_detail.cell(row = row_idx, column = 9).alignment = data_alignment
            sheet_detail.cell(row = row_idx, column = 9).border = border_style
            sheet_detail.cell(row = row_idx, column = 9).fill = has_fill

            sheet_detail.cell(row = row_idx, column = 10).value = row_data["id_teacher"]
            sheet_detail.cell(row = row_idx, column = 10).font = data_font
            sheet_detail.cell(row = row_idx, column = 10).alignment = data_alignment
            sheet_detail.cell(row = row_idx, column = 10).border = border_style
            sheet_detail.cell(row = row_idx, column = 10).fill = has_fill

            sheet_detail.cell(row = row_idx, column = 11).value = row_data["name_teacher"]
            sheet_detail.cell(row = row_idx, column = 11).font = data_font
            sheet_detail.cell(row = row_idx, column = 11).alignment = data_alignment
            sheet_detail.cell(row = row_idx, column = 11).border = border_style
            sheet_detail.cell(row = row_idx, column = 11).fill = has_fill

            sheet_detail.cell(row = row_idx, column = 12).value = row_data["from_day"]
            sheet_detail.cell(row = row_idx, column = 12).font = data_font
            sheet_detail.cell(row = row_idx, column = 12).alignment = data_alignment
            sheet_detail.cell(row = row_idx, column = 12).border = border_style
            sheet_detail.cell(row = row_idx, column = 12).fill = has_fill

            sheet_detail.cell(row = row_idx, column = 13).value = row_data["has_lms"]
            sheet_detail.cell(row = row_idx, column = 13).font = data_font
            sheet_detail.cell(row = row_idx, column = 13).alignment = data_alignment
            sheet_detail.cell(row = row_idx, column = 13).border = border_style
            sheet_detail.cell(row = row_idx, column = 13).fill = has_fill

      # TODO: điều chỉnh độ rộng của cột dựa trên giá trị dài nhất
      for col_idx in range(1, len(data[0]) + 1):
            max_length = 0
            column = get_column_letter(col_idx)
            for cell in sheet_detail[column]:
                  if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
            adjust_width = max_length + 2
            sheet_detail.column_dimensions[column].width = adjust_width

      # TODO: Tạo sheet thứ 2 nội dung về tổng quan
      # FIXME: gom nhóm dữ liệu để tạo file tổng quan

      wb.save("Tình hình soạn thảo LMS.xlsx")


def main():
      data = [
            {"group": "SG001", "id_subject": "ACCO4331", "name_subject": "Quản trị học", "from_day": "31/12/1994", "id_class": "TM123456", "name_class": "Lớp luật Đồng Tháp Mười", "id_unit": "TM", "name_unit": "Đồng Tháp Mười Long An", "id_teacher": "TX001", "name_teacher": "Nguyễn Kim Duy", "department": "Luật", "has_lms": "x"},
            {"group": "SG001", "id_subject": "ACCO4331", "name_subject": "Quản trị học", "from_day": "31/12/1994", "id_class": "TM123456", "name_class": "Lớp luật Đồng Tháp Mười", "id_unit": "TM", "name_unit": "Đồng Tháp Mười Long An", "id_teacher": "TX001", "name_teacher": "Nguyễn Kim Duy", "department": "Luật", "has_lms": "x"},
            {"group": "SG001", "id_subject": "ACCO4331", "name_subject": "Quản trị học", "from_day": "31/12/1994", "id_class": "TM123456", "name_class": "Lớp luật Đồng Tháp Mười", "id_unit": "TM", "name_unit": "Đồng Tháp Mười Long An", "id_teacher": "TX001", "name_teacher": "Nguyễn Kim Duy", "department": "Luật", "has_lms": ""},
            {"group": "SG001", "id_subject": "ACCO4331", "name_subject": "Quản trị học", "from_day": "31/12/1994", "id_class": "TM123456", "name_class": "Lớp luật Đồng Tháp Mười", "id_unit": "TM", "name_unit": "Đồng Tháp Mười Long An", "id_teacher": "TX001", "name_teacher": "Nguyễn Kim Duy", "department": "Xây dựng", "has_lms": "x"},
            {"group": "SG001", "id_subject": "ACCO4331", "name_subject": "Quản trị học", "from_day": "31/12/1994", "id_class": "TM123456", "name_class": "Lớp luật Đồng Tháp Mười", "id_unit": "TM", "name_unit": "Đồng Tháp Mười Long An", "id_teacher": "TX001", "name_teacher": "Nguyễn Kim Duy", "department": "Xây dựng", "has_lms": ""},
            {"group": "SG001", "id_subject": "ACCO4331", "name_subject": "Quản trị học", "from_day": "31/12/1994", "id_class": "TM123456", "name_class": "Lớp luật Đồng Tháp Mười", "id_unit": "TM", "name_unit": "Đồng Tháp Mười Long An", "id_teacher": "TX001", "name_teacher": "Nguyễn Kim Duy", "department": "Quản trị kinh doanh", "has_lms": "x"}
      ]

      create_file_report(data)
      # semester = "242"
      # from_day = "2025-04-21"
      # to_day = "2025-04-27"
      # url_lsa = "http://lsa.ou.edu.vn"
      # file = "242_detail.xlsx"
      # report_final = []
      # list_lsa = get_lsa(semester, url_lsa)
      # list_subject_by_day = get_subject_by_day(semester, from_day, to_day, file)

      # for key, value in list_subject_by_day.items():
      #       temp = {}
      #       temp["group"] = value[0]
      #       temp["id_subject"] = value[1]
      #       temp["name_subject"] = value[2]
      #       temp["from_day"] = value[3]
      #       temp["id_class"] = value[4]
      #       temp["name_class"] = value[5]
      #       temp["id_unit"] = value[6]
      #       temp["name_unit"] = value[7]
      #       temp["id_teacher"] = value[8]
      #       temp["name_teacher"] = value[9]
      #       temp["department"] = value[10]
      #       temp["has_lms"] = "x" if key in list_lsa else ""
      #       report_final.append(temp)

if __name__ == "__main__":
      main()