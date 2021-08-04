from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from string import Template
from pathlib import Path
import csv
import eel
import os
from cosmos import query_authorization_code,used_authorization_code,query_used_authorization_code,save_frequency
import wmi
import time
from email.mime.application import MIMEApplication


c = wmi.WMI()

@eel.expose
def main(letter_title,html_path,csv_path,sender_mail,sender_password,mail_csv_order,name_csv_order,mail_content1_order,mail_content2_order,attached_dir_path,attached_csv_order):

    try:
        # 開啟 CSV 檔案
        all_content = []
        with open(csv_path, newline='') as csvfile:

          # 讀取 CSV 檔案內容
          rows = csv.reader(csvfile)

          # 以迴圈輸出每一列，加入<br>在信中才能換行
          for row in rows:
            row[2] = row[2].replace("\n","<br>")
            all_content.append(row)

    except Exception as e:
        print(e)
        return "CSV路徑有誤，或是內容有缺少"

    # 去除最上面那行
    all_content = all_content[1:]

    try:
        for item in all_content:
            body_dic = {}

            body_dic["name"] = item[int(name_csv_order)]

            # 輸入欄位如果不是NO就新增到客製化欄位中
            if mail_content1_order !="NO":
                body_dic["content1"] = item[int(mail_content1_order)]
            if mail_content2_order !="NO":
                body_dic["content2"] = item[int(mail_content2_order)]

            content = MIMEMultipart()  # 建立MIMEMultipart物件
            content["subject"] = letter_title  # 郵件標題
            content["from"] = sender_mail  # 寄件者
            content["to"] = item[int(mail_csv_order)] # 收件者姓名
            # content.attach(MIMEText("大塚雲端課程"))  # 郵件內容

            # 如果指定路徑下有附件就加到信件中
            if os.path.isfile(attached_dir_path+"/"+item[int(attached_csv_order)]):
                # 添加附件
                part_attach1 = MIMEApplication(open(attached_dir_path+"/"+item[int(attached_csv_order)], 'rb').read())  # 開啟附件
                part_attach1.add_header('Content-Disposition', 'attachment', filename=item[int(attached_csv_order)])  # 為附件命名
                content.attach(part_attach1)

            template = Template(Path(html_path).read_text())
            body = template.substitute(body_dic) # 帶入客製化內容
            content.attach(MIMEText(body, "html"))  # HTML郵件內容

            import smtplib
            with smtplib.SMTP(host="smtp.gmail.com", port="587") as smtp:  # 設定SMTP伺服器
                try:
                    smtp.ehlo()  # 驗證SMTP伺服器
                    smtp.starttls()  # 建立加密傳輸
                    smtp.login(sender_mail, sender_password)  # 登入寄件者gmail
                    smtp.send_message(content)  # 寄送郵件
                    print(" Complete! ")
                except Exception as e:
                    print(" Error message: ", e)
                    return "請檢查Gmail帳密，是否有誤"

    except Exception as e:
        return "輸入資料有誤，請檢查內容"

    # 將授權碼寫到key.txt中
    with open(os.getcwd()+'\web'+"\key.txt", "r", encoding="utf-8") as f:
        key =  f.readlines()[0]

    # 將換行符號\n取代掉(授權碼機制會在txt存入金鑰及到期日共兩行，在換行時會有\n，所以要取代掉)
    key = key.replace("\n", "")

    user_name = query_authorization_code(key)[0]['company_name']
    number = len(all_content)
    save_frequency("寄送{}筆".format(number),user_name)

    return "信件已全部寄出"

# 第一次授權要做的處理
@eel.expose
def comfirm_authorization(Send_email_authorization):
    authorization_result = query_authorization_code(Send_email_authorization)
    # 長度就大於1，代表在授權碼資料庫中找到資料
    if len(authorization_result)>0:
        # 查看授權是否被使用
        if len(query_used_authorization_code(Send_email_authorization))>0:
            return "授權已使用，如有問題請聯繫大塚資訊"
        else:

            # 取得現在的時間(時間截記)
            now_time = int(time.time())

            # 將授權天數轉換成秒數(時間截記)+並加上現在的時間，代表到期日的時截記
            end_day_ts = int(authorization_result[0]["authorization_day"])*86400+now_time

            # 將時間截記轉換成日期
            end_day = time.localtime(end_day_ts)
            end_day = time.strftime("%Y-%m-%d", end_day)

            # 將授權寫到雲端資料庫。並紀錄公司名稱、授權到期日
            used_authorization_code(Send_email_authorization,authorization_result[0]["company_name"],end_day,end_day_ts)

            # 第一次輸入授權碼，儲存到本地端txt檔案中
            with open(os.getcwd()+'\web'+"\key.txt", "w", encoding="utf-8") as f:
                f.write(Send_email_authorization+"\n"+str(end_day_ts))

            # 授權成功
            return "success"
    else:
        # 無此授權碼，授權失敗
        return "failure"


# 開啟應用時檢查授權，如果txt檔案有東西，就代表之前曾經設定過授權，直接進到畫面
@eel.expose
def have_use_authorization():
    # 打開指定路徑檔案查看是否有金鑰紀錄
    with open(os.getcwd()+'\web'+"\key.txt", "r", encoding="utf-8") as f:
        key =  f.readlines()[0]

        # 將換行符號\n取代掉
        key = key.replace("\n", "")

        # 本地CPU序號
        local_cpu_serial_number = c.Win32_BaseBoard()[0].SerialNumber.strip()

        if len(query_authorization_code(key))>0:
            cosmes_cpu_serial_number = query_used_authorization_code(key)[0]["cpu_serial_number"]
            # 本地已有授權，且CPU序號符合，直接帶入使用畫面
            if cosmes_cpu_serial_number == local_cpu_serial_number:
                return "success"
            # CPU序號不符合，授權失敗
            else:
                return "failure"
        else:
            # 本地端還沒有授權過，或是授權失敗
            return "failure"


eel.init(os.getcwd()+'\web') # eel.init(網頁的資料夾)
eel.start(os.getcwd()+'\web'+r"\authorization.html",size = (734,580),port=5020) # eel.start(html名稱, size=(起始大小))