
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import threading
import time

# Global değişkenler ve Lock
email_counter = 1
success_count = 0
failure_count = 0
counter_lock = threading.Lock()
stats_lock = threading.Lock()

# Tek bir e-posta gönderimi
def send_email(smtp_server, port, sender_email, sender_password, email, subject, body):
    global email_counter, success_count, failure_count
    try:
        # SMTP bağlantısını aç ve tek bir e-posta gönder
        with smtplib.SMTP(smtp_server, port) as server:
            server.starttls()  # Güvenli bağlantı
            server.login(sender_email, sender_password)

            msg = MIMEMultipart()
            msg['From'] = sender_email
            msg['To'] = email
            msg['Subject'] = subject
            msg.attach(MIMEText(body, 'html', _charset='utf-8'))

            server.sendmail(sender_email, email, msg.as_string())

            # Thread-safe sayaç artırma ve başarı yazdırma
            with counter_lock:
                print(f"{email_counter}. E-posta gönderildi: {email}")
                email_counter += 1

            # Başarı sayısını artır
            with stats_lock:
                success_count += 1
            return True  # Başarılı
    except Exception as e:
        print(f"E-posta gönderim hatası ({email}): {e}")

        # Başarısızlık sayısını artır
        with stats_lock:
            failure_count += 1
        return False  # Başarısız

# Thread'de e-posta gönder
def threaded_email_send(emails, smtp_server, port, sender_email, sender_password, subject, body):
    for email in emails:
        send_email(smtp_server, port, sender_email, sender_password, email, subject, body)

# Çoklu thread kullanarak e-postaları gönder
def send_emails_with_threads(file_path, smtp_server, port, sender_email, sender_password, subject, body, start_index, end_index, thread_count=4):
    global success_count, failure_count
    success_count = 0
    failure_count = 0

    data = pd.read_excel(file_path)
    email_list = data.iloc[start_index - 1:end_index, column].dropna().tolist()

    # E-postaları thread_count kadar parçaya böl
    email_chunks = [email_list[i::thread_count] for i in range(thread_count)]
    threads = []

    for chunk in email_chunks:
        thread = threading.Thread(
            target=threaded_email_send,
            args=(chunk, smtp_server, port, sender_email, sender_password, subject, body)
        )
        threads.append(thread)
        thread.start()

    # Tüm thread'leri bekle
    for thread in threads:
        thread.join()

    # İstatistik hesaplama
    total_emails = success_count + failure_count
    success_rate = (success_count / total_emails) * 100 if total_emails > 0 else 0
    failure_rate = (failure_count / total_emails) * 100 if total_emails > 0 else 0

    return success_count, failure_count, success_rate, failure_rate

# Program başlangıcı
start_time = time.time()

# ---------------------------------------------------------------------------------------------------------------------------------------------------

#  PARAMETRELER

excel_file = "mails.xlsx"            # mail dosyanizin adini girin
smtp_server = "smtp.yandex.com"      # "smtp.yandex.com" veya "smtp.gmail.com" veya istediginiz mail serveri 
port = 587
sender_email = " "                   # mail gondereceğiniz mail hesabi 
sender_password = ""                 # mail sifreniz 
subject = " HEADER HERE "            # gondereceginiz mailin konusu
body_file = "body.txt"               # mailin icerigini iceren dosyanin adi (line 29)
body = open(body_file, 'r', encoding='utf-8').read()
column = 8                           # E-posta adreslerinin bulunduğu sütun (Excel'de A=0, B=1, C=2, ...)

start_index = 1000                   # su satirdan 
end_index = 1250                     # su satira kadar olanlara mail gondericektir
thread_count = 4                     # es zamanli calistirmak icin thead sayisi onerilen 4 

# ---------------------------------------------------------------------------------------------------------------------------------------------------

success_count, failure_count, success_rate, failure_rate = send_emails_with_threads(
    excel_file, smtp_server, port, sender_email, sender_password, subject, body, start_index, end_index, thread_count
)

# Program bitiş zamanı
end_time = time.time()
total_time = end_time - start_time

# Çalışma süresini ve istatistikleri yazdır
minutes = int(total_time // 60)
seconds = int(total_time % 60)

print("\n--- İstatistikler ---")
print(f"Başarıyla gönderilen e-postalar: {success_count} ({success_rate:.2f}%)")
print(f"Başarısız e-postalar: {failure_count} ({failure_rate:.2f}%)")
print(f"Program çalışma süresi: {minutes} dk {seconds} sn")
