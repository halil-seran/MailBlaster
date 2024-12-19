import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Excel'den e-posta adreslerini oku
def get_email_list(file_path, start_index, end_index):
    data = pd.read_excel(file_path)  # Excel dosyasını oku
    # start_index ve end_index arasında e-posta adreslerini al
    email_list = data.iloc[start_index-1:end_index, 8].dropna().tolist()  # 8. sütun I'ye karşılık geliyor
    return email_list

# Tek bir e-posta gönderimi
def send_email(smtp_server, port, sender_email, sender_password, recipient_email, subject, body):
    try:
        # SMTP sunucusuna bağlan
        with smtplib.SMTP(smtp_server, port) as server:
            server.starttls()  # Güvenli bağlantı
            server.login(sender_email, sender_password)

            # E-posta oluştur
            msg = MIMEMultipart()
            msg['From'] = sender_email
            msg['To'] = recipient_email
            msg['Subject'] = subject
            msg.attach(MIMEText(body, 'plain'))  # Düz metin e-posta içeriği

            # Gönder
            server.sendmail(sender_email, recipient_email, msg.as_string())
            print(f"E-posta gönderildi: {recipient_email}")
    except Exception as e:
        print(f"E-posta gönderim hatası ({recipient_email}): {e}")

# Belirtilen aralıktaki e-postalara gönderim yap
def send_to_range(file_path, smtp_server, port, sender_email, sender_password, subject, body, start_index, end_index):
    email_list = get_email_list(file_path, start_index, end_index)
    for email in email_list:
        send_email(smtp_server, port, sender_email, sender_password, email, subject, body)

# Kullanıcı bilgileri ve parametreler
excel_file = "mails.xlsx"  # Excel dosyanızın adı
smtp_server = "smtp.yandex.com" veya "smtp.gmail.com" # SMTP sunucunu sec
port = 587
sender_email = ""  # Gönderen e-posta adresiniz
sender_password = ""  # Yandex uygulama şifreniz
subject = "Test Mail"  #MAilin konusunu yaz
body = "Merhaba, bu bir test mailidir. Lutfen dikkate almayiniz. İyi çalışmalar!"   # Mailin icerigini yaz

# Belirtilen e-posta aralığını gönder
start_index = 1  # Başlangıç e-posta indexi - su satirdan 
end_index = 2    # Bitiş e-posta indexi - su satira kadar ki
column_index = 8 # mailler excel dosyasinda kacinci sutunda  a =1 b=2 c=3 vs. 
send_to_range(excel_file, smtp_server, port, sender_email, sender_password, subject, body, start_index, end_index)
