import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from string import Template
import os

# SMTP sunucu ayarları
smtp_sunucu = 'smtp.example.com'
smtp_port = 587
kullanici_adi = 'email@example.com'
sifre = 'parolanız'

# Script'in bulunduğu dizini alın
script_dizini = os.path.dirname(os.path.abspath(__file__))

# E-posta şablonunu okuyun
sablon_dosyasi = os.path.join(script_dizini, 'sablon.html')
with open(sablon_dosyasi, 'r', encoding='utf-8') as dosya:
    eposta_sablonu = dosya.read()

# E-posta şablonunu Template nesnesine dönüştürün
eposta_sablonu = Template(eposta_sablonu)

# Excel dosyasından e-posta adreslerini ve diğer bilgileri okuyun
excel_dosyasi = os.path.join(script_dizini, 'eposta_listesi.xlsx')
df = pd.read_excel(excel_dosyasi)

# Sütun isimlerini temizleyin
df.columns = df.columns.str.strip().str.lower().str.replace('-', '').str.replace(' ', '').str.replace('_', '')

# Sütun isimlerini yazdırın (kontrol amaçlı)
print("Sütun İsimleri:", df.columns.tolist())

# Her bir e-posta adresine e-posta gönderin
for index, satir in df.iterrows():
    satir = satir.fillna('')  # Eksik değerleri boş string ile doldurun
    alici_eposta = satir['eposta']
    
    # E-posta mesajını oluşturun
    mesaj = MIMEMultipart('mixed')
    mesaj['From'] = kullanici_adi
    mesaj['To'] = alici_eposta
    mesaj['Subject'] = 'Konu Başlığı'

    # Kişiselleştirilmiş e-posta gövdesi
    eposta_govdesi = eposta_sablonu.safe_substitute(**satir.to_dict())

    # E-posta gövdesini ekleyin
    mesaj_alani = MIMEText(eposta_govdesi, 'html', 'utf-8')
    mesaj.attach(mesaj_alani)

    # **PDF ekini ekleyin**
    pdf_dosyasi = os.path.join(script_dizini, 'dokuman.pdf')
    with open(pdf_dosyasi, 'rb') as dosya:
        mime_pdf = MIMEApplication(dosya.read(), _subtype='pdf')
        mime_pdf.add_header('Content-Disposition', 'attachment', filename=os.path.basename(pdf_dosyasi))
        mesaj.attach(mime_pdf)

    # E-postayı gönderin
    try:
        sunucu = smtplib.SMTP(smtp_sunucu, smtp_port)
        sunucu.starttls()
        sunucu.login(kullanici_adi, sifre)
        sunucu.sendmail(kullanici_adi, alici_eposta, mesaj.as_string())
        sunucu.quit()
        print(f"{alici_eposta} adresine e-posta gönderildi.")
    except Exception as e:
        print(f"{alici_eposta} adresine e-posta gönderilemedi. Hata: {str(e)}")
