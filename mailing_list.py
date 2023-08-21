import smtplib
import xml.etree.ElementTree as ET
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from config import *


from PyQt5.QtWidgets import QApplication, QFileDialog

# Создание приложения PyQt
app = QApplication([])

# Открытие диалогового окна для выбора файла XML
options = QFileDialog.Options()
options |= QFileDialog.ReadOnly
xml_file_path, _ = QFileDialog.getOpenFileName(None, "Выберите файл", "", "XML Files (*.xml)", options=options)
if not xml_file_path:
    print("XML файл не выбран")
    exit()

# Закрытие приложения PyQt
app.quit()


tree = ET.parse(xml_file_path)
root = tree.getroot()


# Читання вмісту HTML-шаблону листа
with open("email_template.html", "r", encoding="utf-8") as template_file:
    email_template = template_file.read()
namespace = {'ss': 'urn:schemas-microsoft-com:office:spreadsheet'}

count_letter = 0
count_failure = 0
log_error = ""

# Перебір елементів у XML
for row in root.findall(".//ss:Row", namespaces=namespace)[1:]: # Ігнорування першого рядка заголовків
    cells = row.findall("ss:Cell/ss:Data", namespaces=namespace) 

    last_name = cells[1].text
    first_name = cells[2].text
    email = cells[6].text
    code = cells[9].text
        
    personalized_message = email_template.replace("{{LAST_NAME}}", last_name).replace("{{FIRST_NAME}}", first_name).replace("{{CODE}}", code)

    # Відправка листа
    try:
        msg = MIMEMultipart()
        msg['From'] = smtp_username
        msg['To'] = email
        msg['Subject'] = "Реєстрація в електронному каталозі НТБ ОНТУ"
        msg.attach(MIMEText(personalized_message, 'html'))

        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(smtp_username, smtp_password)
        server.sendmail(smtp_username, email, msg.as_string())
        server.quit()

        log_error += f"Лист відправлений на {email}\n"
    except Exception as e:
        count_failure += 1
        log_error += f"Помилка при надсиланні листа на {email} {last_name} {first_name}: {e}\n\n"

    count_letter += 1


conclusion = f"Всього відправлено листів: {count_letter}, успішно: {count_letter-count_failure}, невдало: {count_failure}\n"

# Запис логів
with open("mailing_results.txt", "w") as log_file:
    log_file.write(conclusion)
    log_file.write(log_error)

