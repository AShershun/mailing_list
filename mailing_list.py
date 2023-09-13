import smtplib
import xml.etree.ElementTree as ET
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from config import *
from PyQt5.QtWidgets import QApplication, QFileDialog

# Створення програми PyQt
app = QApplication([])

# Відкриття діалогового вікна для вибору файлу XML
options = QFileDialog.Options()
options |= QFileDialog.ReadOnly
xml_file_path, _ = QFileDialog.getOpenFileName(None, "Выберите файл", "", "XML Files (*.xml)", options=options)
if not xml_file_path:
    print("XML файл не выбран")
    exit()

# Закриття програми PyQt
app.quit()

tree = ET.parse(xml_file_path)
root = tree.getroot()

# Читання вмісту HTML-шаблону листа
with open("email_template.html", "r", encoding="utf-8") as template_file:
    email_template = template_file.read()
namespace = {'ss': 'urn:schemas-microsoft-com:office:spreadsheet'}

count_letter = 0
count_failure = 0
log_error = ''

kit_check = input('Додати у лист посилання на комплекти першокурсників? | Y/N: ').strip().lower()

# Перебір елементів у XML
for row in root.findall(".//ss:Row", namespaces=namespace)[1:]: # Ігнорування першого рядка заголовків
    
    cells = row.findall("ss:Cell/ss:Data", namespaces=namespace) 

    last_name = cells[1].text
    first_name = cells[2].text
    email = cells[6].text
    code = cells[9].text
    faculty = cells[10].text
        
    #Посилання першокурсника
    itfrhb = ''
    kit_link = "https://elc.library.ontu.edu.ua/library-w/DocumentSearchForm"

    
    if faculty == "Ф-т технології вина та туристичного бізнесу":
        kit_link = "https://drive.google.com/file/d/1qcDLYsna4kTUpBfU7Eje674WG15MgzKF/view"
    if faculty == "Ф-т комп`ютерних систем та автоматизації":
        kit_link = "https://drive.google.com/file/d/1lQCGt9l0ffMDCc-DCtMPNaaj5RfTXWex/view"
    if faculty == "Ф-т економіки, бізнесу і контролю":
        kit_link = "https://drive.google.com/file/d/19i6_c43KyW5VV0iS1NoSndB52elAiT4_/view"
    if faculty == "Ф-т інноваційних технологій харчування і ресторанно-готельного бізнесу":
        kit_link = "https://drive.google.com/file/d/1BQpclMifZWgJzmX1tBqmAQpa6uT6fEzW/view"
        itfrhb = f' або <a href="https://drive.google.com/file/d/16pLFJVm4oduhqIf8mXUMnBRIBC3WYxJ-/view">Посилання</a>'
    if faculty == "Ф-т технології зерна і зернового бізнесу":
        kit_link = "https://drive.google.com/file/d/1uObvGs7L22I2sh_yZJOfdjSoosLjyXTG/view"
    if faculty == "Ф-т менеджменту, маркетингу і логістики":
        kit_link = "https://drive.google.com/file/d/1Mfa6vxi3avSP_ahK3fYjr0rWTMenK6_j/view"
    if faculty == "Ф-т технології та товарознавства харчових продуктів і продовольчого бізнесу":
        kit_link = "https://drive.google.com/file/d/1rk06P5dayE9cZA588fFQXkEjqoVIYibA/view"
    if faculty == "Ф-т комп`ютерної інженерії, програмування та кіберзахисту":
        kit_link = "https://drive.google.com/file/d/1Mquj4zOtsRxVYTuIsP4osMIhESrTuok3/view"
    if faculty == "Ф-т низькотемпературної техніки та інженерної механіки":
        kit_link = "https://drive.google.com/file/d/1_dOBn8w7g8wjsXcZQsxgHzpcvQPXH3L6/view"
    if faculty == "Ф-т нафти, газу та екології":
        kit_link = "https://drive.google.com/file/d/1anWo0R2t2iQ8Qw83Y-LuytKx2lWr3vMD/view"
        
    if kit_check == 'y':
        kit_paragraph = f'<p><b style="color: rgb(92, 67, 116)">Комплект першокурсника: </b><a href="{kit_link}">Посилання</a>{itfrhb}</p>'
    elif kit_check == 'n':
        kit_paragraph = ''
    else:
        print("Введено неправильну відповідь. Використовуйте 'Y' або 'N'.")    
                    
    personalized_message = email_template.replace("{{LAST_NAME}}", last_name).replace("{{FIRST_NAME}}", first_name).replace("{{CODE}}", code).replace("{{KIT_PARAGRAPH}}", kit_paragraph)

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