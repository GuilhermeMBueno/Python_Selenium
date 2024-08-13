import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from PIL import Image
import time

def capture_screenshot(url, file_path, email, password):
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)
    
    driver.set_window_size(2000, 1320)
    
    driver.get("SITE RAIZ DO TENENT")
    time.sleep(2)

    # Login to SharePoint
    driver.find_element(By.NAME, "loginfmt").send_keys(email)
    driver.find_element(By.NAME, "loginfmt").send_keys(Keys.RETURN)
    time.sleep(2)
    driver.find_element(By.NAME, "passwd").send_keys(password)
    driver.find_element(By.NAME, "passwd").send_keys(Keys.RETURN)
    time.sleep(2)
    driver.find_element(By.ID, "idSIButton9").click()
    time.sleep(5)

    driver.get(url)
    time.sleep(5)  # Esperar a página carregar
    driver.save_screenshot(file_path)
    driver.quit()

def combine_images(image_files, output_file):
    images = [Image.open(file) for file in image_files]

    max_width = max(image.width for image in images)
    total_height = sum(image.height for image in images)

    combined_image = Image.new('RGB', (max_width, total_height))

    current_height = 0
    for image in images:
        combined_image.paste(image, (0, current_height))
        current_height += image.height

    combined_image.save(output_file)

def send_email(to_addresses, subject, body, attachment):
    from_address = "user"
    password = "senha"


    msg = MIMEMultipart()
    msg['From'] = from_address
    msg['To'] = ", ".join(to_addresses)  # Junta os emails com vírgula
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    with open(attachment, 'rb') as attachment_file:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment_file.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename= {attachment}')
        msg.attach(part)

    with smtplib.SMTP('smtp.office365.com', 587) as server:
        server.starttls()
        server.login(from_address, password)
        server.sendmail(from_address, to_addresses, msg.as_string())


def cleanup_files(files):
    for file in files:
        if os.path.exists(file):
            os.remove(file)

# URLs
urls = [
    "Lista de links",
    "Links2",
    "link3"
    ]

email = "user"
password = "senha"

screenshot_files = []

for i, url in enumerate(urls):
    file_path = f"screenshot_{i}.png"
    capture_screenshot(url, file_path, email, password)
    screenshot_files.append(file_path)

output_file = "Output.png"
combine_images(screenshot_files, output_file)

to_addresses = [
    "lista de e-mails a enviar",
    "e-mail2",
    "E-mail3"
]

html_body = """
        Alguma mensagem para o E-mail
"""

send_email(to_addresses, "Titulo do E-mail", html_body, output_file)

# Apagar os arquivos gerados após enviar e-mail
cleanup_files(screenshot_files + [output_file])
