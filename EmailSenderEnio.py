import time
import win32com.client as win32
import base64
from PIL import Image

def send_html_email_with_image(subject, body_html, image_path, sender, receiver):
    # Corrige la orientación de la imagen
    image = Image.open(image_path)
    image = image.rotate(0, expand=True)  # 0 grados de rotación

    # Guarda la imagen corregida en un archivo temporal
    temp_image_path = "temp_image.jpg"
    image.save(temp_image_path)

    # Codifica la imagen como base64
    with open(temp_image_path, "rb") as image_file:
        image_data = image_file.read()
        image_base64 = base64.b64encode(image_data).decode("utf-8")

    # Embed the image into the HTML content
    body_html_with_image = body_html.replace("<img src=\"\"", f"<img src=\"data:image/jpeg;base64,{image_base64}\"")

    # Create the Outlook application instance
    olapp = win32.Dispatch('Outlook.Application')

    # Create a new email item
    mail_item = olapp.CreateItem(0)

    # Set email properties
    mail_item.Subject = subject
    mail_item.BodyFormat = 2  # 2 represents HTML format for the body
    mail_item.HTMLBody = body_html_with_image

    mail_item.Sender = sender
    mail_item.To = receiver

    # Display the email window
    mail_item.Display(True)
    # mail_item.Send()

    # Elimina el archivo temporal después de enviar el correo electrónico
    import os
    os.remove(temp_image_path)


now = time.strftime("%m/%d")
subject = f"Check in {now}"
body_html = """<html>
<body>
<p>Hi [Recipient's Name], good morning,<br>
This is my clock in for today.<br>
Please let me know if you have any questions.<br>
Best regards,<br>
</p>
<p style='margin-bottom:12.0pt'><strong><span style='font-family:"Calibri",sans-serif'>Your Name</span></strong></p>
<p style='margin-bottom:12.0pt'><span style='font-size:9.0pt'>Your Job Title</o:p></span></p>
<p style='margin-bottom:12.0pt'><span style='font-size:9.0pt'>Your Company, LLC<o:p></o:p></span></p>
<p style='margin-bottom:12.0pt'>Email: <a href="mailto:youremail@example.com">youremail@example.com</a></p>
<img src=\"\" alt="Image">
<p style='margin-bottom:12.0pt'><span style='font-size:8.0pt'>This email and any files transmitted with it are confidential and intended solely for the use of the individual or entity to which they are addressed. If you have<br>
received this email in error please notify the system manager. Please note that any views or opinions presented in this email are solely those of the author and do<br>
not necessarily represent those of the Company. Finally, the recipient should check this email and any attachments for the presence of viruses. The Company<br> 
accepts no liability for any damage caused by any virus transmitted by this email.</span></p>
</body>
</html>"""
sender = "youremail@example.com"
receiver = "recipient@example.com"

# Use a general placeholder for the image path
image_path = r"C:\path\to\your\image.jpg"

send_html_email_with_image(subject, body_html, image_path, sender, receiver)