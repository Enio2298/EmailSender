# HTML Email Sender with Embedded Image Using Outlook
Overview

This Python script automates sending an HTML email with an embedded image via Microsoft Outlook. The email is composed using a specified subject, body content, and an image, which is encoded in Base64 and embedded directly in the email body. The email is then sent using the win32com library for Outlook automation.
Features

    Send Emails via Outlook: Automates the email creation and sending process using Microsoft Outlook.
    Embed Image in Email Body: Embeds an image directly into the HTML body of the email using Base64 encoding.
    Support for Multiple Recipients: Allows sending the email to a primary recipient, with additional recipients in CC.
    Customizable Email Content: Modify the subject, body, and image path to tailor the email to your needs.

Prerequisites

    Python 3.x: Ensure you have Python installed on your machine.

    Libraries: Install the required Python libraries by running the following command:

    bash

    pip install pywin32 Pillow

        pywin32: Used to interact with the Microsoft Outlook application.
        Pillow: Used for image handling (rotation and saving).

    Microsoft Outlook: Outlook must be installed and configured on your system for this script to function.

How It Works
Key Steps:

    Image Handling:
        The script reads an image from the specified file path.
        The image is rotated by 0 degrees (i.e., no actual rotation) to ensure it maintains the correct orientation.
        The image is saved temporarily and then encoded as Base64.

    HTML Email Composition:
        The email body is written in HTML format.
        The image is embedded into the HTML content by replacing a placeholder <img> tag with the Base64-encoded image.

    Sending the Email:
        The win32com library is used to interact with Outlook, create a new email, set its subject, body (in HTML), and recipients.
        The email is sent immediately using the .Send() method.

    Temporary File Cleanup:
        The temporary image file is deleted after the email is sent.

Code Breakdown
Functions
send_html_email_with_image

This function takes in several parameters:

    subject: The subject of the email.
    body_html: The HTML body of the email, where the image will be embedded.
    image_path: The local file path of the image to be embedded.
    sender: The sender’s email address.
    receiver: The primary recipient's email address.
    carbon_copy: A string of email addresses to be added in CC, separated by semicolons.

The function:

    Opens the image, rotates it (if needed), and saves it temporarily.
    Encodes the image in Base64 and embeds it in the HTML body.
    Uses Outlook's COM interface to create and send the email.

Email Content Example

In this case, the email is a daily check-in to "Leticia", with a friendly message and a signature block:

html

<p>Hi Leticia, good morning, this is my clock in for today.</p>
<p>Best regards,<br>Enio Rodriguez</p>

The email includes a confidentiality disclaimer at the bottom and embeds an image inline.
Usage Example

python

now = time.strftime("%m/%d")
subject = f"Check in {now}"
body_html = """<html>...</html>"""  # Your email HTML here
image_path = r"C:\path\to\your\image.jpg"
sender = "your-email@example.com"
receiver = "recipient@example.com"
carbon_copy = "cc1@example.com; cc2@example.com"

send_html_email_with_image(subject, body_html, image_path, sender, receiver, carbon_copy)

Image Embedding

The image is embedded into the email using the Base64 format, which allows Outlook to display the image inline without needing to download it from an external source.
Running the Script

    Ensure Outlook is running and you are logged in.
    Update the email subject, body, sender, receiver, CC, and image path in the script as needed.
    Run the script from a Python environment:

    bash

    python send_email.py

Troubleshooting

    Outlook Issues: If the script fails to send the email, ensure that Outlook is properly installed and configured on your machine.
    Image Path: Verify that the image file path is correct. Use an absolute file path to avoid issues.

Conclusion

This script is a convenient tool for automating email sending via Outlook, particularly useful for scenarios requiring consistent or repetitive email templates with embedded images.
