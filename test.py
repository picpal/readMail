import pythoncom
import win32com.client
from bs4 import BeautifulSoup
import json
import os


def on_new_mail(item):
    subject = item.Subject
    body = item.HTMLBody

    # 보낸 사람 정보 확인
    try:
        sender = item.Sender
        sender_email = sender.Address
    except AttributeError:
        sender_email = "Unknown"

    soup = BeautifulSoup(body, "html.parser")

    # 메일 정보
    data = {
        "subject": subject,
        "sender": sender_email,
        "body": soup.prettify(),
        # "body": soup.get_text(),
    }

    # json 데이터로 변환
    json_data = json.dumps(data, indent=10, ensure_ascii=False)
    print(json_data)

    # 첨부파일 다운로드
    filePath = "C:/Users/bwc/Desktop/picpal/"
    for attachment in item.Attachments:
        attachment_path = os.path.join(filePath, attachment.FileName)
        if os.path.exists(attachment_path):
            file_name, file_extension = os.path.splitext(attachment_path)
            i = 1
            fileNumber = "(" + str(i) + ")"
            new_attachment_path = file_name + fileNumber + file_extension
            while os.path.exists(new_attachment_path):
                i += 1
                new_attachment_path = file_name + fileNumber + file_extension
            attachment_path = new_attachment_path

        attachment.SaveAsFile(attachment_path)


class ServiceEvents:
    def OnNewMailEx(self, receivedItemsIDs):
        for ID in receivedItemsIDs.split(","):
            item = outlook.Session.GetItemFromID(ID)
            on_new_mail(item)


outlook = win32com.client.DispatchWithEvents(
    "Outlook.Application", ServiceEvents)

print("Waiting for new emails...")
pythoncom.PumpMessages()
