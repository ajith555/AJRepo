import sys
import win32com.client
import os
import datetime

def fetch_and_save_email(subject_keyword, folder_name):
    outlook = win32com.client.Dispatch('outlook.application').GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder('6').Folders(folder_name)
    messages = inbox.Items

    messages.Sort("[ReceivedTime]", True)

    # Fetch only the latest email with subject containing the specified keyword
    latest_msg = None
    for msg in messages:
        if subject_keyword in str(msg.Subject):
            latest_msg = msg
            break

    # Check if there is a latest email with the specified subject keyword
    if latest_msg:
        # Fetch To, CC, Subject, and From information from the latest email
        latest_subject = latest_msg.Subject
        latest_to = latest_msg.To
        latest_cc = latest_msg.CC
        latest_sender = latest_msg.SenderName

        print("Latest Subject:", latest_subject)
        print("To:", latest_to)
        print("CC:", latest_cc)
        print("From:", latest_sender)

        path = 'P:/Documents/PycharmProjects/pythonLearning/' + folder_name + '/'
        if not os.path.exists(path):
            os.makedirs(path)

        for atch in latest_msg.Attachments:
            # Save attachments with unique file names
            attachment_path = os.path.join(path, atch.FileName)
            if os.path.exists(attachment_path):
                # If the file already exists, add date and time as an extension to the file name
                unique_name = os.path.splitext(atch.FileName)
                now = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                new_attachment_name = f"{unique_name[0]}_{now}{unique_name[1]}"
                attachment_path = os.path.join(path, new_attachment_name)

            atch.SaveAsFile(attachment_path)

        print("Attachments saved successfully.")
    else:
        print(f"No emails with subject containing '{subject_keyword}' found in the folder.")

if __name__ == "__main__":
    # Command-line arguments:
    # sys.argv[0]: Script name (e.g., outlook_email.py)
    # sys.argv[1]: Subject keyword (e.g., 'FFTS')
    # sys.argv[2]: Folder name (e.g., 'ABCTest')
    if len(sys.argv) == 3:
        subject_keyword = sys.argv[1]
        folder_name = sys.argv[2]
        fetch_and_save_email(subject_keyword, folder_name)
    else:
        print("Usage: python outlook_email.py <subject_keyword> <folder_name>")
