import win32com.client as client
import os
import subprocess
import fnmatch

option_dict = {
    1: "CONTAINER LABEL ",
    2: "CONTENT LIST LABEL ",
    3: "FISCHER SHIPPING LABEL ",
}
edi_message = "C:\\EDI_Messages\\"
no_reply = "no.reply@ap.mail.com"
mailL_folder_name = "99.EDI"


def select_option():
    option_message = (
        "Please select OPTION:\n"
        + "1.CONTAINER LABEL\n"
        + "2.FISCHER CONTENT LIST LABEL\n"
        + "3.FISCHER SHIPPING LABEL\n"
    )
    while True:
        user_selected = input(option_message)
        if (
            user_selected.isnumeric()
            and int(user_selected) >= 0
            and int(user_selected) <= 3
        ):
            return int(user_selected)


def user_inputed(option):
    inputted_id = input(option_dict.get(option)).strip()
    return inputted_id


def download_pdf_file(file_type, file_id):
    # create outlook instance
    outlook = client.Dispatch("Outlook.Application")

    # get the namespace object
    namespace = outlook.GetNameSpace("MAPI")

    # get inbox folder
    inbox = namespace.GetDefaultFolder(6)  # Index for inbox is 6
    # get edi mail folder
    # edi_folder = inbox.Folders[0]
    edi_folder = inbox
    for folder in inbox.Folders:
        if folder.name == mailL_folder_name:
            edi_folder = folder
            break

    for message in edi_folder.Items:
        if (file_type in message.subject) and (file_id in message.subject) and (no_reply == message.Sender.Address):
            attachment = message.Attachments[0]
            pdf_file_path = edi_message + attachment.FileName
            attachment.SaveAsFile(pdf_file_path)
            return pdf_file_path


def open_saved_file(path):
    if os.path.exists(path):
        subprocess.Popen([path], shell=True)


def create_folder_if_not_exist(folder):
    if not os.path.exists(folder):
        os.mkdir(folder)


def get_path_of_first_found_file_in_directory(patt):
    for filename in os.listdir(edi_message):
        if fnmatch.fnmatch(filename, patt):
            return edi_message + filename


while True:
    try:
        # select options in [1, 2, 3]
        option = select_option()
        # 0 to terminate
        if option == 0:
            break
        # ID inputted
        id = user_inputed(option)
        # create folder to save EDI file if not existed
        create_folder_if_not_exist(edi_message)
        # if file already downloaded
        patt = f"*{option_dict.get(option)}*{id}*.pdf"
        existed_file = get_path_of_first_found_file_in_directory(patt)
        if existed_file is not None:
            open_saved_file(existed_file)
            # print("open from local directory")
        else:
            # download edi attachment to created folder
            path = download_pdf_file(option_dict.get(option), id)
            # open the downloaded file with shell mode if file existed
            # print("open from mail")
            open_saved_file(path)
    except Exception as e:
        print("No EDI messages")
        print(e)
