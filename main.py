from telethon.tl.types import InputPeerEmpty
from telethon.tl.functions.messages import GetDialogsRequest
import socks
import openpyxl
import os
from telethon.sync import TelegramClient
import credentials


# Excel file constants
FIRST_COL = 'A'
SECOND_COL = 'B'
THIRD_COL = 'C'
FILE = 'XLS FILE NAME.XLS'
SHEET = 'SHEET NAME (Sheet1, Sheet2, etc.)'
DIR = 'FILE DIRECTORY GOES HERE'

# Open XLS file

os.chdir(DIR)
xls_file = openpyxl.load_workbook(FILE)
xls_sheet = xls_file[SHEET]


def scrape_groups():

    global client
    chats = []
    groups = []
    last_date = None
    chunk_size = 200

    # Telegram session login and otp
    client.connect()
    if not client.is_user_authorized():
        client.send_code_request(credentials.phone)
        client.sign_in(credentials.phone, input('Enter OTP code: '))

    result = client(GetDialogsRequest(
        offset_date=last_date,
        offset_id=0,
        offset_peer=InputPeerEmpty(),
        limit=chunk_size,
        hash=0
    ))

    chats.extend(result.chats)

    for chat in chats:
        try:
            if chat.megagroup == True:
                groups.append(chat)
        except:
            continue

    return groups


def write_to_xls(groups):

    for index, target_group in enumerate(groups):
        print("Scraping Group ID:" + str(target_group.id) +
              " - Name: " + str(target_group.title) + "... \n")
        xls_sheet[FIRST_COL+str(index+2)].value = target_group.id
        xls_sheet[SECOND_COL+str(index+2)].value = target_group.username
        xls_sheet[THIRD_COL+str(index+2)].value = target_group.title

    print("Done!")
    xls_file.save(FILE)


# Log in to telegram API
client = TelegramClient('anon', credentials.api_id, credentials.api_hash, proxy=(
    socks.SOCKS5, '127.0.0.1', 7890))

# Scrape and save to Excel file
groups = scrape_groups()
write_to_xls(groups)
