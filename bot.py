from loguru import logger
import telebot
from telebot import types
import os
from docx2pdf import convert

TOKEN = "TOKEN"

BASE_DIR = os.getcwd()
bot = telebot.TeleBot(TOKEN)
logger.add("debug.log", format="{time} {level} {message}",
           level="DEBUG", rotation="10 MB", compression="zip")


logger.info("Bot was started for polling")

@logger.catch
@bot.message_handler(commands=['start', 'help'])
def send_welcome(message: types.Message) -> None:
    introduction = 'Hey man I\'m the docx to pdf conveter bot. Upload the docx file and if God permits I will convert it to pdf'
    bot.send_message(message.from_user.id, introduction)

@bot.message_handler(content_types=['document'])
def handle_docs(message: types.Message) -> None:
    try:
        if str(message.document.file_name)[str(message.document.file_name).find('.') + 1:] == "docx":
            save_file(message)
            pdf_file_name = generate_pdf(str(message.document.file_name))
            send_pdf_file(message, pdf_file_name)
        else:
            bot.send_message(message.from_user.id, "This is not .docx file! Please send Microsoft Word document file one more time.")
    except Exception as e:
        bot.reply_to(message, e)


def save_file(message: types.Message) -> None:
    try:
        file_name = message.document.file_name
        file_id_info = bot.get_file(message.document.file_id)
        downloaded_file = bot.download_file(file_id_info.file_path)
        src = file_name
        with open(BASE_DIR + "/" + src, 'wb') as new_file:
            new_file.write(downloaded_file)
        logger.info("[*] File added:\nFile name - {}\nFile directory - {}".format(str(file_name), str(BASE_DIR)))
        bot.send_message(message.chat.id, "Processing...")
    except Exception as ex:
        logger.error(str(ex))
        bot.send_message(message.chat.id, "[!] error - {}".format(str(ex)))


def generate_pdf(file_name: str) -> str:
    pdf_file_name = file_name.replace(".docx", ".pdf")
    convert(file_name)
    return pdf_file_name

def send_pdf_file(message: types.Message, pdf_file_name: str) -> None:
    directory = BASE_DIR + "\\" + pdf_file_name
    f = open(pdf_file_name, "rb")
    bot.send_document(message.chat.id, f)

bot.polling()
