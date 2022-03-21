import codecs
import logging
import os
import sys
from logging.handlers import RotatingFileHandler

from dotenv import load_dotenv
from telegram.ext import (CommandHandler, ConversationHandler, Filters,
                          MessageHandler, Updater)
from telegram import KeyboardButton, ReplyKeyboardMarkup
from google_sheet import update_google_sheet

import sheet

load_dotenv()

MAIN_MENU = (
    [KeyboardButton('Получить таблицу'), ],)

MAIN_MENU_MARKUP = ReplyKeyboardMarkup(
    MAIN_MENU,
    resize_keyboard=True,
    one_time_keyboard=False)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
log_file = os.path.join(BASE_DIR, 'bot.log')

console_handler = logging.StreamHandler()
file_handler = RotatingFileHandler(
    log_file,
    maxBytes=100000,
    backupCount=3,
    encoding='utf-8'
)
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s, %(levelname)s, %(message)s',
    handlers=(
        file_handler,
        console_handler
    )
)
try:
    TELEGRAM_TOKEN = os.environ['TELEGRAM_TOKEN']
except KeyError as exc:
    logging.error(exc, exc_info=True)
    sys.exit('Не удалось получить переменные окружения')


def start(bot, update):
    bot.message.reply_text(
        codecs.open(
            'instruction.md',
            'r').read(),
        parse_mode='Markdown', reply_markup=MAIN_MENU_MARKUP)


def send_old_table(bot, update):
    bot.message.reply_text('Старая таблица')
    bot.message.reply_document(open('stocks.xlsx', 'rb'))


def send_new_table(bot, update):
    bot.message.reply_text('Актуальная таблица')
    bot.message.reply_document(open('stocks.xlsx', 'rb'))
    update_google_sheet()


def file_manager(bot, update):
    '''
    Используется для обработки файлов и определения того,
    нужно ли добавить остатки из этого файла,
    обновить основную табилцу или внести приход
    '''
    send_old_table(bot, update)
    file_id = bot.message.document['file_id']
    name = bot.message.document['file_name']
    if name == 'stocks.xlsx':
        update.user_data['file_id'] = file_id
        bot.message.reply_text(
            'Введите ДА для подтверждения замены главного файла')
        return 'update_main_file_confirmation'
    file = update.bot.get_file(file_id)
    file.download(name)
    if name == 'поставка.xlsx':
        update.user_data['file_id'] = file_id
        sdfsd = sheet.insert_supplie(name)
        logging.info(sdfsd)
        result = sheet.insert_supplie(name)['erorrs']
    else:
        result = list(sheet.insert_sales(name)['erorrs'])
    send_new_table(bot, update)
    if result:
        bot.message.reply_text('Этих артикулов  нет в таблице. Скорее всего нужно добавить их и их актуальный остаток в файл и провести замену главного файла\n'
                               +
                               '\n'.join(result), parse_mode='Markdown')
    return ConversationHandler.END


def update_main_file(bot, update):
    '''Обновляет главный файл и отправляет его пользователю'''
    if bot.message.text.lower() != 'да':
        bot.message.reply_text('Замена главного файла отменена')
        return ConversationHandler.END

    file_id = update.user_data['file_id']
    name = 'stocks.xlsx'
    file = update.bot.get_file(file_id)
    file.download(name)
    send_new_table(bot, update)
    return ConversationHandler.END


def cancel(bot, update):
    '''Служит стандартным выходом из любого места'''
    bot.message.reply_text('Операция отменена')
    return ConversationHandler.END


updater = Updater(token=TELEGRAM_TOKEN)

start_handler = CommandHandler('start', start)
updater.dispatcher.add_handler(start_handler)


update_main_file_handler = MessageHandler(
    Filters.text & ~Filters.command, update_main_file)

document_handler = MessageHandler(
    Filters.document.file_extension("xlsx"),
    file_manager)

get_table_handler = MessageHandler(
    Filters.text(
        ['Получить таблицу']),
    send_new_table)
updater.dispatcher.add_handler(get_table_handler)

dialog = ConversationHandler(
    entry_points=[document_handler],
    states={
        'update_main_file_confirmation': [MessageHandler(Filters.text & ~Filters.command, update_main_file)],
    },
    fallbacks=[CommandHandler('cancel', cancel)])

updater.dispatcher.add_handler(dialog)

updater.start_polling()
