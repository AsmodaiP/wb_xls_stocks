import codecs
import logging
import os
import sys
from logging.handlers import RotatingFileHandler

from dotenv import load_dotenv
from telegram.ext import (CommandHandler, ConversationHandler, Filters,
                          MessageHandler, Updater)

import sheet

load_dotenv()

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
        parse_mode='Markdown')


def file_manager(bot, update):
    file_id = bot.message.document['file_id']
    name = bot.message.document['file_name']
    if name == 'stocks.xlsx':
        update.user_data['file_id'] = file_id
        bot.message.reply_text('Введите ДА для подтверждения замены главного файла')
        return 'update_main_file_confirmation'
    file = update.bot.get_file(file_id)
    file.download(name)
    result = list(sheet.update_stocks(name)['erorrs'])
    bot.message.reply_text('Обновленная таблица')
    bot.message.reply_document(open('stocks.xlsx', 'rb'))
    if result:
        bot.message.reply_text('Нет этих артикулов в таблице\n'
                               +
                               '\n'.join(result), parse_mode='Markdown')
    return ConversationHandler.END

def update_main_file(bot, update):
    if bot.message.text.lower() != 'да':
        bot.message.reply_text('Замена главного файла отменена')
        return ConversationHandler.END

    bot.message.reply_text('Старая таблица')
    bot.message.reply_document(open('stocks.xlsx', 'rb'))

    file_id = update.user_data['file_id']
    name = 'stocks.xlsx'
    file = update.bot.get_file(file_id)
    file.download(name)


def cancel(bot, update):
    bot.message.reply_text('Операция отменена')
    return ConversationHandler.END


updater = Updater(token=TELEGRAM_TOKEN)

start_handler = CommandHandler('start', start)
updater.dispatcher.add_handler(start_handler)


update_main_file_handler = MessageHandler(Filters.text & ~Filters.command, update_main_file)

document_handler = MessageHandler(
    Filters.document.file_extension("xlsx"),
    file_manager)

dialog = ConversationHandler(
    entry_points=[document_handler],
    states={
        'update_main_file_confirmation':  [MessageHandler(Filters.text & ~Filters.command, update_main_file)],
    },
    fallbacks=[CommandHandler('cancel', cancel)])

updater.dispatcher.add_handler(dialog)

updater.start_polling()