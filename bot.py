import logging
import requests
import time
import sys
import xlsxwriter
import threading
import os
import pandas as pd
import glob
from logging import info
from telegram import Update, InlineKeyboardMarkup, InlineKeyboardButton
from telegram.ext import Updater, CallbackContext, CommandHandler, MessageHandler, ConversationHandler, Filters, \
    CallbackQueryHandler

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                    filename='error.log')
info('Loading dependencies...')


class ST:
    a = str()
    b = str()
    count = 1
    all = 0
    date, token, file = range(3)
    CANCEL_INLINE_KEYBOARD = [[InlineKeyboardButton("ОТМЕНИТЬ ВСЕ", callback_data='chatcancel')]]
    START = [[InlineKeyboardButton("запуск", callback_data='start')]]
    auth = [[InlineKeyboardButton("✘", callback_data='del')], [InlineKeyboardButton("✔", callback_data='send_log')]]


class bot:
    def __init__(self):
        self.bot = Updater('5116492940:AAES0YfQVbVOcaxUdwNSR5ZmZ1YYGIhuptM')
        self.dispatcher = self.bot.dispatcher
        self.chat = ConversationHandler(
            entry_points=[CommandHandler('start', self.start)],
            states={
                ST.date: [MessageHandler(Filters.all, self.date)],
                ST.token: [MessageHandler(Filters.all, self.token)],
                ST.file: [MessageHandler(Filters.all, self.file)]
            },
            fallbacks=[CallbackQueryHandler(self.chat_cancel, pattern=r'^chatcancel$')]
        )
        self.dispatcher.add_handler(self.chat)
        self.dispatcher.add_handler(CommandHandler('clear', self.clear))
        self.dispatcher.add_handler(CallbackQueryHandler(self.send, pattern=r'^start$'))
        self.dispatcher.add_handler(CallbackQueryHandler(self.send_log, pattern=r'^send_log$'))
        self.dispatcher.add_handler(CallbackQueryHandler(self.deli, pattern=r'^del$'))
        self.dispatcher.add_handler(CommandHandler('log', self.log))

    def start(self, update: Update, context: CallbackContext):
        update.message.reply_text('введите дату в формате yyyy-MM-dd\nпример: 2022-01-31',
                                  reply_markup=InlineKeyboardMarkup(ST.CANCEL_INLINE_KEYBOARD, one_time_keyboard=False))
        return ST.date

    def deli(self, update: Update, context: CallbackContext):
        update.callback_query.message.delete()

    def clear(self, update: Update, context: CallbackContext):
        filename_list = glob.glob('*.xlsx')
        for filename in filename_list:
            os.remove(filename)
        else:
            update.message.reply_text(f'успешно удаленны фаилы {filename_list}')

    def send_log(self, update: Update, context: CallbackContext):
        doc = open('error.log', 'rb')
        update.callback_query.message.reply_document(doc)
        doc.close()

    def date(self, update: Update, context: CallbackContext):
        if '.' in update.message.text:
            text = update.message.text.split('.')
            ST.a = '-'.join(text)
        else:
            ST.a = update.message.text
            update.message.reply_text('теперь введите токен:',
                                      reply_markup=InlineKeyboardMarkup(ST.CANCEL_INLINE_KEYBOARD,
                                                                        one_time_keyboard=False))
            ST.a = update.message.text
            return ST.token

    def log(self, update: Update, context: CallbackContext):
        update.message.reply_text('вам отправится фаил с ошибкой подтвердите:',
                                  reply_markup=InlineKeyboardMarkup(ST.auth, one_time_keyboard=False))

    def token(self, update: Update, context: CallbackContext):
        ST.b = update.message.text
        s = requests.Session()
        headers = {
            'X-Mpstats-TOKEN': ST.b,
            'Content-Type': 'application/json',
        }

        response = requests.get('http://mpstats.io/api/user/report_api_limit', headers=headers)
        if response.status_code == 200:
            a = response.json()['available'] - response.json()['use']
            update.message.reply_text(f'токен валидный\nОсталось использований: {a}\nтеперь exel фаил:',
                                      reply_markup=InlineKeyboardMarkup(ST.CANCEL_INLINE_KEYBOARD,
                                                                        one_time_keyboard=False)
                                      )
            info(f'токен:\n{ST.b}\nОсталось использований: {a}')
            return ST.file
        else:
            update.message.reply_text(f'что то не так с кодом\nпопробуйте еще раз',
                                      reply_markup=InlineKeyboardMarkup(ST.CANCEL_INLINE_KEYBOARD,
                                                                        one_time_keyboard=False)
                                      )
            return ST.token

    def file(self, update: Update, context: CallbackContext):
        try:
            os.remove(f'{update.message.chat_id}_{ST.a}.xlsx')
            os.remove(f'{update.message.chat_id}.xlsx')
        except:
            pass
        f = context.bot.getFile(update.message.document.file_id)
        f.download(f'./{update.message.chat_id}.xlsx')
        update.message.reply_text(f'date: {ST.a}\ntoken: {ST.b}\nфаил загружен',
                                  reply_markup=InlineKeyboardMarkup(ST.START, one_time_keyboard=False))
        return ConversationHandler.END

    def send(self, update: Update, context: CallbackContext):
        x = update.callback_query.message.chat.id
        info(f'использован токен: {ST.b}')
        headers = {
            'X-Mpstats-TOKEN': ST.b,
            'Content-Type': 'application/json'
        }
        param = {
            'd1': ST.a,
            'd2': ST.a
        }
        param_f = {
            'd1': ST.a,
            'd2': ST.a,
            'full': 'true'
        }
        param_fbs = {
            'd': ST.a
        }
        numbers = pd.read_excel(f'{x}.xlsx', index_col='Номенклатура')
        numbers.head()
        workbook = xlsxwriter.Workbook(f'./{x}_{ST.a}.xlsx')
        worksheet = workbook.add_worksheet()
        # ширина колон
        worksheet.set_column(0, 0, 30)
        # worksheet.set_column(0, 1, 30)
        worksheet.set_column(0, 1, 30)
        worksheet.set_column(0, 2, 30)
        worksheet.set_column(0, 3, 30)
        worksheet.set_column(0, 4, 20)
        worksheet.set_column(0, 5, 30)
        worksheet.set_column(0, 6, 30)
        worksheet.set_column(0, 7, 30)
        worksheet.set_column(0, 8, 30)
        worksheet.set_column(0, 9, 30)
        worksheet.set_column(0, 10, 30)
        worksheet.write(0, 0, 'Номенклатура')
        worksheet.write(0, 1, 'Цена товара')
        worksheet.write(0, 2, 'Кол-во категорий')
        # worksheet.write(0, 4, 'Выдача по категориям')
        worksheet.write(0, 3, 'Остаток')
        worksheet.write(0, 4, 'Кол-во кл.слов')
        worksheet.write(0, 5, 'Ср. позиция')
        worksheet.write(0, 6, 'Сумма частотности')
        worksheet.write(0, 7, 'Кол-во продаж')
        worksheet.write(0, 8, 'Кол-во продаж(fbs)')
        worksheet.write(0, 9, 'размеры')
        worksheet.write(0, 10, 'итого по складам')
        row = 1
        col = 0
        s = requests.Session()
        try:
            count = 0
            for i in numbers.index:
                update.callback_query.message.edit_text(f'выполненно {count} из {len(numbers.index)}')
                count += 1
                resSales = s.get(f'https://mpstats.io/api/wb/get/item/{int(i)}/sales', headers=headers, params=param)
                resSales.raise_for_status()
                if resSales.status_code != 204:
                    jsonRS = resSales.json()
                else:
                    logging.error(f"{time.localtime().tm_mday}/{time.localtime().tm_mon}/{time.localtime().tm_year}"
                                  f" {time.localtime().tm_hour}:{time.localtime().tm_min}:{time.localtime().tm_sec}"
                                  f" - Error: Code 204; Нет содержимого в ответе (запрос resSales) debug:"
                                  f" Ошибка на номенклатуре: {int(i)}")
                    continue
                resSalesfbs = s.get(f'https://mpstats.io/api/wb/get/item/{int(i)}/balance_by_day',
                                    headers=headers, params=param_fbs)
                resSalesfbs.raise_for_status()
                if resSalesfbs.status_code != 204:
                    jsonRSfbs = resSalesfbs.json()
                else:
                    logging.error(f"{time.localtime().tm_mday}/{time.localtime().tm_mon}/{time.localtime().tm_year} "
                                  f"{time.localtime().tm_hour}:{time.localtime().tm_min}:{time.localtime().tm_sec}"
                                  f" - Error: Code 204; Нет содержимого в ответе (запрос resSalesfbs) debug:"
                                  f" Ошибка на номенклатуре: {int(i)}")
                    continue
                resCategory = s.get(f'https://mpstats.io/api/wb/get/item/{int(i)}/by_category',
                                    headers=headers, params=param)
                resCategory.raise_for_status()
                if resCategory.status_code != 204:
                    jsonRC = resCategory.json()
                else:
                    logging.error(f"{time.localtime().tm_mday}/{time.localtime().tm_mon}/{time.localtime().tm_year}"
                                  f" {time.localtime().tm_hour}:{time.localtime().tm_min}:{time.localtime().tm_sec} "
                                  f"- Error: Code 204; Нет содержимого в ответе (запрос resCategory) debug:"
                                  f" Ошибка на номенклатуре: {int(i)}")
                    continue

                resKeyWords = s.get(f'https://mpstats.io/api/wb/get/item/{int(i)}/by_keywords', headers=headers,
                                    params=param)
                resKeyWords.raise_for_status()

                if resKeyWords.status_code != 204:
                    jsonKeys = resKeyWords.json()
                else:
                    logging.error(f"{time.localtime().tm_mday}/{time.localtime().tm_mon}/{time.localtime().tm_year}"
                                  f" {time.localtime().tm_hour}:{time.localtime().tm_min}:{time.localtime().tm_sec} - "
                                  f"Error: Code 204; Нет содержимого в ответе (запрос resKeyWords) debug:"
                                  f" Ошибка на номенклатуре: {int(i)}")
                    continue
                remains = s.post(f'http://mpstats.io/api/wb/get/item/{int(i)}/orders_by_size', headers=headers,
                                 params=param_f)
                if resSales.status_code == 200 and resCategory.status_code == 200 and resKeyWords.status_code == 200 \
                        and resSalesfbs.status_code == 200:
                    countSale = jsonRS[0]['sales']  # количество продаж
                    balance = jsonRS[0]['balance']  # остаток
                    price = jsonRS[0]['final_price']  # цена товара
                    countCategorie = len(jsonRC['categories'])  # кол-во категорий
                    countKeyWords = len(jsonKeys['words'])  # кол-во ключ. слов
                    outputKeyWords = 0  # сумма выдачи ключевых слов
                    avgPos = 0
                    for j in jsonKeys['words']:
                        avgPos += jsonKeys['words'][j]['avgPos']
                        outputKeyWords += jsonKeys['words'][j]['total']
                    if avgPos:
                        avgPos = avgPos // countKeyWords
                    countSalefbs = 0  # кол-во продаж fbs
                    for k in range(len(jsonRSfbs)):
                        countSalefbs += (jsonRSfbs[k]['sales'] + jsonRSfbs[k]['salesfbs'])
                    # writer.writerow([str(i), countSale, price, str(countCategorie), str(outputCategSum), balance])
                    worksheet.write(row, col, str(int(i)))
                    worksheet.write(row, col + 1, price)
                    worksheet.write(row, col + 2, str(countCategorie))
                    # worksheet.write(row, col+4, str(outputCategSum))
                    worksheet.write(row, col + 3, balance)
                    worksheet.write(row, col + 4, countKeyWords)
                    worksheet.write(row, col + 5, avgPos)
                    worksheet.write(row, col + 6, outputKeyWords)
                    worksheet.write(row, col + 7, countSale)
                    worksheet.write(row, col + 8, countSalefbs)
                    stolb = 9


                    countbalance = 0
                    try:
                        minsize = remains.json()[ST.a][0]
                        maxsize = remains.json()[ST.a][-1]
                        print(minsize, maxsize)
                        for m in remains.json()[ST.a]:
                            countbalance += int(remains.json()[ST.a][m]['balance'])
                        worksheet.write(row, col + 9, f'{minsize}-{maxsize}')
                        worksheet.write(row, col + 10, countbalance)
                    except:
                        pass

                    row += 1

                elif resSales.status_code == 429 or resCategory.status_code == 429 or resKeyWords.status_code == 429 \
                        or resSalesfbs.status_code == 429:
                    logging.error(f"{time.localtime().tm_mday}/{time.localtime().tm_mon}/{time.localtime().tm_year}"
                                  f" {time.localtime().tm_hour}:{time.localtime().tm_min}:{time.localtime().tm_sec}"
                                  f" - Error: Code 429; {jsonRS['message']}; debug: Ошибка на номенклатуре: {int(i)}")
                    # workbook.close()
                    continue
                elif resSales.status_code == 401 or resCategory.status_code == 401 or resKeyWords.status_code == 401 \
                        or resSalesfbs.status_code == 401:
                    logging.error(f"{time.localtime().tm_mday}/{time.localtime().tm_mon}/{time.localtime().tm_year}"
                                  f" {time.localtime().tm_hour}:{time.localtime().tm_min}:{time.localtime().tm_sec}"
                                  f" - Error: Code 401; {jsonRS['message']}; Не правильный токен!; debug:"
                                  f" Ошибка на номенклатуре: {int(i)}")
                    # workbook.close()
                    continue
                elif resSales.status_code == 500 or resCategory.status_code == 500 or resKeyWords.status_code == 500 \
                        or resSalesfbs.status_code == 500:
                    logging.error(f"{time.localtime().tm_mday}/{time.localtime().tm_mon}/{time.localtime().tm_year}"
                                  f" {time.localtime().tm_hour}:{time.localtime().tm_min}:{time.localtime().tm_sec}"
                                  f" - Error: Code 500; {jsonRS['message']}; debug:"
                                  f" Ошибка на номенклатуре: {int(i)}")
                    # workbook.close()
                    continue
                elif resSales.status_code == 202 or resCategory.status_code == 202 or resKeyWords.status_code == 202 \
                        or resSalesfbs.status_code == 202:
                    logging.error(f"{time.localtime().tm_mday}/{time.localtime().tm_mon}/{time.localtime().tm_year}"
                                  f" {time.localtime().tm_hour}:{time.localtime().tm_min}:{time.localtime().tm_sec}"
                                  f" - Error: Code 202; {jsonRS['message']}; debug: Ошибка на номенклатуре: {int(i)}")
                    # workbook.close()
                    continue
                else:
                    # print(f"Номенклатура {int(i)} не найдена!")
                    logging.error(f"{time.localtime().tm_mday}/{time.localtime().tm_mon}/{time.localtime().tm_year}"
                                  f" {time.localtime().tm_hour}:{time.localtime().tm_min}:{time.localtime().tm_sec}"
                                  f" - Error: Что-то не так!; Sales = {resSales.status_code}; Category = "
                                  f"{resCategory.status_code}; KeyWords = {resKeyWords.status_code}; Salesfbs = "
                                  f"{resSalesfbs.status_code}; debug: Ошибка на номенклатуре: {int(i)}")
                    # workbook.close()
                    continue
        except Exception as e:
            # print(e)
            logging.error(f"{time.localtime().tm_mday}/{time.localtime().tm_mon}/{time.localtime().tm_year} "
                          f"{time.localtime().tm_hour}:{time.localtime().tm_min}:{time.localtime().tm_sec}"
                          f" - Error: {e}")
            sys.exit(1)
        workbook.close()
        # df.to_excel(f'./nomenclatures_{preDay.tm_year}_{preDay.tm_mon}_{preDay.tm_mday}.xlsx')
        os.remove(f'{x}.xlsx')
        doc = open(f'{x}_{ST.a}.xlsx', 'rb')
        update.callback_query.message.delete()
        update.callback_query.message.reply_document(doc)
        doc.close()
        os.remove(f'{x}_{ST.a}.xlsx')

    def chat_cancel(self, update: Update, context: CallbackContext):
        update.callback_query.message.reply_text(
            'Ок. Забыли. Это останется между нами...'
        )
        return ConversationHandler.END

    def run(self):
        logging.info('Starting bot...')
        self.bot.start_polling()
        self.bot.idle()


if __name__ == "__main__":
    bot = bot()
    bot.run()
