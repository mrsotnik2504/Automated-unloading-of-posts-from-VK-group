import requests
import time
import xlrd, xlwt
from datetime import datetime
# --------------------Функция получения спарсеных данных--------------------
def request_vk_api(token, version, domain, count, offset, attempt):
    all_posts = []
    for i in range(attempt):
        response = requests.get('https://api.vk.com/method/wall.get', 
                                params={'access_token': token,
                                        'v': version,
                                        'domain': domain,
                                        'count': count,
                                        'offset': offset,
                                        }
                                        )
        data = response.json()['response']['items']
        offset += 100
        all_posts.extend(data)
        time.sleep(0.5)
    return all_posts
# ----------------------------------------------------------------
# --------------------Функция загрузки данных в эксель--------------------
def write_excel(data): 
    """Функция для записи данных в таблицу Excel"""
    # Создаем документ
    wb = xlwt.Workbook() 
    # Добавляем лист к документу
    sheet = wb.add_sheet('Sheet 1') 
    # Записываем данные из списка
    for i, row in enumerate(data): 
        for j, col in enumerate(row): 
            sheet.write(i, j, col) 
    # Сохраняем файл
    wb.save("test.xls") # НАИМЕНОВАНИЕ ФАЙЛА ДЛЯ СОХРАНИНЕНИЯ
# ----------------------------------------------------------------
# --------------------Данные для настройки--------------------
x = [['Дата', 'Текст', 'Картинка']] # Шапка для экселя (первая строка)
token = 'ВАШ ТОКЕН' # токен вк апи
version = '5.131' # версия вк апи
domain = 'НАЗВАНИЕ ГРУППЫ ИЗ ССЫЛКИ' # Ссылка на группу в вк которую планируем спарсить
count = 100 # Примерное количество постов
offset = 0 # с какого поста начинаем
attempt = 2 # Количество попыток (за одну попытку захватывает 100 постов)


# ----------------------------------------------------------------
# --------------------сбор данных в одну переменную--------------------
data = request_vk_api(token, version, domain, count, offset, attempt)
# ----------------------------------------------------------------
# --------Упаковка файлов в подходящем формате для экселя [[],[],[]]--------
for i in range(len(data)):
    date = int(data[i]['date'])
    date = datetime.utcfromtimestamp(date).strftime('%Y-%m-%d')
    txt = data[i]['text']
    # Условия на ошибку так как фото приходят в разных форматах
    try:
        foto = data[i]['attachments'][0]['doc']['url']
    except:
        try:
            foto = data[i]['attachments'][0]['photo']['sizes'][7]['url']
        except:
            try:
                foto = data[i]['attachments'][0]['photo']['sizes'][7]['url']
            except:
                try:
                    foto = data[i]['attachments'][0]['video']['image'][0]['url']
                except:
                    foto = 'Нет картинки'
                    try:
                        foto = data[i]['copy_history'][0]['attachments'][0]['photo']['sizes'][7]['url']
                    except:
                        foto = 'Нет картинки'
    x.append([date, txt, foto])
# ----------------------------------------------------------------
# print(data[0]['attachments'][0]['video']['image'][2]['url'])
# print(data[111]['copy_history'][0]['attachments'][0]['photo']['sizes'][7]['url'])
# --------------Запуск функции по записи в эксель-----------
write_excel(x)
# ----------------------------------------------------------