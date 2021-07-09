import requests
from urllib.request import urlopen
from urllib.error import HTTPError
import json
import datetime
import zipfile
import time
import os


def getHTML(url):
    try:
        html = urlopen(url)
    except HTTPError as e:
        return None
    return html


def extractDate(d):
    d['name'] = d['name'].split(' ')[1]
    d['name'] = datetime.datetime.strptime(d['name'], '%d.%m.%Y')
    return d


def findFresh(allCuts):
    return max(allCuts, key=lambda d: d['name'])


def urlToDictList(cutURL):
    cutList = getHTML(cutURL)
    allCuts = json.loads(cutList.read().decode('utf-8'))
    allCuts = map(extractDate, allCuts)
    allCuts = list(allCuts)
    return allCuts


def findCutId(allCuts):
    return findFresh(allCuts)['id']


def extractZip(zipfile_path, root_directory):
    zayavki = root_directory
    with zipfile.ZipFile(zipfile_path, 'r') as zip_ref:

        excel_name = zip_ref.namelist()[0]
        excel_path = zayavki + '/' + excel_name
        # print(excel_path)
        excel_extension = excel_name.split(".")[1]
        # print(excel_extension)
        try:
            time_now = datetime.datetime.now()
            time_now = datetime.datetime.strftime(time_now, "%d-%m-%y_%H-%M")
            # print(time_now)
        except:
            print('problems with reading date!')
        new_excel_path = zayavki + '/request_' + time_now + '.' + excel_extension
        # print(new_excel_path)

        # with open(zayavki + '/request_name.txt', 'w') as txt:
        #   txt.write(excel_path)
        zip_ref.extractall(zayavki + '/')

        if os.path.exists(excel_path):
            os.rename(excel_path, new_excel_path)

    if os.path.exists(zipfile_path):
        os.remove(zipfile_path)


def main():
    # получаем список срезов
    allCuts = urlToDictList('https://stat.gov.kz/api/rcut/ru')
    # print(allCuts)

    # находим свежайщий cutId
    cutId = findCutId(allCuts)
    # print(cutId)

    juridicalFace = 742679

    # по г.АЛМАТЫ
    katoId = 741880

    # запрос по json
    data = {"conditions":
                [{"classVersionId": 2153, "itemIds": [juridicalFace]},
                 {"classVersionId": 213, "itemIds": [katoId]}
                 ],
            "cutId": cutId}

    # {"classVersionId": 12621, "itemIds": [307545]}

    r = requests.post(url="https://stat.gov.kz/api/sbr/request/?api", json=data)

    # номер заявки
    requestNUMBER = json.loads(r.text)['obj']
    # print(requestNUMBER)

    zayavki = 'zayavki'

    # пока заявка не обработана, пытаемся найти GUID
    while True:
        # считываем статус заявки
        status = json.loads(urlopen(f'https://stat.gov.kz/api/sbr/requestResult/{requestNUMBER}/ru').read())
        print(status)
        try:
            guid = status['obj']['fileGuid']
            # print(guid)
            # если нашли, скачиваем файл
            request = requests.get(f'https://stat.gov.kz/api/sbr/download?bucket=SBR&guid={guid}')
            a = request.headers["Content-Disposition"].split('filename=')[1].strip("\"")
            # request_name,extension=a.split('.')
            # print(request_name.strip("\""))

            if not os.path.exists(zayavki):
                os.makedirs(zayavki)

            zipfile_path = zayavki + '/' + a

            with open(zipfile_path, 'wb') as file:
                file.write(request.content)

            # разархивировать xlsx
            extractZip(zipfile_path, zayavki)

            # выходим из цикла
            break
        except KeyError:
            print('Key error!')
        except TypeError:
            print('Not ready')
            time.sleep(1)
        else:
            print('Something went wrong!')
            break

    # print(guid)

    # d=[{"id":677,"name":"На 01.05.2022 года"},{"id":677,"name":"На 01.06.2021 года"},{"id":674,"name":"На 01.05.2021 года"},{"id":671,"name":"На 01.04.2021 года"},{"id":668,"name":"На 01.03.2021 года"},{"id":662,"name":"На 01.02.2021 года"},{"id":658,"name":"На 01.01.2021 года"},{"id":655,"name":"На 01.12.2020 года"},{"id":652,"name":"На 01.11.2020 года"},{"id":649,"name":"На 01.10.2020 года"},{"id":646,"name":"На 01.09.2020 года"},{"id":643,"name":"На 01.08.2020 года"},{"id":640,"name":"На 01.07.2020 года"},{"id":637,"name":"На 01.06.2020 года"},{"id":634,"name":"На 01.05.2020 года"},{"id":631,"name":"На 01.04.2020 года"},{"id":628,"name":"На 01.03.2020 года"},{"id":625,"name":"На 01.02.2020 года"},{"id":618,"name":"На 01.01.2020 года"},{"id":615,"name":"На 01.12.2019 года"},{"id":612,"name":"На 01.11.2019 года"},{"id":609,"name":"На 01.10.2019 года"},{"id":606,"name":"На 01.09.2019 года"},{"id":603,"name":"На 01.08.2019 года"},{"id":600,"name":"На 01.07.2019 года"},{"id":597,"name":"На 01.06.2019 года"},{"id":594,"name":"На 01.05.2019 года"},{"id":591,"name":"На 01.04.2019 года"},{"id":587,"name":"На 01.03.2019 года"},{"id":585,"name":"На 01.02.2019 года"},{"id":580,"name":"На 01.01.2019 года"},{"id":577,"name":"На 01.12.2018 года"},{"id":574,"name":"На 01.11.2018 года"},{"id":571,"name":"На 01.10.2018 года"},{"id":567,"name":"На 01.09.2018 года"},{"id":564,"name":"На 01.08.2018 года"},{"id":561,"name":"На 01.07.2018 года"},{"id":558,"name":"На 01.06.2018 года"},{"id":555,"name":"На 01.05.2018 года"},{"id":552,"name":"На 01.04.2018 года"},{"id":547,"name":"На 01.03.2018 года"},{"id":544,"name":"На 01.02.2018 года"},{"id":538,"name":"На 01.01.2018 года"}]

    # print(extractDate(d[0]))


if __name__ == "__main__":
    main()
