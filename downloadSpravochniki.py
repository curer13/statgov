import requests
import os
import openpyxl


def getExtension(request):
    extension = request.headers["Content-Disposition"].split('filename=')[1].strip("\"").split('.')[1]
    return extension


spravocniki_name='spravochniki'

if not os.path.exists(spravocniki_name):
    os.makedirs(spravocniki_name)

with open(spravocniki_name+'/2153.xlsx', 'wb') as file:
    file.write(requests.get(f'https://stat.gov.kz/api/getFile/?docId=ESTAT346269').content)

with open(spravocniki_name+'/213.xlsx', 'wb') as file:
    file.write(requests.get(f'https://stat.gov.kz/api/getFile/?docId=ESTAT346264').content)

with open(spravocniki_name+'/4855.xlsx', 'wb') as file:
    file.write(requests.get(f'https://stat.gov.kz/api/getFile/?docId=ESTAT346270').content)

with open(spravocniki_name+'/15.xlsx', 'wb') as file:
    file.write(requests.get(f'https://stat.gov.kz/api/getFile/?docId=ESTAT346260').content)

with open(spravocniki_name+'/1989.xlsx', 'wb') as file:
    file.write(requests.get(f'https://stat.gov.kz/api/getFile/?docId=ESTAT346268').content)

with open(spravocniki_name+'/17.xlsx', 'wb') as file:
    file.write(requests.get(f'https://stat.gov.kz/api/getFile/?docId=ESTAT346262').content)

# Can't get the downloading file name!!!
request12621=requests.get(f'https://stat.gov.kz/api/getFile/?docId=ESTAT346266')
extension12621=getExtension(request12621)
# print(request12621.headers["Content-Disposition"])
with open(spravocniki_name+'/12621.'+extension12621, 'wb') as file:
    file.write(request12621.content)

