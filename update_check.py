#! /usr/bin/python3

import json
import requests
from datetime import datetime
import time

UPDATE_URL = 'https://bitbucket.org/api/2.0/repositories/compass_dataservices/agilysys-import-export-tools/downloads'

response = requests.get(UPDATE_URL)
data = response.json()
recent_update = data['values'][0]
last_update_date = recent_update['created_on']
update_name = recent_update['name']
download_link = recent_update['links']['self']['href']

last_update_year = int(last_update_date[:4])
last_update_month = int(last_update_date[5:7])
last_update_day = int(last_update_date[8:10])
last_update_hour = int(last_update_date[11:13])
last_update_minute = int(last_update_date[14:16])
last_update_dt = datetime(last_update_year, last_update_month, last_update_day, last_update_hour, last_update_minute)

current_dt = datetime(2016, 1, 14, 20 , 1)

if current_dt < last_update_dt:
    print('new update available')
else:
    print('running latest version')

print('Most recent update published on {0}.'.format(last_update_date))
print('Download: {0} \nLink: {1}'.format(update_name, download_link))


def download_update(url, name):
    update = requests.get(url)
    with open(name, 'wb+') as file:
        file.write(update.content)

    print('Update downloaded successfully')

