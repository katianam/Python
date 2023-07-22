import requests
import pandas as pd
import xlrd
import json
import time
import urllib.error
import urllib.parse
import urllib.request
import math
import xlwt
import json
import re

def main():

    file_name = 'C:\\Users\\Katie\\Downloads\\STORES.xls'
    base_url = "https://maps.googleapis.com/maps/api/distancematrix/json?origins="

    wb = xlrd.open_workbook(file_name)
    tractsws = wb.sheet_by_index(3)
    census_tract = tractsws.col_values(0)
    tlat = tractsws.col_values(1)
    tlong = tractsws.col_values(2)
    pop = tractsws.col_values(3)
    mpop = tractsws.col_values(4)

    storesws = wb.sheet_by_index(1)
    store = storesws.col_values(0)
    storenm = storesws.col_values(1)
    storetype = storesws.col_values(2)
    slat = storesws.col_values(3)
    slong = storesws.col_values(4)

    workbook = xlwt.Workbook()
    worksheet = workbook.add_sheet("API_output")

    r=0
    for i in range(1,tractsws.nrows):
            json_answers = list()  # = []
            coord_list = []
            min_dur = 500
            for j in range(1,storesws.nrows):
                if abs(float(tractsws.cell_value(i,1)) - float(storesws.cell_value(j,3))) < .17 and abs(float(tractsws.cell_value(i,2)) - float(storesws.cell_value(j,4))) <.17 :
                    r=r+1
                    url = base_url+str(tractsws.cell_value(i,1))+'%2C'+str(tractsws.cell_value(i,2))+'&destinations='+str(storesws.cell_value(j,3))+'%2C'+str(storesws.cell_value(j,4))+'&key=***'
                    payload={}
                    headers = {}
                    response = requests.request("GET", url, headers=headers, data=payload)
                    json_answers.append(response.text)

            for answer in json_answers:
                json_data = json.loads(answer)
                dur1 = json_data['rows'][0]['elements'][0]['duration']['text']
                dist1 = json_data['rows'][0]['elements'][0]['distance']['text']
                dur = re.findall('\d+',  dur1)
                if int(dur[0]) < min_dur:
                    min_dur = int(dur[0])
                    dest = json_data['destination_addresses'][0]
                    org = json_data['origin_addresses'][0]
                    dist = re.findall('\d+',  dist1)
                    closest_loc =  [dur1, dist1, org, dest]
                    closest_loc_detail = [str(tractsws.cell_value(i,0))]
                    continue
                temp_list = { dur1, dist1}
                coord_list.append(temp_list)
            print(closest_loc)
            worksheet.write(i, 1, closest_loc[0])
            worksheet.write(i, 2, closest_loc[1])
            worksheet.write(i, 3, closest_loc[2])
            worksheet.write(i, 4, closest_loc[3])
            worksheet.write(i, 5, closest_loc_detail[0])
            workbook.save("API_OUTPUT.xls")

main()
