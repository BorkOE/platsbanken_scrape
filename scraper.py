'''
PLATSBANKEN API SCRAPER

- Hämta annonser
    - Obs, hämtar max 100 annonser per söktext
- Skriv annonser till excel
    - Markera nya annonser grönt, gamla gult
'''

import requests
import datetime
import pandas as pd
import numpy as np
import os
from excel_printer import Printer
from settings import söktexter, URL, BASLÄNK, excel_path, minne_path, color_dict

def get_ads(söktext):
    payload = {"filters":[
        {
            "type":"freetext",
            "value":söktext,
        },
        ],
            "fromDate":None,
            "order":"relevance",
            "maxRecords":100,
            "startIndex":0,
            "toDate":datetime.datetime.now().isoformat()[:-3] + 'Z',
            "source":"pb"}
    response = requests.post(URL, json=payload)

    # respons som dict
    res = eval(response.text
        .replace('true', 'True')
        .replace('false', 'False')
        )

    return res.get('ads', [])

def get_result_frame(söktexter):
    res_list = []
    for söktext in söktexter:
        ads = get_ads(söktext)
        if ads:
            temp_res = pd.DataFrame(ads)
            temp_res['söktext'] = söktext
            res_list.append(temp_res)

    res = pd.concat(res_list, axis=0).reset_index(drop=True)
    res[''] = np.nan    # Ska hålla färg

    # Enkel tvätt
    res = res[[
        '', 'id', 'title', 'occupation', 'workplace', 'workplaceName',
        'positions', 'lastApplicationDate', 'publishedDate', 'söktext', 
        ]]
    res = res.sort_values('publishedDate', ascending=False)

    res['lastApplicationDate'] = pd.to_datetime(res['lastApplicationDate']).dt.strftime(r'%Y-%m-%d')
    res['publishedDate'] = pd.to_datetime(res['publishedDate']).dt.strftime(r'%Y-%m-%d %H:%M')
    res = res.drop_duplicates('id')
    res['länk'] = BASLÄNK + res['id'] + ' '

    return res.reset_index(drop=True)

def write_id_to_memory(frame):
    if not os.path.exists(minne_path):
        frame['id'].astype(str).to_csv(minne_path, index=False)
    else:
        data = frame.loc[frame[''] == 'grön', 'id']
        if data.empty:
            return
        data.to_csv(minne_path, index=False,
                          mode='a', header=False)

def get_memory():
    if not os.path.exists(minne_path):
        return []
    return pd.read_csv(minne_path)['id'].astype(str).to_list()

res = get_result_frame(söktexter)
mem_id = get_memory()
res.loc[res['id'].isin(mem_id), ''] = 'gul' # Igenkända id får gul färg
res[''] = res[''].fillna('grön')        # Alla andra grön
write_id_to_memory(res)

# Print
printer = Printer(excel_path)
printer.append(res, 'jobb', {'A':4, 'B':10, 'C': 40, 'D':30, 'E': 15, 
                             'F': 35, 'G':8, 'H:I':15, 'J':20, 'K':30}, 
                index=False, wrap_values=True,
                hyperlink_cols=['K'],
                color_dict=color_dict
                    )
printer.run()

