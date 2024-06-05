import argparse
import json
import os
import pathlib
import sys

from pathlib import Path
from pprint import pprint

import openpyxl
import requests
import requests_cache

from openpyxl import workbook

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--user", "-u", type=str, action="store", default="alexander.channui@rpmliving.com")
    parser.add_argument("--password", "-p", type=str, action="store", default=os.getenv("ZENDESK_PASSWORD"))
    args = parser.parse_args()


    # session = requests.Session()
    session = requests_cache.CachedSession('cache')
    session.auth = (args.user, args.password)
    pprint(session.auth)
    next_url = "https://roscoeproperties.zendesk.com/api/v2/macros.json?per_page=200"
    
    macros = []
    while next_url:
        pprint(next_url)
        response = session.get(next_url)
        response.raise_for_status()
        data = response.json()
        # print(json.dumps(response.json(), indent=2))
        next_url = data['next_page']
        macros.extend(data['macros'])
    
    pprint('done')
    print(len(macros))


    

if __name__ == "__main__":
    main()
