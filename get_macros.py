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

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.descriptors import Integer

def main():
    # argument parsing
    parser = argparse.ArgumentParser()
    parser.add_argument("--user", "-u", type=str, action="store", default=os.getenv("ZENDESK_USERNAME"))
    parser.add_argument("--password", "-p", type=str, action="store", default=os.getenv("ZENDESK_PASSWORD"))
    args = parser.parse_args()


    # request macros from zendesk
    # session = requests.Session()
    session = requests_cache.CachedSession('cache')
    session.auth = (args.user, args.password)
    # print debugging
    # print(session.auth)
    next_url = "https://roscoeproperties.zendesk.com/api/v2/macros.json?per_page=200&include=usage_30d"
    groups_url = "https://roscoeproperties.zendesk.com/api/v2/groups"
    
    # getting all macros and narrowing down to all active macros
    macros = get_macro_list(session, next_url)
    active_macros = [item for item in macros if item["active"]]
    # print debugging
    # print(len(active_macros))

    # getting all group ids
    groups = get_macro_list(session, groups_url, "groups")
    # print debugging
    # print(len(groups))
    # dict of group id and corresponding name
    group_map = {item["id"]:item["name"].replace("/", "") for item in groups if not item["deleted"]}

    # creates dict of macros by group id
    grouped_macros = dict()
    sort_macros(active_macros, grouped_macros)


    # creates workbook with each group having their own sheet
    wb = Workbook()
    create_workbook(group_map, grouped_macros, wb)
        



    #print debugging
    # print(wb.get_sheet_names())
    del wb['Sheet']
    # print debugging
    # print(wb.get_sheet_names())
    wb.save("macro_review.xlsx")

    # print debugging
    # print(len(macros))
    print('done')

def sort_macros(active_macros, grouped_macros):
    for macro in active_macros:
        resriction = macro["restriction"]
        # filtering out null and non Group restrictions
        if resriction is None:
            print(f"restriction is null, macro id number: {macro['id']}")
            continue
        if resriction["type"] != "Group":
            print(f"restriction not group, macro id number: {macro['id']}")
            continue
        # creates new 
        for id in resriction['ids']:
            if id not in grouped_macros:
                grouped_macros[id] = []
            grouped_macros[id].append(macro)

def create_workbook(group_map, grouped_macros, wb):
    for id, macro_list in grouped_macros.items():
        # sheet creation
        # print debugging
        # print(group_map[id])
        ws1 = wb.create_sheet(group_map[id])
        # header
        ws1.append(['Name', "ID", "Created", "Updated", "Group", "Usage 30 Days", "Action", "Action Taken"])

        #set font
        font = Font(bold=True, underline='single')
        for cell in ws1['1:1']:
            cell.font = font

        # fill sheet with macros from dict
        for macro in macro_list:
            macro_groups = ",".join (group_map[item] for item in macro['restriction']['ids'])
            ws1.append([macro["title"], macro["id"], macro["created_at"], macro["updated_at"], macro_groups, macro["usage_30d"]])
        
        # set id to show in regular form not scientific notation
        colB = ws1['B']
        for cell in colB:
            cell.number_format = '0'
        ws1.column_dimensions['B'].hidden=True



# grabs all data from zendesk from url with key to specify what to grab
def get_macro_list(session, next_url, key="macros"):
    results = []
    while next_url:
        pprint(next_url)
        response = session.get(next_url)
        response.raise_for_status()
        data = response.json()
        # print(json.dumps(response.json(), indent=2))
        next_url = data['next_page']
        results.extend(data[key])
    return results


    

if __name__ == "__main__":
    main()
