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

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--user", "-u", type=str, action="store", default="alexander.channui@rpmliving.com")
    parser.add_argument("--password", "-p", type=str, action="store", default=os.getenv("ZENDESK_PASSWORD"))
    args = parser.parse_args()


    # session = requests.Session()
    session = requests_cache.CachedSession('cache')
    session.auth = (args.user, args.password)
    pprint(session.auth)
    next_url = "https://roscoeproperties.zendesk.com/api/v2/macros.json?per_page=200&include=usage_30d"
    groups_url = "https://roscoeproperties.zendesk.com/api/v2/groups"
    
    macros = get_macro_list(session, next_url)

    groups = get_macro_list(session, groups_url, "groups")
    pprint(len(groups))
    group_map = {item["id"]:item["name"].replace("/", "") for item in groups if not item["deleted"]}

    active_macros = [item for item in macros if item["active"]]
    pprint(len(active_macros))
    grouped_macros = dict()
    for macro in active_macros:
        resriction = macro["restriction"]
        if resriction is None:
            print(f"restriction is null, macro id number: {macro['id']}")
            continue
        if resriction["type"] != "Group":
            print(f"restriction not group, macro id number: {macro['id']}")
            continue
        for id in resriction['ids']:
            if id not in grouped_macros:
                grouped_macros[id] = []
            grouped_macros[id].append(macro)


    wb = Workbook()
    for id, macro_list in grouped_macros.items():
        print(group_map[id])
        ws1 = wb.create_sheet(group_map[id])
        ws1.append(["Name", "ID", "Created", "Updated", "Group", "Usage 30 Days", "Action", "Action Taken"])
        for macro in macro_list:
            macro_groups = ",".join (group_map[item] for item in macro['restriction']['ids'])
            ws1.append([macro["title"], macro["id"], macro["created_at"], macro["updated_at"], macro_groups, macro["usage_30d"]])

    wb.save("macro_review.xlsx")

    
    pprint('done')
    print(len(macros))


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
