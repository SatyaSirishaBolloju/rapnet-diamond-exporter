"""
RapNet Diamond Data Fetcher & Excel Exporter
Author: [Your Name]
Description: Automates fetching filtered diamond listing data from RapNet API and
exports it to Excel sheets, each based on a shape-color-clarity combination.
"""

import requests
from jproperties import Properties
import json
import pandas as pd
import os
import datetime
from openpyxl import load_workbook

# Load properties from external config file
configs = Properties()

# Global columns to extract from RapNet response
columns = [
    'seller.companyName', 'location.countryCode', 'shape', 'displaySize',
    'color', 'clarity', 'cut', 'polish', 'symmetry', 'displayFluorescence',
    'displayPrice.displayPricePerCarat', 'displayPrice.displayListDiscount',
    'displayPrice.displayTotalPrice', 'displayDepthPercent', 'displayTablePercent',
    'displayMeasurments', 'shade', 'displayInclusions', 'displayLabComment',
    'displayKeyToSymbols', 'memberComment', 'sellerID'
]

def loadProperties():
    """Load token and filter inputs from config file."""
    with open('market_input_sample.txt', 'rb') as read_prop:  
        configs.load(read_prop)

def getPropertyValue(key):
    """Fetch a config value by key."""
    return configs.get(key).data

def getNameList():
    """Retrieve the list of saved searches from RapNet."""
    token = getPropertyValue("token")  
    url = 'https://api.example.com/savesearch/names'
    headers = {'Authorization': f'Bearer {token}'}

    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        return response.json()["data"]["namesList"]
    except Exception as e:
        print(f"Failed to get names list: {e}")
        return None

def getFilterCriteria(namesListId, size_range, color, clarity):
    """Fetch search filter parameters for a saved search."""
    token = getPropertyValue("token")
    url = "https://api.example.com/savesearch/list"
    headers = {'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'}
    payload = {"pageNumber": 1, "recordsPerPage": 1, "savedSearchIDs": [namesListId]}

    try:
        response = requests.post(url, headers=headers, json=payload)
        response.raise_for_status()
        filters = response.json()["data"]["savedSearchList"][0]["filter"]

        if size_range:
            size_from, size_to = size_range.split(':')
            filters['size'] = {
                'isSpecificSize': True,
                'sizeGrids': [f"{size_from} - {size_to}"],
                'sizeFrom': size_from.strip(),
                'sizeTo': size_to.strip()
            }

        filters['color'] = {'isWhiteColor': True, 'colorFrom': color, 'colorTo': color}
        filters['clarity'] = {'clarityFrom': clarity, 'clarityTo': clarity}

        diamond_attrs = {
            'Shape': filters.get('shape', {}).get('shapes', [''])[0],
            'Color': f"{color}",
            'Size': f"{size_from} - {size_to}" if size_range else '',
            'Clarity': f"{clarity}",
            'Fluorescence Intensity': ', '.join(filters.get('fluorescence', {}).get('fluorescenceIntensities', [])),
            'Depth': f"{filters.get('depth', {}).get('depthPercentFrom', '')} - {filters.get('depth', {}).get('depthPercentTo', '')}",
            'Table': f"{filters.get('table', {}).get('tablePercentFrom', '')} - {filters.get('table', {}).get('tablePercentTo', '')}",
            'Labs': ', '.join(filters.get('labs', [])),
            'Cut': f"{filters.get('finish', {}).get('cutFrom', '')} - {filters.get('finish', {}).get('cutTo', '')}",
            'Polish': f"{filters.get('finish', {}).get('polishFrom', '')} - {filters.get('finish', {}).get('polishTo', '')}",
            'Symmetry': f"{filters.get('finish', {}).get('symmetryFrom', '')} - {filters.get('finish', {}).get('symmetryTo', '')}",
            'Crown Height': f"{filters.get('crown', {}).get('crownHeightFrom', '')} - {filters.get('crown', {}).get('crownHeightTo', '')}",
            'Crown Angle': f"{filters.get('crown', {}).get('crownAngleFrom', '')} - {filters.get('crown', {}).get('crownAngleTo', '')}",
            'Pavilion Angle': f"{filters.get('pavilion', {}).get('pavilionAngleFrom', '')} - {filters.get('pavilion', {}).get('pavilionAngleTo', '')}"
        }

        return filters, diamond_attrs

    except Exception as e:
        print(f"Error in filter criteria: {e}")
        return None, None

def aggregateCounts(body):
    """Get total diamond count from search filters."""
    token = getPropertyValue("token")
    url = "https://api.example.com/diamondsearch/aggregations"
    headers = {'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'}

    try:
        response = requests.post(url, headers=headers, json=body)
        response.raise_for_status()
        return response.json()["data"]["totalDiamondCount"]
    except Exception as e:
        print(f"Error aggregating counts: {e}")
        return None

def extract_nested_data(data, key):
    """Safely extract nested data from dictionary using dot notation."""
    try:
        for k in key.split('.'):
            data = data[k]
        return str(data)
    except:
        return None

def fetch_and_process_diamonds():
    """Main processing function that loops filters and fetches data."""
    loadProperties()
    saved_searches = getPropertyValue("load_saved_search").split(',')
    size_ranges = getPropertyValue("size_range").split(',')
    colors = getPropertyValue("colors").split(',')
    clarities = getPropertyValue("clarities").split(',')

    names_list = getNameList()
    data_dict = {}
    attributes_dict = {}

    for item in names_list:
        if item["name"] in saved_searches:
            for size in size_ranges:
                for color in colors:
                    for clarity in clarities:
                        filters, attrs = getFilterCriteria(item["id"], size, color, clarity)
                        if not filters: continue

                        attributes_dict[(item["name"], size, color, clarity)] = attrs

                        try:
                            with open("filter.json", "r") as f:
                                body = json.load(f)
                            body["filter"] = filters

                            count = aggregateCounts(body)
                            if not count: continue

                            url = "https://api.example.com/diamondsearch/search?start=1&size=250"
                            token = getPropertyValue("token")
                            headers = {'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'}

                            response = requests.post(url, headers=headers, json=body)
                            response.raise_for_status()
                            diamonds = response.json()["data"]["diamonds"]

                            for diamond in diamonds:
                                key = (item["name"], size, color, clarity)
                                data_dict.setdefault(key, []).append({
                                    col.split('.')[-1]: extract_nested_data(diamond, col) for col in columns
                                })

                        except Exception as e:
                            print(f"Error fetching diamonds: {e}")
    return data_dict, attributes_dict

def save_to_excel(data_dict, attributes_dict):
    """Write results to Excel with attribute headers and listings."""
    current_time = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    size_ranges = getPropertyValue("size_range").split(',')
    searches = getPropertyValue("load_saved_search").split(',')

    col_map = {
        'companyName': 'SELLER', 'countryCode': 'CTY', 'shape': 'SHAPE', 'displaySize': 'SIZE',
        'color': 'COL', 'clarity': 'CLA', 'cut': 'CUT', 'polish': 'POL', 'symmetry': 'SYM',
        'displayFluorescence': 'FLU', 'displayPricePerCarat': '$/CT', 'displayListDiscount': '%RAP',
        'displayTotalPrice': 'AMT', 'displayDepthPercent': 'TD', 'displayTablePercent': 'TB',
        'displayMeasurments': 'MEASUREMENTS', 'shade': 'SHADE', 'displayInclusions': 'INCLUSION',
        'displayKeyToSymbols': 'KEY TO SYMBOLS', 'displayLabComment': 'Lab Comment',
        'memberComment': 'MEMBER COMMENTS', 'sellerID': 'ID'
    }

    attr_map = {
        'Shape': 'SHAPE', 'Color': 'COL', 'Size': 'SIZE', 'Clarity': 'CLA',
        'Fluorescence Intensity': 'FLU', 'Depth': 'TD', 'Table': 'TB', 'Labs': 'LAB',
        'Cut': 'CUT', 'Polish': 'POL', 'Symmetry': 'SYM',
        'Crown Height': 'CH', 'Crown Angle': 'CA', 'Pavilion Angle': 'PA'
    }

    for search in searches:
        for size in size_ranges:
            filename = f"{search}{size.replace(':', '')}_{current_time}.xlsx"
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                for key, data in data_dict.items():
                    if key[0] != search or key[1] != size: continue
                    sheet = f"{search}{key[2]}{key[3]}"[:31]

                    attr_df = pd.DataFrame([attributes_dict[key]])
                    data_df = pd.DataFrame(data)
                    attr_df.rename(columns=attr_map, inplace=True)
                    data_df.rename(columns=col_map, inplace=True)

                    attr_df.to_excel(writer, index=False, sheet_name=sheet)
                    data_df.to_excel(writer, index=False, sheet_name=sheet, startrow=len(attr_df)+2)

            print(f"Saved: {filename}")

def main():
    data_dict, attributes_dict = fetch_and_process_diamonds()
    save_to_excel(data_dict, attributes_dict)

if _name_ == "_main_":
    main()