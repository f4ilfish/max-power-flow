import json
import csv


def csv_to_list(path: str) -> [dict]:
    """ Parse .сsv to list of dictionaries"""

    dict_list = []

    with open(path, newline='') as csv_data:
        csv_dic = csv.DictReader(csv_data)

        # Creating empty list and adding dictionaries (rows)
        for row in csv_dic:
            dict_list.append(row)

    return dict_list


def json_to_dic(path: str) -> dict:
    """ Parse .json to dic """

    with open(path, "r") as json_data:
        dictionary = json.load(json_data)

    return dictionary
