import json
import csv


def csv_to_list(path: str):
    """ Parse .сsv to list of dictionaries"""

    diclist = []

    with open(path, newline='') as csv_data:
        csv_dic = csv.DictReader(csv_data)

        # Creating empty list and adding dictionaries (rows)
        for row in csv_dic:
            diclist.append(row)

    return diclist


def json_to_dic(path: str):
    """ Parse .json to dic """

    with open(path, "r") as json_data:
        dictionary = json.load(json_data)

    return dictionary
