import pandas as pd
import json
from pathlib import Path

def main():
    file_path_excel = Path(__file__).with_name("Werbestandorte.xlsx")
    file_path_json = Path(__file__).with_name("standort_daten.json")

    generateJson(file_path_excel, file_path_json)

    print("Json data has been written into standort_daten.json")


def generatePopup(row): 
    link_start = "<img src= https://www.wtm-aussenwerbung.de/wp-content/uploads/"
    link_end =  " style='width: 15vw; min-width: 200px;'>"

    popupText = "<h3>" + row["Name"] + "</h3><br> Werbeträger: " + row["Werbetraeger"] + "<br>Ort: " + row["Ort"] + "<br>Standort: " + row["Standort"] + "<br>Maße: " + row["Maße"] + "<br>Beleuchtet: " + row["Beleuchtung"] + "<br>Buchungsinterball: " + row["Buchungsintervall"] + "<br>Vorlaufzeit: " + row["Vorlaufzeit"] + "<br><hr><br>"

    popupString = ""
    popupString += "<div class='mCustomScrollbar' data-mcs-theme='rounded-dark'>"
    popupString += "<div class='cf'>"
    popupString += '<div>'
    popupString += popupText
    popupString += '</div>'
    popupString += '<div>'
    if row["Bild1"]: popupString += link_start + row["Bild1"] + link_end
    if row["Bild2"]: popupString += link_start + row["Bild2"] + link_end
    popupString += link_start + "wtm-aussenwerbung.webp" + link_end
    popupString += '</div>'
    popupString += '</div>'
    return popupString

def generateJson(file_path_excel, file_path_json):
    data = {"Großflächen": {"type": "FeatureCollection", "features": []}}


    excel = pd.read_excel(file_path_excel)
    excel.fillna('',inplace=True)

    for index, row in excel.iterrows():
        feature = {"type": "Feature","properties" : {"Name": "","popupContent": "","Category": ""},"geometry": {"type": "Point","coordinates": ""}}
        
        category = row["Werbetraeger"]
        if category not in data:
            data[category] = {"type": "FeatureCollection", "features": []}

        feature["properties"]["Name"] = row["Name"]
        feature["properties"]["popupContent"] = generatePopup(row)
        feature["properties"]["Category"] = category
        feature["geometry"]["coordinates"] = [float(x) for x in row["Koordinaten"].split(',')]

        data[category]["features"].append(feature)

    with open(file_path_json, 'w') as json_file:
        json.dump(data, json_file, indent=4)


main()