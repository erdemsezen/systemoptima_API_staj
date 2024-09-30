import xlsxwriter
import requests
import json
import pandas as pd
from datetime import datetime
API_KEY = ''

# Initialize coordinates for places.
arnavutkoy = {"point": {"latitude": 41.19727068753848, "longitude": 28.741250438999806}}
avcilar = {"point": {"latitude": 41.0405730168935, "longitude": 28.709556453751343}}
bagcilar = {"point": {"latitude": 41.04906698065621, "longitude": 28.834664191108487}}
bahcelievler = {"point": {"latitude": 40.99622873131968, "longitude": 28.849846111455278}}
bakirkoy = {"point": {"latitude": 40.9780617215533, "longitude": 28.81685552211479}}
basaksehir = {"point": {"latitude": 41.097783827380496, "longitude": 28.718937090595055}}
bayrampasa = {"point": {"latitude": 41.05408219740858, "longitude": 28.898039091907066}}
beylikduzu = {"point": {"latitude": 40.99127043272676, "longitude": 28.630840985557573}}
buyukcekmece = {"point": {"latitude": 41.03462695354059, "longitude": 28.479472705306918}}
catalca = {"point": {"latitude": 41.14841927953814, "longitude": 28.45566521943994}}
esenler = {"point": {"latitude": 41.084453616022, "longitude": 28.83187709167017}}
esenyurt = {"point": {"latitude": 41.04181264339045, "longitude": 28.646968895958135}}
fatih = {"point": {"latitude": 41.01873502537419, "longitude": 28.94151974810067}}
gungoren = {"point": {"latitude": 41.02164905246183, "longitude": 28.868242550288777}}
kucukcekmece = {"point": {"latitude": 41.01704183130148, "longitude": 28.769507808950355}}
silivri = {"point": {"latitude": 41.076038630448764, "longitude": 28.18535640071609}}
zeytinburnu = {"point": {"latitude": 40.99069859200749, "longitude": 28.891346921440512}}

adalar = {"point": {"latitude": 40.876854086690805, "longitude": 29.084977635239333}}
atasehir = {"point": {"latitude": 40.98708143276146, "longitude": 29.11883299281788}}
kadikoy = {"point": {"latitude": 40.98293492618971, "longitude": 29.056015520099415}}
kartal = {"point": {"latitude": 40.91177361196957, "longitude": 29.189872707846032}}
maltepe = {"point": {"latitude": 40.945044626611754, "longitude": 29.1328822311992}}
pendik = {"point": {"latitude": 40.9806590827888, "longitude": 29.36478844387813}}
sancaktepe = {"point": {"latitude": 40.99586989782604, "longitude": 29.21455860606777}}
sultanbeyli = {"point": {"latitude": 40.97959457326417, "longitude": 29.270938595515396}}
sile = {"point": {"latitude": 41.16298630377569, "longitude": 29.58643091043711}}
tuzla = {"point": {"latitude": 40.88046739251423, "longitude": 29.356263237489113}}
umraniye = {"point": {"latitude": 41.03374837005961, "longitude": 29.101402980385302}}
uskudar = {"point": {"latitude": 41.01995740305791, "longitude": 29.03816892255646}}

avrasya = {"point": {"latitude": 41.03655488413734, "longitude": 29.26317930221558}}
temmuz15 = {"point": {"latitude": 40.904269071036985, "longitude": 29.306421875953674}}
fsm = {"point": {"latitude": 40.985579711382854, "longitude": 29.02375996112824}}

# Indicate origins, destinations and waypoints. Also indicate their names.
origins = [arnavutkoy, avcilar, bagcilar, bahcelievler, bakirkoy, basaksehir, bayrampasa, beylikduzu, buyukcekmece, catalca, esenler, esenyurt, fatih, gungoren, kucukcekmece, silivri, zeytinburnu]
destinations = [adalar, atasehir, kadikoy, kartal, maltepe, pendik, sancaktepe, sultanbeyli, sile, tuzla, umraniye, uskudar]
waypoints = [avrasya, temmuz15, fsm]
originNames = ["ARNAVUTKÖY", "AVCILAR", "BAĞCILAR", "BAHÇELİEVLER", "BAKIRKÖY", "BAŞAKŞEHİR", "BAYRAMPAŞA", "BEYLİKDÜZÜ", "BÜYÜKÇEKMECE", "ÇATALCA", "ESENLER", "ESENYURT", "FATİH", "GÜNGÖREN", "KÜÇÜKÇEKMECE", "SİLİVRİ", "ZEYTİNBURNU"]
destinationNames = ["ADALAR", "ATAŞEHİR", "KADIKÖY", "KARTAL", "MALTEPE", "PENDİK", "SANCAKTEPE", "SULTANBEYLİ", "ŞİLE", "TUZLA", "ÜMRANİYE", "ÜSKÜDAR"]
waypointNames = ["Avrasya Tuneli", "15 Temmuz Sehitler Koprusu", "Fatih Sultan Mehmet Koprusu"]

numberOfWaypoints = len(waypoints)
numberOfOrigins = len(origins)
numberOfDestinations = len(destinations)

# Select Departure time
departureTime = "2024-09-25T07:00:00+03:00" # Write now for the current time YYYY-MM-DDTHH:mm:ss+03:00

# Initialize excel file name
if (departureTime == "now"):
    dt = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
else:
    dt = datetime.fromisoformat(departureTime).strftime("%Y-%m-%d_%H-%M-%S")

# Initialize options
options = {
    "departAt": departureTime,
    "travelMode": "car",
    "traffic": "historical" # departureTime has to be "now" to use "live" for this option
}
body1 = {
    "origins": origins,
    "destinations": waypoints,
    "options": options
}
body2 = {
    "origins": waypoints,
    "destinations": destinations,
    "options": options
}
headers = {'Content-Type': 'application/json'}

# Request data from the API
call1 = requests.post('https://api.tomtom.com/routing/matrix/2?key='+API_KEY, json=body1, headers=headers)
call2 = requests.post('https://api.tomtom.com/routing/matrix/2?key='+API_KEY, json=body2, headers=headers)

print(call1.status_code, call1.reason)
print(call1.text)
print(call2.status_code, call2.reason)
print(call2.text)


# index = origin * numberOfWaypoints + waypoint
routes1 = pd.json_normalize(json.loads(call1.text), "data")
# index = waypoint * numberOfDestinations + destination
routes2 = pd.json_normalize(json.loads(call2.text), "data")

# Process Requested data

excelMatrix = []

i = 0
for firstHalf in routes1.itertuples(index=True):
    for j in range(numberOfDestinations):
        origin = firstHalf[1]
        midpoint = firstHalf[2]
        destination = routes2.loc[i*3+j, 'destinationIndex']
        duration = firstHalf[4] + routes2.loc[i*3+j, 'routeSummary.travelTimeInSeconds']
        distance = firstHalf[3] + routes2.loc[i*3+j, 'routeSummary.lengthInMeters']
        calculatedValue = [originNames[origin], destinationNames[destination], waypointNames[midpoint], duration, distance]
        excelMatrix.append(calculatedValue)
    if (i == numberOfWaypoints-1):
        i = 0
    else:
        i += 1

# Write into an excel file
filename = "TomTom_" + dt + ".xlsx"
wb = xlsxwriter.Workbook(filename)
ws = wb.add_worksheet('Data Sheet')

ws.write(0, 0, "Origin")
ws.write(0, 1, "Destination")
ws.write(0, 2, "Midpoint")
ws.write(0, 3, "Duration (s)")
ws.write(0, 4, "Distance (m)")

index = 1 
for row in excelMatrix:
    for item in range(5):
        ws.write(index, item, row[item])
    index += 1

wb.close()   