import xlsxwriter
import requests
import json
from datetime import datetime
API_KEY = ''

# coordinates are in [longitude, latitude]
levent = [29.016724526882175, 41.08147607930026]
bagcilar = [28.856792449951175, 41.03381141422386]
bakirkoy = [28.880572915077213, 40.9835974693]
ozu = [28.996811807155613, 41.00576319620463]
sabihaGokcen = [29.033315185874994, 41.0464491047922]
kadikoy = [29.061651527881626, 41.09132951805778]
avrasya = [29.26317930221558, 41.03655488413734]
temmuz15 = [29.306421875953674, 40.904269071036985]
fsm = [29.02375996112824, 40.985579711382854]

# Indicate locations in europe side, asia side and waypoints. Also indicate their names.
europeLocations = [levent, bagcilar, bakirkoy]
asiaLocations = [ozu, sabihaGokcen, kadikoy]
waypoints = [avrasya, temmuz15, fsm]
europeLocationNames = ["Levent", "Bagcilar", "Bakirkoy"]
asiaLocationNames = ["Ozyegin Universitesi", "Sabiha Gokcen Havalimani", "Kadikoy"]
waypointNames = ["Avrasya Tuneli", "15 Temmuz Sehitler Koprusu", "Fatih Sultan Mehmet Koprusu"]

locations = europeLocations + asiaLocations + waypoints
locationNames = europeLocationNames + asiaLocationNames + waypointNames

numberOfWaypoints = len(waypoints)
europeSide = len(europeLocations)
asiaSide = len(asiaLocations)
numberOfLocations = len(locations) - numberOfWaypoints

# Request Data from OpenRoute Service Matrix API
body = {"locations":locations,"metrics":["distance","duration"],"units":"km"}
headers = {
    'Accept': 'application/json, application/geo+json, application/gpx+xml, img/png; charset=utf-8',
    'Authorization': API_KEY,
    'Content-Type': 'application/json; charset=utf-8'
}
call = requests.post('https://api.openrouteservice.org/v2/matrix/driving-car', json=body, headers=headers)

# Print raw response data for debugging
print(call.status_code, call.reason)
print(call.text)
responseData = json.loads(call.text)

# Initialize date and time
dt = datetime.now()
filename = 'Openroute_' + dt.strftime('%Y-%m-%d_%H-%M-%S') + '.xlsx'

# Process requested data and write them into a matrix to use as input for excel file.
distances = responseData['distances']
durations = responseData['durations']

excelMatrix = []

for i in range(numberOfLocations):
    if (i < europeSide):
        for j in range(asiaSide):
            for k in range(numberOfWaypoints):
                origin = i
                midpoint = numberOfLocations+k
                destination = europeSide+j
                duration = durations[origin][midpoint] + durations[midpoint][destination]
                distance = distances[origin][midpoint] + distances[midpoint][destination]
                calculatedValue = [locationNames[origin], locationNames[destination], locationNames[midpoint], duration, distance]
                excelMatrix.append(calculatedValue)
    else:
        for j in range(asiaSide):
            for k in range(3):
                origin = i
                midpoint = numberOfLocations+k
                destination = j
                duration = durations[origin][midpoint] + durations[midpoint][destination]
                distance = distances[origin][midpoint] + distances[midpoint][destination]
                calculatedValue = [locationNames[origin], locationNames[destination], locationNames[midpoint], duration, distance]
                excelMatrix.append(calculatedValue)

# Write into an excel file
wb = xlsxwriter.Workbook(filename)
ws = wb.add_worksheet('Data Sheet')

ws.write(0, 0, "Origin")
ws.write(0, 1, "Destination")
ws.write(0, 2, "Midpoint")
ws.write(0, 3, "Duration (s)")
ws.write(0, 4, "Distance (km)")

index = 1 
for row in excelMatrix:
    for item in range(5):
        ws.write(index, item, row[item])
    index += 1

wb.close()