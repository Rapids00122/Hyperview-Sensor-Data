import pandas as pd
import requests
from datetime import datetime

# Constants
TOKEN_URL = 'https://msu.hyperviewhq.com/connect/token'
SENSOR_API_URL = 'https://msu.hyperviewhq.com/api/asset/sensors'
SUMMARY_API_URL = 'https://msu.hyperviewhq.com/api/asset/sensorsDailySummaries/numeric/last7Days'
HUMIDITY_TYPE_ID = '466799ea-0e25-e211-8183-001c42e521d8'
TEMPERATURE_TYPE_ID = '52835710-56f9-4311-babb-67b21b423c7d'
HEADERS = {
    'accept': 'application/json',
    'Authorization': None  # Will be set after fetching the token
}

# Utility Functions 
def extract_rack_name(PDUName):
    '''Extract rack name from PDU Name.'''
    parts = PDUName.split('-')
    return parts[1]

def extract_date(timestamp):
    '''Extract date from ISO format tiemstamp.'''
    dt = datetime.fromisoformat(timestamp)
    return dt.date().strftime('%Y-%m-%d')

def get_access_token():
    """Get access token for API requests."""
    response = requests.post(TOKEN_URL, data={
        'client_id': '8dcaf879-9077-42b1-a7b3-05dbce43c5b0',
        'client_secret': '08c1b6ec-9800-42d6-ada1-be2244e76e38',
        'grant_type': 'client_credentials'
    }, headers={'Content-Type': 'application/x-www-form-urlencoded'})
    response.raise_for_status()
    return response.json()['access_token']

# Step 1: Generate Hyperview Access Token
headers = {
    'Content-Type': 'application/x-www-form-urlencoded'
}

body = {
    'client_id': '8dcaf879-9077-42b1-a7b3-05dbce43c5b0',
    'client_secret': '08c1b6ec-9800-42d6-ada1-be2244e76e38',
    'grant_type': 'client_credentials'
}

response = requests.post('https://msu.hyperviewhq.com/connect/token', data=body, headers=headers)

content = response.json()

#Access Token
token = content['access_token']

# Step 2: Send API request to Hyperview
headers = {
    'accept': 'application/json',
    'accept-language': 'en-US,en;q=0.9',
    'content-type': 'application/json',
    'origin': 'https://msu.hyperviewhq.com',
    'priority': 'u=1, i',
    'Authorization': f'Bearer {token}'
}

#List of all unique rack IDs with sensors
racksWithSensors = [
    "127e5e7e-3583-4a1c-9085-298f0227a505",
    "688f6ac3-dc01-44e2-a3aa-448acd6d92ca",
    "80d75e0c-558b-4c5e-afb2-a34501441b04",
    "2a5c3f17-2fc2-4c96-9715-8e1dc8d6fae9",
    "8ddcebc2-0212-4a1e-91e8-9c6552cbd592",
    "02e9aa2c-c2ce-4921-887f-0591563e66f8",
    "80185bcb-f9ec-498b-8824-bfe8066301d0",
    "cbf43771-8874-46ff-9c7b-419f4bbfd9fc",
    "3d2c9201-66f3-470b-830d-0dac74ba713d",
    "e9ba155b-d6cf-4903-b413-064c9516971e",
    "668bf286-099c-491c-af2c-582097328088",
    "16a6a582-f273-4b31-8d6b-6d4ef93c4a5a",
    "0fa28de8-c9e2-4479-ad1e-c099168ee692",
    "7126e2f2-e0e7-4848-82fe-13b1cbfcaf80",
    "63026f9d-6eaa-429d-98fd-71e616896809",
    "2c993106-9c06-40ca-87d0-a3fd5bddfafc",
    "d0851c6b-3fce-4bb5-bc2c-1b7ed57af00f",
    "bcd8590c-e40a-49b6-a9f8-9a6c991d1d8d",
    "84a4b1ec-b523-4c3a-9a36-da4b5e652208",
    "42e4a5b9-1a63-4bc5-84cd-51be2cdf9b91",
    "312ab6c2-e78c-43d8-92c0-e625ec927097",
    "bbd75b94-88da-4033-b171-83e0b26eab3a",
    "4ee8f2cc-6c76-404d-a7d0-03d7e508f6be",
    "8fc988d3-484c-40c6-900f-e1cfcc62e35e",
    "d07b5e39-f445-4f09-aed5-fa6c93821871",
    "f205d673-56a2-4e48-b2d2-e9cbe229ec73",
    "e7665821-b33f-45ba-9ec1-e6f57b86330a",
    "16d47c3d-42d6-4728-93b6-339917db8ece",
    "b45b1ccd-1c87-4833-955b-060ee8396486",
    "01483494-075f-4df7-895d-e971eb1aff70",
    "497b67a6-c237-4094-ad91-4d9e041d45d5",
    "a418d44c-ea94-415e-a3ae-4c8a746dfb3b",
    "41bd638d-de74-407b-9933-e1143060ca5e",
    "de4c2191-4075-4157-a70b-7bc4de605898",
    "8a36be22-a47b-48d4-9c85-388a3b89cfd3",
    "ba6d5e3f-c3fd-411b-8471-ed1ee5a2023d",
    "b904cae6-70fa-415e-88c9-3b743ac3afc3",
    "a38c2480-51ca-4254-97c8-ec340041e0f4",
    "eb4b4184-d36d-4f17-a501-c118b2727d5f",
    "eaeb6081-5839-402e-8db2-57873f42ab44",
    "17ce3511-0cd3-444a-9c6a-69b144807f80",
    "03f8fe53-a33f-412e-8e2d-0121e5604c41"
]

rackCompleteSensorList = []

#grabs the full list of sensor data for each rack
for rackId in racksWithSensors:
    response = requests.get(f'https://msu.hyperviewhq.com/api/asset/sensors/{rackId}', headers=headers)
    response = response.json()
    rackCompleteSensorList.append(response)

#print(len(rackCompleteSensorList))

#Ids for what the temp and humidity sensors look like inside of the sensor data from above
humidityTypeId = '466799ea-0e25-e211-8183-001c42e521d8'
temperatureTypeId = '52835710-56f9-4311-babb-67b21b423c7d'

sensorIdsListDict = []

#iterates through each racks full sensor list to grab each racks unique temp/humidity sensor ids
for rack in rackCompleteSensorList:

    #Empty Dict to add each rack to!
    dict = {}
    nameAppended = False

    #iterates over each rack
    for sensorList in rack:
        if sensorList['sensorTypeId'] == humidityTypeId:

            if not nameAppended:
                dict['PDUName'] = sensorList['sourceAssetDisplayName']
                nameAppended = True

            humidityUniqueId = sensorList['id']
            dict['humidityId'] = humidityUniqueId
        if sensorList['sensorTypeId'] == temperatureTypeId:

            if not nameAppended:
                dict['PDUName'] = sensorList['sourceAssetDisplayName']
                nameAppended = True

            temperatureUniqueId = sensorList['id']
            dict['temperatureId'] = temperatureUniqueId

        if len(dict) == 3:
            sensorIdsListDict.append(dict)
            break

#print(len(sensorIdsListDict))

#Part 3 use unique sensor Ids for each rack to grab summary data
headers = {
    'accept': 'application/json',
    'Authorization': f'Bearer {token}'
}

summarySensorData = []

for rack in sensorIdsListDict:
    rackDict = {}
    rackName = extract_rack_name(rack['PDUName'])
    rackDict["Rack"] = rackName

    #grabs humidity summary data
    response = requests.get(f'https://msu.hyperviewhq.com/api/asset/sensorsDailySummaries/numeric/last7Days?sensorIds={rack["humidityId"]}', headers=headers)
    response = response.json()

    humidityDataPointList = []
    for x in response:
        for data_point in x['sensorDataPoints']:
            humidityDict = {}
            humidityDict['date'] = extract_date(data_point['r'])
            humidityDict['average'] = data_point['avg']
            humidityDict['maximum'] = data_point['max']
            humidityDict['minimum'] = data_point['min']
            humidityDict['last'] = data_point['lst']
            humidityDataPointList.append(humidityDict)
    rackDict['humiditySummaryData'] = humidityDataPointList
            
    #grabs temperature summary data
    response = requests.get(f'https://msu.hyperviewhq.com/api/asset/sensorsDailySummaries/numeric/last7Days?sensorIds={rack["temperatureId"]}', headers=headers)
    response = response.json()

    temperatureDataPointList = []
    for x in response:
        for data_point in x['sensorDataPoints']:
            temperatureDict = {}
            temperatureDict['date'] = extract_date(data_point['r'])
            temperatureDict['average'] = data_point['avg']
            temperatureDict['maximum'] = data_point['max']
            temperatureDict['minimum'] = data_point['min']
            temperatureDict['last'] = data_point['lst']
            temperatureDataPointList.append(temperatureDict)
    rackDict['temperatureSummaryData'] = temperatureDataPointList

    summarySensorData.append(rackDict)

# Flatten the data
temperature_data = []
humidity_data = []

for item in summarySensorData:
    rack_name = item['Rack']
    for entry in item['temperatureSummaryData']:
        temp_entry = {
            'Rack': rack_name,
            'Date': entry['date'],
            'Average': entry['average'],
            'Maximum': entry['maximum'],
            'Minimum': entry['minimum'],
            'Last': entry['last']
        }
        temperature_data.append(temp_entry)
    for entry in item['humiditySummaryData']:
        hum_entry = {
            'Rack': rack_name,
            'Date': entry['date'],
            'Average': entry['average'],
            'Maximum': entry['maximum'],
            'Minimum': entry['minimum'],
            'Last': entry['last']
        }
        humidity_data.append(hum_entry)

# Convert to DataFrame
temp_df = pd.DataFrame(temperature_data, columns=['Rack', 'Date', 'Average', 'Maximum', 'Minimum', 'Last'])
humidity_df = pd.DataFrame(humidity_data, columns=['Rack', 'Date', 'Average', 'Maximum', 'Minimum', 'Last'])

# Write to Excel
with pd.ExcelWriter('rack_data.xlsx') as writer:
    temp_df.to_excel(writer, sheet_name='TemperatureSummary', index=False)
    humidity_df.to_excel(writer, sheet_name='HumiditySummary', index=False)
