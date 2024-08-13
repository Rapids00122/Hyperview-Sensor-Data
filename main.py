import pandas as pd
import requests
from datetime import datetime
import win32com.client as win32
import os

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
    """
    Extracts the rack name from a PDU name.

    Args:
        PDUName (str): The name of the PDU, formatted with hyphens.

    Returns:
        str: The extracted rack name.
    """
    return PDUName.split('-')[1]

def extract_date(timestamp):
    """
    Extracts and formats the date from a timestamp string.

    Args:
        timestamp (str): The timestamp in ISO format.

    Returns:
        str: The formatted date as 'YYYY-MM-DD'.
    """
    return datetime.fromisoformat(timestamp).date().strftime('%Y-%m-%d')

def get_access_token():
    """
    Fetches an access token from the Hyperview API.

    Returns:
        str: The access token for authorization in API requests.
    """
    response = requests.post(TOKEN_URL, data={
        'client_id': '8dcaf879-9077-42b1-a7b3-05dbce43c5b0',
        'client_secret': '08c1b6ec-9800-42d6-ada1-be2244e76e38',
        'grant_type': 'client_credentials'
    }, headers={'Content-Type': 'application/x-www-form-urlencoded'})
    response.raise_for_status()
    return response.json()['access_token']

def fetch_sensor_data(rack_id):
    """
    Fetches sensor data for a specific rack from the Hyperview API.

    Args:
        rack_id (str): The unique identifier of the rack.

    Returns:
        dict: The response JSON containing sensor data.
    """
    response = requests.get(f'{SENSOR_API_URL}/{rack_id}', headers=HEADERS)
    response.raise_for_status()
    return response.json()

def process_sensor_data(rack_data):
    """
    Processes sensor data for a given rack to extract unique humidity and temperature sensor IDs.

    Args:
        rack_data (list of dict): A list of sensor data dictionaries, where each dictionary contains details of a sensor, including 'sensorTypeId' and 'sourceAssetDisplayName'.

    Returns:
        dict: A dictionary with the following keys:
            - 'humidityId' (str): The unique identifier for the humidity sensor, if present.
            - 'temperatureId' (str): The unique identifier for the temperature sensor, if present.
            - 'PDUName' (str): The display name of the PDU, extracted from the sensor data if not already present.

    Notes:
        The function assumes that the sensor data includes 'sensorTypeId' values matching predefined constants for humidity and temperature sensors. 
        It stops processing once both humidity and temperature sensor IDs are found.
    """
    result = {}
    for sensor in rack_data:
        if sensor['sensorTypeId'] == HUMIDITY_TYPE_ID:
            result['humidityId'] = sensor['id']
            if 'PDUName' not in result:
                result['PDUName'] = sensor['sourceAssetDisplayName']
        elif sensor['sensorTypeId'] == TEMPERATURE_TYPE_ID:
            result['temperatureId'] = sensor['id']
            if 'PDUName' not in result:
                result['PDUName'] = sensor['sourceAssetDisplayName']
        if 'humidityId' in result and 'temperatureId' in result:
            break
    return result

def fetch_summary_data(sensor_id):
    """
    Fetches summary data for a specific sensor from the Hyperview API.

    Args:
        sensor_id (str): The unique identifier of the sensor.

    Returns:
        list: A list of dictionaries containing the summary data points for the sensor.
    """
    response = requests.get(f'{SUMMARY_API_URL}?sensorIds={sensor_id}', headers=HEADERS)
    response.raise_for_status()
    return response.json()

def send_email(file_path):
    # Initialize Outlook
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)

    current_date = datetime.now().date()

    # Email details
    mail.To = 'lakos@msu.edu'
    mail.Subject = f'Weekly Data Center Sensor Report - {current_date}'
    mail.Body = 'Please find the attached weekly report.'

    # Attach the file
    mail.Attachments.Add(file_path)

    # Send the email
    mail.Send()

def main():
    """
        Main function to execute the data extraction and processing workflow.

        This function performs the following tasks:
        1. Fetches the access token for authorization.
        2. Retrieves sensor data for each rack.
        3. Collects summary data for temperature and humidity sensors.
        4. Flattens the collected data and saves it to an Excel file.
    """

    # Step 1: Generate Hyperview Access Token
    token = get_access_token()
    HEADERS['Authorization'] = f'Bearer {token}'

    # List of all unique rack IDs with sensors
    racks_with_sensors = [
        "127e5e7e-3583-4a1c-9085-298f0227a505",
        "688f6ac3-dc01-44e2-a3aa-448acd6d92ca",
        "80d75e0c-558b-4c5e-afb2-a34501441b04",
        "2a5c3f17-2fc2-4c96-9715-8e1dc8d6fae9",
        "8ddcebc2-0212-4a1e-91e8-9c6552cbd592",
        "02e9aa2c-c2ce-4921-887f-0591563e66f8",
        "80185bcb-f9ec-498b-8824-bfe8066301d0",
        "cbf43771-8874-46ff-9c7b-419f4bbfd9fc",
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

    # Step 2: Fetch sensor data for each rack
    rack_complete_sensor_list = [fetch_sensor_data(rack_id) for rack_id in racks_with_sensors]

    # Process each rack's sensor data
    sensor_ids_list_dict = [process_sensor_data(rack) for rack in rack_complete_sensor_list]

    # Step 3: Use unique sensor IDs to fetch summary data
    summary_sensor_data = []
    for rack in sensor_ids_list_dict:
        rack_dict = {
            "Rack": extract_rack_name(rack['PDUName'])
        }

        humidity_data = fetch_summary_data(rack['humidityId'])
        rack_dict['humiditySummaryData'] = [
            {
                'date': extract_date(dp['r']),
                'average': dp['avg'],
                'maximum': dp['max'],
                'minimum': dp['min'],
                'last': dp['lst']
            }
            for data in humidity_data for dp in data['sensorDataPoints']
        ]

        temperature_data = fetch_summary_data(rack['temperatureId'])
        rack_dict['temperatureSummaryData'] = [
            {
                'date': extract_date(dp['r']),
                'average': dp['avg'],
                'maximum': dp['max'],
                'minimum': dp['min'],
                'last': dp['lst']
            }
            for data in temperature_data for dp in data['sensorDataPoints']
        ]

        summary_sensor_data.append(rack_dict)

    # Flatten the data
    temperature_data = [
        {
            'Rack': item['Rack'],
            'Date': entry['date'],
            'Average': entry['average'],
            'Maximum': entry['maximum'],
            'Minimum': entry['minimum'],
            'Last': entry['last']
        }
        for item in summary_sensor_data for entry in item['temperatureSummaryData']
    ]

    humidity_data = [
        {
            'Rack': item['Rack'],
            'Date': entry['date'],
            'Average': entry['average'],
            'Maximum': entry['maximum'],
            'Minimum': entry['minimum'],
            'Last': entry['last']
        }
        for item in summary_sensor_data for entry in item['humiditySummaryData']
    ]

    # Convert to DataFrame and write to Excel
    temp_df = pd.DataFrame(temperature_data, columns=['Rack', 'Date', 'Average', 'Maximum', 'Minimum', 'Last'])
    humidity_df = pd.DataFrame(humidity_data, columns=['Rack', 'Date', 'Average', 'Maximum', 'Minimum', 'Last'])

    # Convert 'Date' to datetime
    temp_df['Date'] = pd.to_datetime(temp_df['Date'])
    humidity_df['Date'] = pd.to_datetime(humidity_df['Date'])

    # Adds in a 'Row' column that takes the last two str digits of the 'Rack' colunmn to get the row
    temp_df['Row'] = temp_df['Rack'].str[-2:]
    humidity_df['Row'] = humidity_df['Rack'].str[-2:]

    # Calculate weekly averages
    # Creates duplicate weekly temp and humidity data frame
    temp_weekly = temp_df
    humidity_weekly = humidity_df

    # Drops the 'Row' column from the original temp_df and humidity_df as they are no longer needed
    temp_df = temp_df.drop(columns=['Row'])
    humidity_df = humidity_df.drop(columns=['Row'])

    # Drops the 'Rack' and 'Date' columns from the temp_weekly and Humidity_weekly as these data frames will
    # only be recording the averages of a whole row and will only include a start and end date
    temp_weekly = temp_weekly.drop(columns=['Rack','Date', 'Last'])
    humidity_weekly = humidity_weekly.drop(columns=['Rack','Date', 'Last'])

    # Groups the temp_weekly and humidity_weekly data frames by their row and calcualtes the mean of all the
    # data in the data frame and then resets the index back to the original order
    temp_weekly = temp_weekly.groupby(['Row']).mean().reset_index()
    humidity_weekly = humidity_weekly.groupby(['Row']).mean().reset_index()

    # Grabs the min and max values of the original data frames and applies those as start and end dates for the new data frames
    temp_weekly['Start Date'] = temp_df['Date'].min()
    humidity_weekly['Start Date'] = humidity_df['Date'].min()
    temp_weekly['End Date'] = temp_df['Date'].max()
    humidity_weekly['End Date'] = humidity_df['Date'].max()
    
    # Rearrange weekly average columns
    temp_weekly = temp_weekly[['Row', 'Start Date', 'End Date', 'Average', 'Maximum', 'Minimum']]
    humidity_weekly = humidity_weekly[['Row', 'Start Date', 'End Date', 'Average', 'Maximum', 'Minimum']]

    # Rounds all the data to .2 decimal points
    temp_df[['Average', 'Maximum', 'Minimum', 'Last']] = temp_df[['Average', 'Maximum', 'Minimum', 'Last']].round(2)
    humidity_df[['Average', 'Maximum', 'Minimum', 'Last']] = humidity_df[['Average', 'Maximum', 'Minimum', 'Last']].round(2)
    temp_weekly[['Average', 'Maximum', 'Minimum']] = temp_weekly[['Average', 'Maximum', 'Minimum']].round(2)
    humidity_weekly[['Average', 'Maximum', 'Minimum']] = humidity_weekly[['Average', 'Maximum', 'Minimum']].round(2)

    # Converts data frames from pd.datetime back to strtime
    temp_df['Date'] = temp_df['Date'].dt.strftime('%Y-%m-%d')
    humidity_df['Date'] = humidity_df['Date'].dt.strftime('%Y-%m-%d')
    temp_weekly['Start Date'] = temp_weekly['Start Date'].dt.strftime('%Y-%m-%d')
    temp_weekly['End Date'] = temp_weekly['End Date'].dt.strftime('%Y-%m-%d')
    humidity_weekly['Start Date'] = humidity_weekly['Start Date'].dt.strftime('%Y-%m-%d')
    humidity_weekly['End Date'] = humidity_weekly['End Date'].dt.strftime('%Y-%m-%d')

    # writes to Excel
    with pd.ExcelWriter('Rack_Sensor_Data.xlsx') as writer:
        temp_df.to_excel(writer, sheet_name='Temperature Data', index=False)
        temp_weekly.to_excel(writer, sheet_name='Temperature Row Weekly Averages', index=False)
        humidity_df.to_excel(writer, sheet_name='Humidity Data', index=False)
        humidity_weekly.to_excel(writer, sheet_name='Humidity Row Weekly Averages', index=False)

    excel_file = os.path.abspath("Rack_Sensor_Data.xlsx")

    send_email(excel_file)

if __name__ == "__main__":
    main()
