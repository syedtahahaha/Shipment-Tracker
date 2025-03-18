import http.client
import urllib.parse
import json
import requests
import csv
from dotenv import load_dotenv
import os
from datetime import datetime
import time

# Load environment variables from .env file
load_dotenv()

# Get FedEx API credentials
FEDEX_CLIENT_ID = os.getenv('FEDEX_CLIENT_ID')
FEDEX_CLIENT_SECRET = os.getenv('FEDEX_CLIENT_SECRET')

# Get DHL API credentials
DHL_API_KEY = os.getenv('DHL_API_KEY')


# Function to get FedEx authorization (OAuth token)
def getFedExBearerAuthorization():
    url = "https://apis.fedex.com/oauth/token"

    payload = {
        "grant_type": "client_credentials",
        "client_id": FEDEX_CLIENT_ID,
        "client_secret": FEDEX_CLIENT_SECRET
    }
    headers = {
        'Content-Type': "application/x-www-form-urlencoded"
    }

    response = requests.post(url, data=payload, headers=headers)
    print("FedEx OAuth response:", response.status_code)

    if response.status_code == 200:
        authorization = response.json()['access_token']
        return authorization
    else:
        print("Error obtaining FedEx token.")
        return None


# Function to get DHL tracking result
def getDHLTrackingResult(trackingNumber):
    # Encode parameters for the URL
    params = urllib.parse.urlencode({
        'trackingNumber': trackingNumber,
        'service': 'express'  # This can be adjusted based on the specific service you need
    })

    # Set up headers with DHL API key
    headers = {
        'Accept': 'application/json',
        'DHL-API-Key': DHL_API_KEY  # API Key loaded from environment
    }

    # Set up connection to DHL API
    connection = http.client.HTTPSConnection("api-eu.dhl.com")

    # Send GET request to DHL API
    connection.request("GET", "/track/shipments?" + params, "", headers)

    # Get the response from the server
    response = connection.getresponse()

    # Parse response data
    status = response.status
    reason = response.reason
    data = json.loads(response.read())

    print(f"Processing DHL tracking for: {trackingNumber}")
    # Print the response status and reason for debugging
    # print("Status: {} and reason: {}".format(status, reason))

    # Check if status code is 200 (OK) and process the response data
    if status == 200:
        #print("Tracking Info:")

        # Extracting shipment information
        shipment = data['shipments'][0]
       # print(shipment)

        # Extracting relevant fields similar to FedEx response
        dhl_tracking_number = shipment['id']
        delivery_status = shipment['status']['description']  # Status description
        status_remarks = shipment['status'].get('remark', 'N/A')  # Get remark (if available)
        lasteventdate = shipment['status']['timestamp']
        lasteventlocation = shipment['status']['location']['address']['addressLocality']

        # Return the information in the same format as FedEx tracking result
        return [dhl_tracking_number, delivery_status, status_remarks, lasteventdate, lasteventlocation]
    else:
        # If the response status is not 200, print an error message
        print(f"DHL tracking failed for {trackingNumber}: {status}")
        return None

    # Close the connection
    connection.close()


# Function to process FedEx tracking
def processFedExTracking(trackingNumber, token):
    url = "https://apis.fedex.com/track/v1/trackingnumbers"
    headers = {
        'Content-Type': "application/json",
        'Authorization': "Bearer " + token
    }

    payload = '{ "trackingInfo": [ { "trackingNumberInfo": { "trackingNumber": "' + trackingNumber.strip() + '" } } ], "includeDetailedScans": true }'

    response = requests.post(url, data=payload, headers=headers)

    print(f"Processing FedEx tracking for: {trackingNumber}")

    if response.status_code == 200:
        tracking_data = response.json()
        #print(tracking_data) #for debugging

        try:
            trackingNumber = tracking_data['output']['completeTrackResults'][0]['trackingNumber']
            deliveryStatus = tracking_data['output']['completeTrackResults'][0]['trackResults'][-1]['latestStatusDetail']['statusByLocale']
            try:
                reasons = tracking_data['output']['completeTrackResults'][0]['trackResults'][-1]['latestStatusDetail']['ancillaryDetails']
            except KeyError:
                reasons = ""

            deliveryStatusDescription = tracking_data['output']['completeTrackResults'][0]['trackResults'][-1]['latestStatusDetail']['description']
            scanEvent = tracking_data['output']['completeTrackResults'][0]['trackResults'][-1]['scanEvents'][0]
            latestStatusDate = scanEvent['date']

            try:
                latestStatusLocation = scanEvent['scanLocation']['city']
            except KeyError:
                latestStatusLocation = scanEvent['scanLocation'].get('Country', 'Unknown')  # Fallback to 'Country' or 'Unknown' if neither are available

  
            return [trackingNumber, deliveryStatus, deliveryStatusDescription, latestStatusDate, latestStatusLocation, reasons]
        except KeyError as e:
            print(f"KeyError occurred: {e} - FedEx tracking failed for: {trackingNumber}.")
            return None
        
    else:
        print(f"FedEx tracking failed for {trackingNumber}: {response.status_code}")
        return None


# Variables and Parameters
token = getFedExBearerAuthorization()
with open('1. Tracking Numbers.txt', "r") as text, open('3. Tracking Results.csv', "w", newline='') as output:
    writer = csv.writer(output, lineterminator='\n')
    writer.writerow(['Tracking Number', 'Updates'])

    # Complete API request
    while True:
        trackingNumber = text.readline()

        if not trackingNumber:
            break
        
        trackingNumber = trackingNumber.strip()

        # Check if the tracking number is FedEx (12 digits) or DHL (10 digits)
        if len(trackingNumber) == 12:  # FedEx tracking number
            if token:  # Check if FedEx token is available
                result = processFedExTracking(trackingNumber, token)
                if result:
                    try:
                        # Extracting relevant details for DHL from the response
                        fedex_tracking_number = result[0]
                        delivery_status = result[1] + " - " + result[2] + ' at ' + result[4]  # Status description + remarks
                        lasteventdate = datetime.fromisoformat(result[3]).strftime("%d/%m/%Y %I:%M%p")  # date
                        actionDescription = [entry['actionDescription'] for entry in result[5]]
                        reasonDescription = [entry['reasonDescription'] for entry in result[5]]
                        
                        # Combine reasonDescription and actionDescription
                        combined_descriptions = []

                        for reason, action in zip(reasonDescription, actionDescription):
                            if reason and action:
                                combined_descriptions.append(f"{reason} - {action}")
                            elif reason:
                                combined_descriptions.append(reason)
                            elif action:
                                combined_descriptions.append(action)

                        # Join the descriptions with ' | ' separator for the final output
                        combined_descriptions_string = ' | '.join(combined_descriptions)

                        # Writing the extracted details into CSV
                        if combined_descriptions_string == '':
                            writer.writerow([fedex_tracking_number, lasteventdate + ' - ' + delivery_status + ' | '.join(combined_descriptions)])
                        else:
                            writer.writerow([fedex_tracking_number, lasteventdate + ' - ' + delivery_status + ' - Details: ' + ' | '.join(combined_descriptions)])

                        # Explicitly flush the output after writing to ensure data is written immediately
                        output.flush()

                    except KeyError as e:
                        print(f"Error extracting data for tracking number {trackingNumber}: {e}")
                        writer.writerow([trackingNumber, 'Error', 'Error', 'Error'])
                        output.flush()
            else:
                print("FedEx token is missing or expired.")
                writer.writerow([trackingNumber, 'Error', 'Error', 'Error'])
                output.flush()

        elif len(trackingNumber) == 10:  # DHL tracking number
            result = getDHLTrackingResult(trackingNumber)
            if result:
                try:
                    # Extracting relevant details for DHL from the response
                    dhl_tracking_number = result[0]
                    delivery_status = result[1] + ' at ' + result[4] + " - " + result[2]  # Status description + remarks
                    lasteventdate = datetime.fromisoformat(result[3]).strftime("%d/%m/%Y %I:%M%p")  # date

                    # Writing the extracted details into CSV
                    writer.writerow([dhl_tracking_number, lasteventdate + ' - ' + delivery_status])
                    output.flush()  # Ensure data is immediately written to the file
                except KeyError as e:
                    print(f"Error extracting data for tracking number {trackingNumber}: {e}")
                    writer.writerow([trackingNumber, 'Error', 'Error', 'Error'])
                    output.flush()

            # Add a 5-second wait between DHL requests to respect the rate limit
            time.sleep(5.1)
        else:
            print(f"Invalid tracking number: {trackingNumber}")
            writer.writerow([trackingNumber, 'Error', 'Error', 'Error'])
            output.flush()

# Add a message after everything is done
print("\nAll tracking numbers have been processed.")
print("\nClose this window to proceed.")
