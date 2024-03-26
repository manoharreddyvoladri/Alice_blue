import time
import openpyxl
import pandas as pd
import datetime
from pya3 import Aliceblue

# Replace with your Aliceblue API credentials
username = '1243674'
api_key = 'AbLLDhIKc064Zt2d5sSdGoJPoz3fDybL247ppJHx42tWli28eKMLdoHtCSNmIhkGye1Ub3r8QhBh1xFLtDtK6R2ixO2j5k38Vki7ijzxjNLeyhAufedQkuQRgojLZKqD'

# Set up the Aliceblue API client
alice = Aliceblue(username, api_key)

def get_and_print_data():
    # Get the session ID (only needed once)
    session_id = alice.get_session_id()['sessionID']

    # Parameters for historical data request
    days = 1  # We only need today's data
    exchange = 'MCX'
    spot_symbol = 'CRUDEOILM'
    interval = '1'
    indices = False
    from_date = datetime.datetime.now() - datetime.timedelta(days=days)
    to_date = datetime.datetime.now()

    # Get the token for the specified symbol
    token = alice.get_instrument_by_symbol(exchange, spot_symbol)

    # Get the historical data
    data = alice.get_historical(token, from_date, to_date, interval, indices)

    # Convert the data to a pandas DataFrame
    data = pd.DataFrame(data)

    # Convert the 'datetime' column to a datetime object
    data['datetime'] = pd.to_datetime(data['datetime'])

    # Calculate the 'time' column
    data['time'] = data['datetime'].apply(lambda x: x.time())

    # Calculate Profit Metrics (Optional)
    data['Differece'] = data['close'] - data['open']
    data['profit_percentage'] = ((data['close'] / data['open']) - 1) * 100

    # Access today's data (last row)
    todays_data = data.iloc[-1]

    # Print current time and today's data
    current_time = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print(f"{current_time}         {todays_data['time']}               {todays_data['open']:.2f}      {todays_data['close']:.2f}      {todays_data['Differece']:.2f}             {todays_data['profit_percentage']:.2f}%")

    # Create a new Excel workbook and worksheet if they don't exist
    if not hasattr(get_and_print_data, 'workbook') or not hasattr(get_and_print_data, 'worksheet'):
        get_and_print_data.workbook = openpyxl.Workbook()
        get_and_print_data.worksheet = get_and_print_data.workbook.active

    # Write the column names to the first row of the worksheet only once
    if not hasattr(get_and_print_data, 'column_names_written'):
        get_and_print_data.column_names_written = True
        column_names = ["Date Time", "market time", "Open", "Close", "Difference", "Profit(%)"]
        get_and_print_data.worksheet.append(column_names)

    # Append the data to the worksheet
    new_row = get_and_print_data.worksheet.max_row + 1
    get_and_print_data.worksheet.append([datetime.datetime.now(), todays_data['time'], todays_data['open'], todays_data['close'], todays_data['Differece'], todays_data['profit_percentage']])

   # Save the workbook
    get_and_print_data.workbook.save('live_data.xlsx')

print(f"  Date      Time           market time              Open          Close         Difference      Profit(%)    ")
while True:
    get_and_print_data()
    time.sleep(25)