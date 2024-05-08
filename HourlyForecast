import requests
import pandas as pd
from datetime import datetime


#get current date and time
current_date = datetime.now()
print("Current Date and Time:", current_date)

#set output path
excel_file = "hourly_forecast.xlsx"

#zipcodes Latitude and Longitude
zip_codes = {
    "Enter ZipCode": (xx.xxxx, -xxx.xxxx)
}

def get_hourly_forecast(latitude, longitude):
    #get forecast grid endpoint
    points_url = f"https://api.weather.gov/points/{latitude},{longitude}"
    response = requests.get(points_url)
    data = response.json()

    #Get forecast hourly endpoint
    hourly_forecast_url = data['properties']['forecastHourly']

    #retrieve hourly forecast data
    response = requests.get(hourly_forecast_url)
    hourly_forecast_data = response.json()

    #extract temp and relative humifity for each hour
    hourly_forecast = []
    for forecast in hourly_forecast_data['properties']['periods']:
        print(forecast)

        time = str(datetime.fromisoformat(forecast['startTime'][:-6]).strftime('%m/%d/%Y %H:%M')) + " " + "PDT"
        temperature = forecast['temperature']
        temperature_unit = forecast['temperatureUnit']
        relative_humidity = forecast['relativeHumidity']['value']
        short_forecast = forecast['shortForecast']
        wind_speed = forecast['windSpeed']
        wind_direction = forecast['windDirection']
        probability_Of_Precipitation = str(forecast['probabilityOfPrecipitation']['value'])
        probability_Of_Precipitation_unit = forecast['probabilityOfPrecipitation']['unitCode'].split(":")[1]

        dewpoint = forecast['dewpoint']['value']
        dewpoint_unit = forecast['dewpoint']['unitCode'].split(":")[1]
        daytime = forecast['isDaytime']


        if datetime.fromisoformat(forecast['startTime'][:-6]).strftime('%Y, %m, %d') > str(current_date.strftime('%Y, %m, %d')):
            hourly_forecast.append({
                'Effective Date': time,
                'DewPoint': dewpoint,
                'Dewpoint Unit': dewpoint_unit,
                'Probability Of Precipitation': probability_Of_Precipitation,
                'Probability Of Precipitation Unit': probability_Of_Precipitation_unit,
                'Temp': temperature,
                'Temp Unit': temperature_unit,
                'RELH': relative_humidity,
                'Short Forecast': str(short_forecast),
                'Wind Speed': wind_speed,
                'Wind Direction': wind_direction,
                'Day Time': daytime,
            })
    return hourly_forecast

with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
    for zip_code, coordinates in zip_codes.items():
        #get hourly forecast data
        hourly_forecast_data = get_hourly_forecast(coordinates[0],coordinates[1])
        weather_dataframe = pd.DataFrame(hourly_forecast_data)
        weather_dataframe.to_excel(writer,index=False,sheet_name=zip_code)
        pd.DataFrame().to_excel(writer,sheet_name=zip_code, index=False)

print(f"Hourly forecast data has been written to {excel_file}")
