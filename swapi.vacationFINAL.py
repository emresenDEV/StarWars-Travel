#-----------------IMPORTS-----------------

import requests, json
from openpyxl import Workbook

#-----------------API-----------------

planetAPI = 'https://swapi.dev/api/planets/' #60 Planets
starshipAPI = 'https://swapi.dev/api/starships/' #36 Starships
vehicleAPI = 'https://swapi.dev/api/vehicles/' #39 Vehicles

#-----------------OPENPYXL WORKBOOK SETUP-----------------+6

wb = Workbook()
planet_ws = wb.active
planet_ws.title = "Star Wars Vacation Spots"

starship_ws = wb.create_sheet("Getting There | Starships")

vehicle_ws = wb.create_sheet("Getting Around | Vehicles")

#-----------------HEADER SETUP-----------------

#Planet Worksheet
headers = ['Planet', 'Climate', 'Terrain', 'Population', 'Film Count']
for index, header in enumerate(headers, 1):
    planet_ws.cell(row=1, column=index, value=header)

#Starship Worksheet
headers2 = ['Starship', 'Type of Ship', 'Price', 'Maximum Speed', 'Total Seats', 'Crew Size', 'Ship Size', 'Passenger Experience']
for index, header in enumerate(headers2, 1):
    starship_ws.cell(row=1, column=index, value=header)

#Vehicle Worksheet
headers3 = ['Vehicle', 'Price', 'Speed Limit', 'Maximum Group Size', 'Crew Size', 'Vehicle Size', 'Passenger Experience']
for index, header in enumerate(headers3, 1):
    vehicle_ws.cell(row=1, column=index, value=header)

#-----------------API CONNECTION & DATA PULLING-----------------

#Planet Worksheet
planets = []

response = requests.get(planetAPI)
planetData = response.json()
planets += planetData['results']

while planetData['next']:
    response = requests.get(planetData['next'])
    planetData = response.json()
    planets += planetData['results']
    
#Starship Worksheet
starships = []

response2 = requests.get(starshipAPI)
starshipData = response2.json()
starships += starshipData['results']

while starshipData['next']:
    response2 = requests.get(starshipData['next'])
    starshipData = response2.json()
    starships += starshipData['results']
    
#Added to remove starships without passenger capacity                           #Mionne's code
starships[:] = (ship for ship in starships if ship['passengers'] != "0")        #Mionne's code
starships[:] = (ship for ship in starships if ship['passengers'] != "n/a")      #Mionne's code
starships[:] = (ship for ship in starships if ship['passengers'] != "unknown")  #Mionne's code

#Vehicle Worksheet
vehicles = []

response3 = requests.get(vehicleAPI)
vehicleData = response3.json()
vehicles += vehicleData['results']

while vehicleData['next']:
    response3 = requests.get(vehicleData['next'])
    vehicleData = response3.json()
    vehicles += vehicleData['results']
    
#Added to remove vehicles without passenger capacity                        #Mionne's code
vehicles[:] = (veh for veh in vehicles if veh['passengers'] != "0")         #Mionne's code
vehicles[:] = (veh for veh in vehicles if veh['passengers'] != "n/a")       #Mionne's code
vehicles[:] = (veh for veh in vehicles if veh['passengers'] != "unknown")   #Mionne's code

#-----------------POPULATING DATA IN 'STAR WARS VACATION SPOTS' WORKSHEET-----------------

#Planet Worksheet
for row_index, planet in enumerate(planets, start=2):
    for col_index, header in enumerate(headers, start=1):
        if header == "Planet":
            value = planet.get("name")
        elif header == "Climate":
            value = planet.get("climate")
        elif header == "Terrain":
            value = planet.get("terrain")
        elif header == "Population":
            value = planet.get("population")
            if value == "unknown" or value == "":
                value = "Unknown"
        elif header == "Film Count":
            value = len(planet.get("films"))
        planet_ws.cell(row=row_index, column=col_index, value=value)

#Starship Worksheet
for row_index, starship in enumerate(starships, start=2):
    for col_index, header in enumerate(headers2, start=1):
        if header == "Starship":
            value = starship.get("name")
        elif header == "Type of Ship":
            value = starship.get("starship_class")
        elif header == "Price":
            value = starship.get("cost_in_credits")
        elif header == "Maximum Speed":
            value = starship.get("max_atmosphering_speed")
            if value == "Unknown" or value == [] or value == "" or value == "n/a":
                value = "This will be a slow ride. The starship ride IS your vacation."
        elif header == "Total Seats":
            value = starship.get("passengers")
        elif header == "Crew Size":
            value = starship.get("crew")
        elif header == "Ship Size":
            value = starship.get("length", "Unknown Size")
            try:
                value_float = float(value) #TRYing to convert value to a float (float because we may have decimals somewhere in the data)
                if 1 <= value_float <= 20:
                    value = "Small"
                elif 20.1 <= value_float <= 500:
                    value = "Medium"
                elif 500.1 <= value_float <= 4000:
                    value = "Large"
                elif 4000.1 <= value_float:
                    value = "Massive"
            except ValueError: #if value is not a number and pops an error, it will be a string that prints:
                value = "Unknown Size" #We can change this to anything we want it to say in cell on the worksheet.
        elif header == "Passenger Experience":          #NOTE: I used the cruise ship crew to passenger ratio as a guide: passengers / crew = how many passengers each crew member is responsible for. 1:1 is top notch. 
            passengers = starship.get("passengers", "0") #Defaults to 0 if no passengers
            crew = starship.get("crew", "1") #Defaults to 1 if no crew (this is IMPORTANT: you cannot divide by 0)
            if passengers.isnumeric() and crew.isnumeric(): #Checks if passengers and crew are numbers, does not convert anything
                passengers = int(passengers) #Converts passengers to an int
                crew = int(crew) #Converts crew to an int
                if crew > 0:
                    ratio = passengers / crew   #cruise ship quality of care ratio formula
                    if ratio <= 1.0:
                        value = "Luxury"
                    elif ratio >= 1.1 and ratio <= 2.9:
                        value = "Comfort"
                    elif ratio >= 3:
                        value = "Economy"
                else:
                    value = "Risky Experience"
            else:
                value = "Mystery Experience"    
        starship_ws.cell(row=row_index, column=col_index, value=value)

#Vehicle Worksheet
for row_index, vehicle in enumerate(vehicles, start=2):
    for col_index, header in enumerate(headers3, start=1):
        if header == "Vehicle":
            value = vehicle.get("model")
        elif header == "Price":
            value = vehicle.get("cost_in_credits")
        elif header == "Speed Limit":
            value = vehicle.get("max_atmosphering_speed")
        elif header == "Maximum Group Size":
            value = vehicle.get("passengers")
        elif header == "Crew Size":
            value = vehicle.get("crew")
        elif header == "Vehicle Size":
            value = vehicle.get("length", "Unknown Size")
            try:
                value_float = float(value) #TRYing to convert value to a float
                if 1 <= value_float <= 20:
                    value = "Small"
                elif 20.1 <= value_float <= 500:
                    value = "Medium"
                elif 500.1 <= value_float <= 4000:
                    value = "Large"
                elif 4000.1 <= value_float:
                    value = "Massive"
            except ValueError: #if value is not a number, it will be a string
                value = "Unknown Size" #we could also assign a default value here if the value cannot be converted to a float.
        elif header == "Passenger Experience":          #NOTE: I used cruise ship crew to passenger ratio as a guide: passengers / crew = how many passengers each crew member is responsible for. 1:1 is top notch.
            passengers = vehicle.get("passengers", "0") #Defaults to 0 if no passengers
            crew = vehicle.get("crew", "1") #Defaults to 1 if no crew data (this is important as you cannot divide by 0)
            if passengers.isnumeric() and crew.isnumeric(): #Checks if passengers and crew are numbers, does not convert anything
                passengers = int(passengers) #Converts passengers to an int
                crew = int(crew) #Converts crew to an int
                if crew > 0:
                    ratio = passengers / crew #cruise ship quality of care ratio formula
                    if ratio <= 1.0:
                        value = "Luxury"
                    elif ratio >= 1.1 and ratio <= 2.9:
                        value = "Comfort"
                    elif ratio >= 3:
                        value = "Economy"
                else:
                    value = "Risky Experience"
            else:
                value = "Mystery Experience"   

        vehicle_ws.cell(row=row_index, column=col_index, value=value)
                
#-----------------SAVE-----------------

wb.save("week_4/spreadsheets/StarWarsVacationSpots.xlsx")
