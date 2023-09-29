# StarWars Travel

Using Python and (openpyxl) )Excel, this app generates data about the fictional Planets from the Star Wars universe to help the user decide which Planet they'd like to visit, how they would get there, and how they would get around on said planet. The data from the [API](https://swapi.dev/documentation) was last updated in 2014.

Column names are indicated within [brackets].

API Documentation: https://swapi.dev/documentation

This project was a part of the Vets in Tech Python Course. Notes were made within the code to indicate the collaboration of others. All of the code which doesn't include a comment indicating collaboration is of my creation.

-----------------DOCUMENTATION-----------------

**Worksheet 1: Star Wars Vacation Spots**

*Which is the best Planet to vacation on in the Star Wars universe?*

Data shows details about each Planet to make the decision process easier.

Just like planning any trip, it's helpful to see your destination options [Planet], the Climate [Climate], Terrain [Terrain], and Population [Population]. Because it is Star Wars, I added the number of films that said Planet featured on [Film Count]. The user may want to visit a planet where many films took place so see some famous sights. Or the user may want to avoid planets with crowds and where filming took place so that they can enjoy their vacation in relaxation mode. The goal is to provide the user with information relevent to making an informed decision.

**Worksheet 2: Getting There | Starships**

*What is the best way to get to my destination planet? Which Starship is right for me?*

Data shows details about each Starship to make the decision process easier. The end user would ideally be able to make an informed decision regarding their transportation to the Planet they selected. In this fictional universe, the data was limited; therefore, it is assumed that all  Starships travel to all planets.

Within this dataset, users will only see Starships that allow passengers. Users can view the name of the Starship [Starship], the type of starship [Type of Ship], its maximum speed [Maximum Speed], the total amount of potential seating [Total Seats], the size of the crew [Crew Size], the size of the ship [Ship Size] (which was initially in meters but using logic, the ships are categorized as small, medium, large, and massive), and the experience they can expect to have on the Starship [Passenger Experience]. The passenger experience was calculated using a cruise ship experience ratio where 1:1 is defined as a luxurious cruise. The user will see the type of experience they can expect from said Starship (Luxury, Economy, Comfort, and Mystery Experience; which means that the API had missing data.

**Worksheet 3: Getting Around | Vehicles**

*How will I get around on my choice planet? Which Vehicle is right for me?*

Data shows details about each Vehicle to help the user make an informed decision.

Within the dataset, the user will see a list of vehicle options [Vehicle], the Price of the Vehicle (in credits, so there is no $ symbol) [Price], the maximum amount of passengers allowed [Maximum Group Size], the number of crew members [Crew Size], the size of the ship was in meters and using logic, it is arranged in a small, medium, and large output [Vehicle Size], along with the expected passenger experience [Passenger Experience]. The same cruise ship logic was applied to the passenger experience to provide the user with the experience they can expect when traveling around their chosen planet.

---------------------GUI-----------------------

Now that I have the data, how will the end user access the data in a user-friendly way? A Graphical User Interface is the answer, of course. For this next piece of the project, I chose to use the [tkinter](https://www.tutorialspoint.com/python/python_gui_programming.htm) Python module (the specific tutiorial that was used is linked).

**Goals:**

1. Upon accessing the GUI, the user is prompted with a short welcome message.

   1. Problems Encountered:
   2. Solutions:
2. Users will start by choosing their desiderd **planet**. With 60 planet options, this can be a difficult decision. The users will be able to decide which feature is most important to them and what about that feature is most important so that the GUI can generate a few (ideally) planet options for them to choose from.

   1. Problems Encountered:
   2. Solutions:
   3. Additional Notes:
      1. In a perfect world, the user would then be prompted to select more features to filter through but at this time we are focusing on one feature at a time.
      2. For example: Climate. The user will select Climate as the deciding factor and then which type of climate they are looking to visit. Then a list of planets with said climate will generate for the user.
3. Next, the users will be needing to get to their desired planet. There are 36 **Starships** to choose from. Users will choose which feature is most important to them and then further define the parameters of said feature in order to generate a list of Starship options for their trip.

   1. Problems Encountered:
   2. Solutions:
4. Finally, the GUI will help users make an informed decision about how they will get around the chosen planet. As above, the users will select the feature they care about the most and the parameters of said feature in order to generate a list of Vehicle options.

   1. Problems Encountered:
   2. Solutions:

This is, of course, a fictional universe and the data is provided by the StarWars [API](https://swapi.dev/documentation) - last updated in 2014. This project was really fun to make and I enjoyed playing with the data in unexpected ways.
