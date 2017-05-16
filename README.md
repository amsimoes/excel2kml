# excel2kml
Python (v3.5) script to generate a KML file to mark points on googlemaps from an excel sheet, with custom icons' colors.


# Dependencies
* `# pip3 install openpyxl`
* `# pip3 install unidecode`

# Instructions
Firstly, there has to be a column with the places' names as you provide their range on the script.
Then, point the column which have the respective pair of (Latitude,Longitude) coordinates.


Finally assign whether there is a colors column for the markers if you intend to have custom ones. 
(Attention: The default color is RED for no custom colors)


![First step](https://i.gyazo.com/00b779e0d4e41fe8c5fe5e158cdf0e57.png)

If everything worked as intended you shall have the kml file fully functional waiting to be imported. 
Access "My Maps", and in any map on a empty layer, you'll see the Import button. 
Just upload the kml file and it will mark every point, with it's color, on the right spot on the map.


![Second step](https://i.gyazo.com/501771cab28420cea55d6421388af6d3.png)

Voilá!


![Final](https://i.gyazo.com/d604070a5af369a63f94a4b5a56bfb4c.png)

# TODO

[TODO]: 
1. Add more colors possibilities. 
2. Option to have column with points descriptions. 
3. Point if user have Lat,Long or Long,Lat format.
4. If cell is blank on colors column -> default color. 
5. Get user input for default color as well as Input for layer's name and Map's name.
