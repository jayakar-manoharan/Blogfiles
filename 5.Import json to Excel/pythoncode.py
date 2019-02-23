# pythoncode.py
# importing the required module 
import xlwings as xw
import urllib.request
import json
from datetime import time
from datetime import date
from datetime import datetime
import datetime

#geojson function which is to be called from python code
def geojson():
    #set the active sheet(The Excel file with the data needs to be open and active)
    sht = xw.sheets.active

    #Set the url ID
    urld="https://earthquake.usgs.gov/earthquakes/feed/v1.0/summary/4.5_day.geojson"
    #open the url ID
    wurl=urllib.request.urlopen(urld)
    
    #Get the status of the connection estblished (html page returns the value 200 after succesful connection)
    if (wurl.getcode()==200):

            # Read the URL data 
            data=wurl.read()

            # Assign url data to a JSON variable
            tjson=json.loads(data)

            #Gets the title from the json object
            if "title" in tjson["metadata"]:
                
                #Assign Title data to A1
                sht.range(1,1).value=(tjson["metadata"]["title"])
                #Set A1 cell colour to Yellow (R,G,B)=(255,255,51)                
                sht.range(1,1).color=(255,255,51)

                #Assign cout value to a variable
                count=tjson["metadata"]["count"]

                #Assign the count value to cell A2 and colour to Red (R,G,B)=(255,0,0)                
                sht.range(1,2).value=("No. of events recorded : " +str(count))
                sht.range(1,2).color=(255,0,0)

                #Assign valuse to the Respective fields and Set Row colour to Blue (R,G,B)=(135,206,250)             
                sht.range(2,1).value=("Date/Time")
                sht.range(2,1).color=(135,206,250)
                
                sht.range(2,2).value=("Latitude")
                sht.range(2,2).color=(135,206,250)

                sht.range(2,3).value=("Longitude")
                sht.range(2,3).color=(135,206,250)

                sht.range(2,4).value=("Altitude")
                sht.range(2,4).color=(135,206,250)

                sht.range(2,5).value=("Location")
                sht.range(2,5).color=(135,206,250)

                sht.range(2,6).value=("Magnitude")
                sht.range(2,6).color=(135,206,250)

                
                #Loop for Time property in Features json object and assign to corresponding field to excel             
                for i,item in enumerate(tjson["features"]):
                  t=item["properties"]["time"]
                  q=datetime.datetime.fromtimestamp(t / 1e3)
                  sht.range(i+3,1).value=(q)

                #Loop for Co-ordinates property in Geometry object and assign to corresponding field to excel
                #cord=(Lattitude,Longitude,Altitude)             
                for i,item in enumerate(tjson["features"]):
                   cord=item["geometry"]["coordinates"]
                   sht.range(i+3,2).value=(cord)

                #Loop for place property in Features object and assign to corresponding field to excel             
                for i,item in enumerate(tjson["features"]):
                   place=item["properties"]["place"]
                   sht.range(i+3,5).value=(place)

                #Loop for magnitude property in Features object and assign to corresponding field to excel             
                for i,item in enumerate(tjson["features"]):
                   mag=item["properties"]["mag"]
                   sht.range(i+3,6).value=(mag)
 
                #Autofit the cells Height and width as per the datas in the Excel Rows/Column
                sht.range('A1:FF1048576').autofit()
        
    else:
        #If connection not established then show error    
        print("Received error")

