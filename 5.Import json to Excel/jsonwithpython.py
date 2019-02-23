import xlwings as xw
import matplotlib.pyplot as plt 
import urllib.request
import json
from datetime import time
from datetime import date
from datetime import datetime
import datetime


#grapth function which is to be called from python code

def printResults(data):
    tjson=json.loads(data)

    if "title" in tjson["metadata"]:
        print(tjson["metadata"]["title"])
    count=tjson["metadata"]["count"]
    print("No. of events recorded : " +str(count))
    for i in tjson["features"]:
        t=i["properties"]["time"]
        q=datetime.datetime.fromtimestamp(t / 1e3)
        #print(datetime.datetime.fromtimestamp(t).strftime('%Y-%m-%d %H:%M:%S'))
        print(q)



def geojson():
    urld="https://earthquake.usgs.gov/earthquakes/feed/v1.0/summary/4.5_day.geojson"
    wurl=urllib.request.urlopen(urld)
    print("Result code: " +str(wurl.getcode()))

    if (wurl.getcode()==200):
        data=wurl.read()
        printResults(data)
    else:
        print("Received error")

geojson()
