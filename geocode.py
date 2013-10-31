import xlrd
import requests
from time import sleep
import json
from slugify import slugify

class Event:
    def __init__(self, evtype, address, count):
        self.type = slugify(evtype)
        self.address = address
        self.count = count

    def gettype(self):
        return self.type

    def getaddress(self):
        return self.address


    def getcount(self):
        return self.count

    def geocode(self):
        print("Geocoding '{}'... ".format(self.address)),

        res = requests.get("http://maps.googleapis.com/maps/api/geocode/json?address={}&sensor=false".format(self.address+" Brunswick ME 04011")).json()

        try:
            self.lat = res.get("results")[0].get("geometry").get("location").get("lat")
        except:
            self.lat = 0
        try:
            self.lon = res.get("results")[0].get("geometry").get("location").get("lng")
        except:
            self.lon = 0

        print ("got ({}, {})").format(self.lat, self.lon)

    def getlat(self):
        try:
            return self.lat
        except:
            return 0

    def getlon(self):
        try:
            return self.lon
        except:
            return 0

    def __str__(self):
        return "{count} {type} events at {address} ({lat}, {lon})".format(count=self.getcount(), type=self.gettype(), address=self.getaddress(), lat=self.getlat(), lon=self.getlon())


events = []

sheet0 = xlrd.open_workbook("ron_data.xlsx", encoding_override="cp1252").sheet_by_index(0)

col1 = sheet0.col(0)
col2 = sheet0.col(1)
col3 = sheet0.col(2)

for row1, row2, row3 in zip(col1[1:], col2[1:], col3[1:]):
    event = Event(row1.value, row2.value, row3.value)
    events.append(event)

data = []

for event in events:
    sleep(0.35)
    event.geocode()
    data.append({"lat": event.getlat(), "lng": event.getlon(), "count": event.getcount(), "type": event.gettype()})

json.dump(data, open("data.json", "w"))

