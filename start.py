"""
module for handleing 6339 Case 1 data
"""
import os
import math
import csv
from xlrd import open_workbook

try:
    import cPickle as pickle
except ImportError:
    import pickle

DIR = '..\\..\\case1\\southShore\\'

class WorkBook(object):
    """wrapper for handling excel workbooks"""
    def __init__(self, file_name, cell_offset, sheet_name=None):
        work_book = open_workbook(file_name)
        if sheet_name is None:
            self.sheet = work_book.sheet_by_index(0)
        else:
            self.sheet = work_book.sheet_by_name(sheet_name)
        self.sheet_as_array = []
        for i in range(cell_offset[0], self.sheet.nrows):
            self.sheet_as_array.append([])
            for j in range(cell_offset[1], self.sheet.ncols):
                self.sheet_as_array[-1].append(self.sheet.cell(i, j).value)
            if all([len(str(x)) < 1 for x in self.sheet_as_array[-1]]):
                del self.sheet_as_array[-1]
                break

    def get_data_as_array(self):
        """"returns data as a list of lists"""
        return self.sheet_as_array


class Inventory(object):
    """class for getting data out of data file"""
    def __init__(self):
        self.data = WorkBook(DIR + 'inv.xlsx', [0, 0]).get_data_as_array()
        part = Inventory.get_header('Product')
        year = Inventory.get_header('Year')
        month = Inventory.get_header('Period')
        inv = Inventory.get_header('Inventory')
        self.data_dict = {}
        for i in self.data:
            if i[part] not in self.data_dict:
                self.data_dict[i[part]] = {}
            self.data_dict[i[part]][(i[month], i[year])] = i[inv]

    @staticmethod
    def get_headers():
        """provdies headerss for inventory file"""
        return ['Product', 'Plant', 'Year', 'Period', 'Inventory', 'Transit', 'Inspection']
    @staticmethod
    def get_header(header):
        """get header index for a field"""
        return Inventory.get_headers().index(header)
    def period_inventory(self, part, month, year):
        """return inventory for a accounting period"""
        try:
            return self.data_dict[part][(month, year)]
        except KeyError:
            return 0

class Sales(object):
    """opens an reads all the sales files"""
    def __init__(self):
        self.file_list = os.listdir(DIR)
        i = 0
        while i < len(self.file_list):
            if self.file_list[i][:12] != 'Data - Sales':
                self.file_list.pop(i)
            else:
                i += 1

        self.data = []
        for i in self.file_list:
            current_file = WorkBook(DIR + i, [15, 6], 'Table')
            self.data = self.data + current_file.get_data_as_array()

    @staticmethod
    def get_headers():
        """column names for the sales file"""
        return ['year', 'week', 'day', 'Plant', 'PlantCity', 'PlantCountry', 'CountryAgain'\
        , 'Region', 'CustomerCity', 'Postal Code', 'IndustryCOde', 'Industry', 'CustomerNumber'\
        , 'Product', 'Description', 'Model', 'Item', 'Qty', 'Price', 'Amount']

    @staticmethod
    def get_header_index(header):
        """returns column index based on header"""
        return Sales.get_headers().index(header)


    def aggeragate_part(self):
        """nested dict of sales volume part -> month,year"""
        aggergate = {}
        part = self.get_header_index('Product')
        day = self.get_header_index('day')
        qty = self.get_header_index('Qty')
        for i in self.data:
            if i[part] not in aggergate:
                aggergate[i[part]] = {}
            if (i[day][6:7], i[day][:4]) not in aggergate[i[part]]:
                aggergate[i[part]][(i[day][6:7], i[day][:4])] = 0
            aggergate[i[part]][(i[day][6:7], i[day][:4])] += int(i[qty])
        return aggergate



class ZipCode(object):
    """class for getting zipcode spatial info"""
    def __init__(self):
        self.load_us_zip_codes()
        self.load_ca_zip_codes()

    def load_us_zip_codes(self):
        """reads us zip codes"""
        self.us_zips = {}
        with open('us_postal_codes.csv') as csv_file:
            reader = csv.reader(csv_file)
            for i in reader:
                self.us_zips[i[1]] = [i[6], i[7]]

    def load_ca_zip_codes(self):
        """reads ca zip codes"""
        self.ca_zips = {}
        with open('ca_postal_codes.csv') as csv_file:
            reader = csv.reader(csv_file)
            for i in reader:
                self.ca_zips[i[0]] = [i[3], i[4]]

    def get_cordinates(self, zipcode):
        """returns tuple of lat and long"""
        try:
            float(zipcode)
            return self.us_get_cordinates(zipcode)
        except ValueError:
            return self.ca_get_cordinates(zipcode)

    def us_get_cordinates(self, zipcode):
        """returns cordinates of us zip codes"""
        return self.us_zips[str(zipcode)]
    def ca_get_cordinates(self, zipcode):
        """returns cordinates of ca zip codes"""
        return self.ca_zips[zipcode[:3]]

    def distance_between_zips(self, zipcode1, zipcode2):
        """get distance between two points"""
        cord1 = self.get_cordinates(zipcode1)
        cord2 = self.get_cordinates(zipcode2)
        return self.distance_on_sphere(cord1[0], cord1[1], cord2[0], cord2[1])

    @staticmethod
    def distance_on_sphere(lat1, long1, lat2, long2):
        """distinace between two cordinates"""
        degrees_to_radians = math.pi/180.0
        phi1 = (90.0 - float(lat1))*degrees_to_radians
        phi2 = (90.0 - float(lat2))*degrees_to_radians
        theta1 = float(long1)*degrees_to_radians
        theta2 = float(long2)*degrees_to_radians
        cos = (math.sin(phi1)*math.sin(phi2)*math.cos(theta1 - theta2)\
        + math.cos(phi1)*math.cos(phi2))
        arc = math.acos(cos)*6371
        return arc


class Warehouse(object):
    """class for stroing warehouse information"""
    ware = [['el paso', 50, 79905], ['ste-croix', 11, 'G0S 2H0'], ['coaticook', 13, 'J1A 2S4']]
    def from_city(self, city):
        """zip code from city name"""
        try:
            city = city.lower()
            return [x[0] for x in self.ware].index(city)
        except ValueError:
            city = city[1:].lower()
            return [x[0] for x in self.ware].index(city)

    def from_factory(self, factory):
        """zip code from factory number"""
        return [x[1] for x in self.ware].index(int(factory))

def dump_forecast(forecast):
    """pickle a forecast class"""
    pickle.dump(forecast, open('forecast.p', 'wb'))

def load_forecast():
    """depickle a forecast class"""
    return pickle.load(open('forecast.p', 'rb'))

class Forecast(object):
    """aggeragtes sales data and returns a forecast"""
    theta = .25
    def __init__(self):
        sales = Sales()
        self.start_year = 2013
        self.aggergated = sales.aggeragate_part()
    def generate_forecast(self, month, year):
        """creates forecast and writes to a csv"""
        data = []
        for key in self.aggergated.keys():
            forecast = self.get_part_forecast(key, month, year)
            data.append(["'" + key, forecast])
        with open('forcast-' + str(month) + '-' + str(year) + '.csv', 'w', newline='') as csv_file:
            writer = csv.writer(csv_file)
            writer.writerows(data)
    def all_forecasts(self):
        """genrerates a csv with all months forecasts"""
        data = []
        for year in range(2013, 2016):
            for month in range(1, 13):
                for part in self.aggergated:
                    forecast = self.get_part_forecast(part, month, year)
                    data.append([year, month, part, forecast])
        return data

    def forecast_verus_inv(self):
        """get comparision of inventory veruses forecast"""
        forecast = self.all_forecasts()
        inventory = Inventory()
        i = 0
        while i < len(forecast):
            forecast[i].append(inventory.period_inventory(forecast[i][2], forecast[i][1],\
            forecast[i][0]))
            i += 1
        with open('inv_forecast.csv', 'w', newline='') as csv_file:
            writer = csv.writer(csv_file)
            writer.writerows(forecast)

    def get_part_forecast(self, part, month, year):
        """creates forcast for part"""
        volume_list = self.volume_list(part, month, year)
        forecast = volume_list[-1]
        for i in range(len(volume_list) - 1, 0, -1):
            forecast = forecast*(1 - self.theta) + volume_list[i]*self.theta
        return forecast


    def volume_list(self, part, month, year):
        """creates list of demand in chronological order"""
        volume_list = []
        while year >= self.start_year:
            volume_list.append(self.volume_lookup(part, month, year))
            if month == 1:
                month = 12
                year -= 1
            else:
                month -= 1
        return volume_list

    def volume_lookup(self, part, month, year):
        """get specfic time period volume"""
        try:
            return self.aggergated[part][(str(month), str(year))]
        except KeyError:
            return 0
