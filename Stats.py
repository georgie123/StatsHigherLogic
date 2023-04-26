
# Le script extrait les stats AMS à partir d'un export CSV des tables customer_ams et customer.

import os
from datetime import date, datetime
from tabulate import tabulate as tab
import pandas as pd
import matplotlib as mpl
import matplotlib.pyplot as plt
from matplotlib.patches import Polygon
from matplotlib.collections import PatchCollection
from mpl_toolkits.basemap import Basemap
import numpy as np
from PIL import Image, ImageOps
import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
import warnings
warnings.simplefilter(action='ignore', category=UserWarning)

today = date.today()

shp_simple_countries = r'C:/Users/Georges/OneDrive/_data/simple_countries/simple_countries'
shp_simple_areas = r'C:/Users/Georges/OneDrive/_data/simple_areas/simple_areas'
inputCountryConversion = r'C:/Users/Georges/OneDrive/_data/countries_conversion.xlsx'
inputSpecialtyConversion = r'C:/Users/Georges/OneDrive/_data/specialties_conversion.xlsx'

workDirectory = r'C:/Users/Georges/Downloads/'

dfAms = pd.read_csv(workDirectory + 'customer_ams.csv', sep=';', encoding='cp1252',
                            usecols=['id', 'customer_id', 'email', 'role'])

dfCustomer = pd.read_csv(workDirectory + 'customer.csv', sep=';', encoding='cp1252',
                            usecols=['id_customer', 'email', 'speciality', 'country'])
dfCustomer.rename(columns={'email': 'email_customer'}, inplace=True)

# print('\ncustomer_ams')
# print(tab(dfAms.head(10), headers='keys', tablefmt='psql', showindex=False))
# print(dfAms.shape[0])

# print('\nCustomer')
# print(tab(dfCustomer.head(10), headers='keys', tablefmt='psql', showindex=False))
# print(dfCustomer.shape[0])

# JOIN CUSTOMER_AMS AND CUSTOMER
df = pd.merge(dfAms, dfCustomer, left_on='customer_id', right_on='id_customer', how='left')[['id', 'role', 'email', 'speciality', 'country']]

print('\nAMS')
print(tab(df.head(10), headers='keys', tablefmt='psql', showindex=False))
print(df.shape[0])
