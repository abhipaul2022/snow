import pandas as pd
from openpyxl import load_workbook
import openpyxl as pxl
from datetime import date,datetime,timedelta
from openpyxl.styles import Font,Alignment
import requests

url = "https://ibm.ent.box.com/folder/159988225988?id=abhipaul@in.ibm.com"
r = requests.get(url)
open('My PIR Report (MC45ODEzMDUwMA).xlsx', 'wb').write(r.content)
df = pd.read_excel('My PIR Report (MC45ODEzMDUwMA).xlsx')

#urllib.urlretrieve(filep, "My PIR Report (MC45ODEzMDUwMA).xlsx") 



   
