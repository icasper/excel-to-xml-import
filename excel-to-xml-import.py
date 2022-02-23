import pandas as pd
from xml.dom import minidom

dom = minidom.parse('StationMap.xml')
stationMap = dom.getElementsByTagName('StationMap')
for x in range(0,stationMap.length):
    print(stationMap[x].toxml())

dfe = pd.read_excel (r'D:\Projects\python\excel-xml-import\StationCode_Torque Media.xlsx', sheet_name='Torque Media')

for x in range(0,dfe.index.stop):
    print(dfe.loc[x].at['Company'].strip(), dfe.loc[x].at['Code'].strip())
