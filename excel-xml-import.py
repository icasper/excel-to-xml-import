import pandas as pd
from xml.dom import minidom

mydoc = minidom.parse('StationMap.xml')

item = mydoc.getElementsByTagName('StationMap')


print(item[1].firstChild.data)


#dfe = pd.read_excel (r'D:\Projects\python\excel-xml-import\StationCode_Torque Media.xlsx', sheet_name='Torque Media')
#print(dfe.index)
#for x in range(0,245):
#    print(dfe.loc[x].at['Company'], dfe.loc[x].at['Code'])