import pandas as pd
from xml.dom import minidom




doc = minidom.parse('StationMap.xml')
print( doc.nodeName)
print(doc.firstChild.tagName)
StationMap = doc.getElementsByTagName('TelmarId')
stat=(StationMap[0])
print(stat[0])


#dfe = pd.read_excel (r'D:\Projects\python\excel-xml-import\StationCode_Torque Media.xlsx', sheet_name='Torque Media')
#print(dfe.index)
#for x in range(0,245):
#    print(dfe.loc[x].at['Company'], dfe.loc[x].at['Code'])