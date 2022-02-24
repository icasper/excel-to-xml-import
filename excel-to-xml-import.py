"""
Script              :   excel-to-xml-import.py
Author              :   Ian Casper
Creation Date       :   23.2.2022
Description         :   Helper script to import data
                        from an excel sheet an XML Document
Change Date         :   24.2.2022
Change Reason       :   Update pandas code to write out matrix
                        from excel to XML file
                        
"""
# Import the Required modules

import pandas as pd
from xml.dom import minidom

# Get the Document Object Model

doc = minidom.parse('StationMap.xml')
print(doc.nodeName)
print(doc.firstChild.tagName)
ArrayOfStationMap = doc.getElementsByTagName(doc.firstChild.tagName)
print(ArrayOfStationMap.length)

# Creat New StationMap

newStation = doc.createElement('StationMap')
newBroadcasterId = doc.createElement('BroadcasterId')
newTelmarId = doc.createElement('TelmarId')
newLanguage = doc.createElement('Language')
doc.firstChild.appendChild(newStation)
newStation.appendChild(newBroadcasterId) 
newStation.appendChild(newTelmarId)
newStation.appendChild(newLanguage)
newBroadcasterId.appendChild(doc.createTextNode('ALGOA'))
newTelmarId.appendChild(doc.createTextNode('ALG'))
newLanguage.appendChild(doc.createTextNode('E'))

# End of New StationMap

# Tidy up the xml with prettyxml

pretty_xml_as_string = doc.toprettyxml()
print(pretty_xml_as_string)

# Write doc back to xml file

with open('StationMap.xml','w') as f:
    f.write(pretty_xml_as_string)

#dfe = pd.read_excel (r'D:\Projects\python\excel-xml-import\StationCode_Torque Media.xlsx', sheet_name='Torque Media')

#for x in range(0,dfe.index.stop):
#    print(dfe.loc[x].at['BroadcasterId'].strip(), dfe.loc[x].at['TelmarId'].strip())
