"""
Script              :   excel-to-xml-import.py
Author              :   Ian Casper
Creation Date       :   23.2.2022
Description         :   Helper script to import data
                        from an excel sheet an XML Document
Change Date         :   24.2.2022
Change Reason       :   Update pandas code to write out array
                        from excel to XML file

"""

# I M P O R T A N T   I N F O R M A T I O N ! ! !

"""

1. Either make sure that the XML file to be parsed has the correct name and path. Or change the name and path in line as indicated

2. Either make sure that the Excel file to be parsed has the correct name, path and sheet name. Or change the name,path and sheet name 
   in line as indicated.

3. Make sure that the excel sheet has the following column headers:-
   BroadcasterId, TelmarId

"""


# Import the Required modules

import pandas as pd
from xml.dom import minidom

# Get the Document Object Model

doc = minidom.parse('StationMap.xml') # Change XML Path and filen name here if it is different than this one
print(doc.nodeName)
print(doc.firstChild.tagName)
ArrayOfStationMap = doc.getElementsByTagName(doc.firstChild.tagName)
print(ArrayOfStationMap.length)

# Get Excel sheet into pandas array

dfe = pd.read_excel (r'D:\Projects\python\excel-xml-import\StationCode_Torque Media.xlsx', sheet_name='Torque Media') # Change Excel path, file name and sheet name here

# Iterate through pandas array and update each new StationMap created in XML DOM

for x in range(0,dfe.index.stop):

# Creat New StationMap

    newStation = doc.createElement('StationMap')
    newBroadcasterId = doc.createElement('BroadcasterId')
    newTelmarId = doc.createElement('TelmarId')
    newLanguage = doc.createElement('Language')
    doc.firstChild.appendChild(newStation)
    newStation.appendChild(newBroadcasterId) 
    newStation.appendChild(newTelmarId)
    newStation.appendChild(newLanguage)
    newBroadcasterId.appendChild(doc.createTextNode(dfe.loc[x].at['BroadcasterId'].strip()))
    newTelmarId.appendChild(doc.createTextNode(dfe.loc[x].at['TelmarId'].strip()))
    newLanguage.appendChild(doc.createTextNode('E'))

# End of New StationMap

# Tidy up the xml with prettyxml

pretty_xml_as_string = doc.toprettyxml()
print(pretty_xml_as_string)

# Write doc back to xml file

with open('StationMap.xml','w') as f:
    f.write(pretty_xml_as_string)
