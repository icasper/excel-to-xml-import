import pandas as pd

dfe = pd.read_excel (r'D:\Projects\python\excel-xml-import\StationCode_Torque Media.xlsx', sheet_name='Torque Media')
dfx = pd.read_xml (r'D:\Projects\python\excel-xml-import\tits.xml')
print(dfe.loc[1].at['Company'])
print(dfe.loc[1].at['Code'])
print(dfe.index)
print (dfx.index)
for x in range(0,245):
    print(dfe.loc[x].at['Company'], dfe.loc[x].at['Code'])
for y in range(0, 3):
    print(dfx.loc[y].at['n35237'])
