import pandas as pd

import openpyxl
print(openpyxl.__version__)

#import xlsxwriter
# fName = sys.argv[1]
fName = "DiscountGrowers_MyView1.xls"
df = pd.read_excel(fName)
#print the column names
df.rename(columns=lambda x: x.replace("\r","").replace("\n","").replace(" ","").replace("%","").replace("$","").replace("/","").replace("-","").replace("+",""), inplace=True)
df['EPS1YrGrowth'] = (df["MeanEPSEstNextYear"] - df["MeanEPSEstThisYr"])/df["MeanEPSEstThisYr"]
df['SustainableGrowth'] = df['ROETTM']/100*(1-(df['PayoutRatioTTM']/100))
df['CoreGrowth'] = df['SustainableGrowth']
df['DivYield']=df['DividendYieldTTM']/100
df['CostOfGrowth'] = df['CoreGrowth']/df['ROETTM']/100*df['EPSTTM']
df['SurplusEarnings'] = df['EPSTTM'] - (df['DividendAmount']/df['SharesHeld'])-df['CostOfGrowth']
df['ShareShrink'] = df['SurplusEarnings']/df['CurrentPrice']
df['ddrm'] = df['DivYield']+df['CoreGrowth']+df['ShareShrink']
df = df.drop('DividendYieldTTM', 1)

outputName = "Processed.xlsx"
writer = pd.ExcelWriter(outputName, engine='xlsxwriter')
df.to_excel(writer, "Main")
workbook  = writer.book
# Add a format. Light red fill with dark red text.
redFormat = workbook.add_format({'bg_color': '#FFC7CE',
                               'font_color': '#9C0006'})

# Add a format. Green fill with dark green text.
greenFormat = workbook.add_format({'bg_color': '#C6EFCE',
                               'font_color': '#006100'})
currencyFmt = workbook.add_format({'num_format': '$#,##0.00'})
pctFmt = workbook.add_format({'num_format': '0.00%'})
worksheet = writer.sheets["Main"]
worksheet.set_column('D:G', 18, currencyFmt)
worksheet.set_column('M:M', 18, currencyFmt)
worksheet.set_column('N:N', 18, currencyFmt)
worksheet.set_column('Z:AA', 18, currencyFmt)
worksheet.set_column('AC:AD', 18, currencyFmt)
worksheet.set_column('AH:AH', 18, currencyFmt)
worksheet.set_column('AR:AR', 18, pctFmt)
worksheet.set_column('AK:AN', 18, pctFmt)

# Dividend eval
worksheet.conditional_format('B3:K12', {'type': 'cell',
                                         'criteria': 'between',
                                         'minimum': 30,
                                         'maximum': 70,
                                         'format': format1})

worksheet.conditional_format('AR2:AR99', {'type':     'cell',
                                        'criteria': '<',
                                        'value':    .09,
                                        'format':   redFormat})
worksheet.conditional_format('AR2:AR99', {'type':     'cell',
                                        'criteria': '>',
                                        'value':    .12,
                                        'format':   greenFormat})
writer.save()


#df2 = pd.read_excel(outputName, sheetname="Main")
#writer2 = pd.ExcelWriter("Test.xlsx", engine='xlsxwriter')
#df2.to_excel(writer2, "Main")




#worksheet = writer2.sheets["Main"]
#worksheet.set_column('AQ:AR', 18, pctFmt)
#worksheet.conditional_format('AR2:AR14', {'type': '3_color_scale'})
#writer2.save()

#worksheet.conditional_format('AR2:AR14', {'type': '3_color_scale'})
#workbook  = writer.book
#worksheet = writer.sheets['Main']
#worksheet.conditional_format('AR0:AR255', {'type': '3_color_scale'})
#writer.save()
