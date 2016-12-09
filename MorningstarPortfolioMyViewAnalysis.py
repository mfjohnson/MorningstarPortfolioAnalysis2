import pandas
import sys

#import xlsxwriter
fName = sys.argv[1]
df = pandas.read_excel(fName)
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


#writer = pandas.ExcelWriter('CurrentHoldingsProcessed.xlsx',engine='xlsxwriter')

writer = pandas.ExcelWriter("Processed-"+fName)
df.to_excel(writer, "Main")
#workbook  = writer.book
#worksheet = writer.sheets['Main']
#worksheet.conditional_format('AR0:AR255', {'type': '3_color_scale'})
writer.save()
