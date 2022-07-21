import pandas as pd

resultFile = open('./CompareQuotes/Results/results.txt', 'w')

readOld = pd.read_excel(f'./CompareQuotes/Quotes/OLD.xlsx', engine='openpyxl', skiprows=8, usecols=['Manufacturer Part #', 'Qty'])
xl = pd.ExcelFile('./CompareQuotes/Quotes/OLD.xlsx')
print(xl.sheet_names[0], file=resultFile)
readOld.dropna(inplace=True)
#print(readOld.to_string())

readNew = pd.read_excel(f'./CompareQuotes/Quotes/NEW.xlsx', engine='openpyxl', skiprows=8, usecols=['Manufacturer Part #', 'Qty'])
#print(readNew.columns)
readNew.dropna(inplace=True)

groupOld = readOld.groupby(['Manufacturer Part #']).sum()
group = open('./CompareQuotes/Groups/groupsOld.txt', 'w')
print(groupOld.to_string(), file=group)

groupNew = readNew.groupby(['Manufacturer Part #']).sum()
group = open('./CompareQuotes/Groups/groupsNew.txt', 'w')
print(groupNew.to_string(), file=group)

resultFile = open('./CompareQuotes/Results/results.txt', 'a')
merged = groupOld.merge(groupNew, on='Manufacturer Part #', how='outer', sort=True, suffixes=['Old', 'New']).sort_index(axis=1)
merged = merged.loc[merged['QtyNew'] != merged['QtyOld']]
print(merged.to_string(), file=resultFile)