import pandas as pd

xls = pd.ExcelFile('BOM file for Data processing.xlsx')
df = pd.read_excel(xls, 'Source')
groupedItems = df.groupby('Item Name')['Item Name'].count()
output = {}

for item in groupedItems.items():
	output[item[0]] = {
		'rawMaterial': [['#', 'Item Description', 'Quantity', 'Unit']],
		'finalProdcut': [['#', 'Item Description', 'Quantity', 'Unit'], [1, item[0], 1, 'Pc']]
	}

for item in df.itertuples():
	index, itemName, level, rawMaterial, quantity, unit = item
	if level == '.1':
		currentRawMaterial = output[itemName]['rawMaterial']
		output[itemName]['rawMaterial'].append([len(currentRawMaterial) + 1, rawMaterial, quantity, unit])

with pd.ExcelWriter('output.xlsx') as writer:
	for key, value in output.items():
		outputDF = pd.DataFrame(
			[['Finished Good List']] +
			value['finalProdcut'] +
			[['End of FG']] +
			[['Raw Material List']] +
			value['rawMaterial'] +
			[['End of RM']]
		)
		outputDF.to_excel(writer, sheet_name=key, header=False, index=False)
