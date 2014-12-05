export = open('C:/Users/BeamaR01/Dropbox/Compass/Agilysys Formatting Project/MI_Exp.txt')
output = open('C:/Users/BeamaR01/Dropbox/Compass/Agilysys Formatting Project/MI_Imp.txt', 'x')
priceLevels = open('C:/Users/BeamaR01/Dropbox/Compass/Agilysys Formatting Project/Price_Levels.txt')
items = []
newPriceMap = {}
for line in priceLevels:
	values = line.split(',')
	iD = values[0]
	prices = values[2]
	prices = prices.rstrip()
	prices = prices.replace(";",",")
	newPriceMap[iD] = prices
	updateString = '"U",' + str(iD) + ",,,,," + prices + ",,,,,,,,,,,,,,,,,\r\n"
	output.write(updateString)
print("completed")