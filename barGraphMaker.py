import xlrd

xl_workbook = xlrd.open_workbook('data.xlsx')	# Open the workbook
xl_sheet = xl_workbook.sheet_by_index(0)	# Select the zero-indexed sheet

val_list1=[int(item) for item in xl_sheet.col_values(1)[1:]]	# List of the vals in first column from the sheet
val_list2=[int(item) for item in xl_sheet.col_values(2)[1:]]	# List of the vals in the second column from the sheet
y_list=val_list1+val_list2
years=[int(year) for year in xl_sheet.col_values(0)[1:]]
years=sorted(years)

y_values=set(y_list)
y_values=sorted(list(y_values),reverse=True)
x_values=set(years)
x_values=sorted(list(x_values),reverse=True)
x_indent=" "*max([int(len(str(item)))+2 for item in y_list])	# spacing for the x-axis to be placed properly after the y-axis, accounting for any y-axis size

vals_to_place=[]	# This will be a list of all of the values from val_list1 and val_list2 with this pattern: val_list1[0], val_list2[0], val_list1[1], val_list2[1], etc.
vals_to_string=[]	# Every pair of values in this list represent the indexes for the + and ~ bars for each year/x-coordinate (indexes 0-1: first x-coordinate, indexes 2:3: second x-coordinate)

def search(lst, value):		# Search for a value in a given list
    return [i for i, ltr in enumerate(lst) if ltr == value]

print "___________________________________"
print "|             Key:                |"
print "| The || bar symbolizes {}    |".format(xl_sheet.cell_value(0,1).encode('utf-8'))
print "| The II bar symbolizes {}   |".format(xl_sheet.cell_value(0,2).encode('utf-8'))
print "|_________________________________|\n"

############################# MAIN ############################################

i=0
while i<len(val_list1) and i<len(val_list2):	# Setting both vals_to_place and vals_to_string
	vals_to_place.append(val_list1[i])
	vals_to_place.append(val_list2[i])
	vals_to_string.append(None)
	vals_to_string.append(None)
	i+=1

largest_yval_len=max([len(str(y_value)) for y_value in y_values])
for y_value in y_values:
	barsLine=" "	# This string represents where the + and ~ bars will be placed on the bar chart
	if len(str(y_value))==largest_yval_len:
		x_coordinate = "{}-|".format(y_value)
	else:
		x_coordinate = "{}{}-|".format(" "*(largest_yval_len-len(str(y_value))),y_value)

	for item in search(vals_to_place,y_value):	# Checks for the index where y_value is in val_to_place
		vals_to_string[item]=y_value
	i=0
	while i < len(vals_to_string):
		if type(vals_to_string[i])==int and type(vals_to_string[i+1])==int:
			barsLine+="|| II   "
			i+=2
		elif type(vals_to_string[i])==int and not type(vals_to_string[i+1])==int:
			barsLine+="||      "
			i+=2
		elif not type(vals_to_string[i])==int and type(vals_to_string[i+1])==int:
			barsLine+="   II   "
			i+=2
		else:
			barsLine+="        "
			i+=2
	print x_coordinate+barsLine

print x_indent+"________"*len(set(years))	# Print the line for the x-axis

indented=False
for item in years:		# Print the x-coordinate values with proper indentation
	if indented==True:
		print str(item)+"   ",
	else:
		indented=True
		print x_indent+" "+str(item)+"   ",