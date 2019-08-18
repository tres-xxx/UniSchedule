import xlrd
import os.path
import sys

''' How the sheet is organizated
0 - NRC
1 - Sesion
2 - Alfa numerico
3 - Nombre de la asignatura
4 - Opciones disponibles para las electivas y seminarios
5 - Creditos
6 - Dia
7 - Hora inicio
8 - Hora fin
9 - Programa
10 - Jornada
11 - Nivel/Semestre
12 - Modalidad
'''

# Function to pass all the data from the sheet to a local table
def loadTable(sheet):
	tableD = []
	for i in range(0,sheet.nrows):
		column = []
		for j in range(0,sheet.ncols):
			column.append(sheet.cell_value(i,j))
		tableD.append(column)
	return tableD

# Load all asignatures of an specified program
def loadSomething(tableD,sel,opt): # tableD = all the data | sel = string filter | opt = option of the string in the tableD
	# Select all asignatures of an specified program
	program = []
	if(len(tableD) == 0): 
		return program
	if(tableD[0][0] == "NRC"): # Load data for the column
		for i in range(1,len(tableD)):
			if(tableD[i][opt] == sel):
				program.append(tableD[i])
	else:
		for i in range(0,len(tableD)):
			if(tableD[i][opt] == sel):
				program.append(tableD[i])
	return program

# Search for double NRCs
def uniqueNRC(possibleAs):
	possibleNRC = []
	if(len(possibleAs) != 0): 
		possibleNRC.append(possibleAs[0][0])
	for i in range(1,len(possibleAs)):
		different = True
		for j in range(0,len(possibleNRC)):
			if(possibleAs[i][0] == possibleNRC[j]):
				different = False
		if different == True:
			possibleNRC.append(possibleAs[i][0])
	return possibleNRC	
	
# Dummy menu for a menu
def menuSelectSchedule():
	print "Select an option:"
	print "1. Create a new schedule (delete all previous choices)"
	print "2. Modify a previous schedule"
	print "3. Print schedule"
	print "0. Close program"

# Menu to modify asignatures
def menuCRDAsignature():
	print "Select an option:"
	print "1. List all the asignatures"
	print "2. Delete an assignature"
	print "3. Add a new assignature"
	print "0. Create schedule (close menu asignatures)"

# Print all the name of the asignatures in a numeric list
def printAllAsignatures(table):
	if(len(table) == 0):
		print "Empty list"
	else:
		print "List of asignatures"
		for i in range(0,len(table)):
			print (i+1,".",table[i][0][3].encode('utf-8'),"|",table[i][0][10].encode('utf-8'))

# Dummy return option of an asignature
def optionMenuAs(tableD,asignatures = []):
	menuCRDAsignature()
	sel = raw_input("Introduce an option: ")
	asig = []
	if(len(asignatures) != 0):
		asig = asignatures
	while(sel != "0"):
		if(sel == "3"):
			sel = raw_input("Introduce the alpha numeric: ")
			possibleAs = loadSomething(tableD,sel,2)
			sel = raw_input("Choose a time (0-Everything,1-Morning&Afternoon,3-Night) = ")
			while((int(sel) < 0) or (int(sel) > 3)):
				print "No valid option"
				sel = raw_input("Choose a time (0-Everything,1-Morning&Afternoon,3-Night) = ")
			if(int(sel) == 3):
				possibleAs = loadSomething(possibleAs,"NOCHE",10)
			sel = raw_input("Choose a modality (0-Everything,1-Presencial,2-Virtual) = ")
			while((int(sel) < 0) or (int(sel) > 2)):
				print "No valid option"
				sel = raw_input("Choose a modality (0-Everything,1-Presencial,2-Virtual) = ")
			if(sel == "1"):
				possibleAs = loadSomething(possibleAs,"PRESENCIAL",12)
			elif(sel == "2"):
				possibleAs = loadSomething(possibleAs,"VIRTUAL",12)
				
			if(len(possibleAs) == 0): 
				print "No asignature found"
			else:
				asig.append(possibleAs)
				print "Asignature added"
		elif(sel == "1"):
			printAllAsignatures(asig)
		elif(sel == "2"):
			printAllAsignatures(asig)
			if len(asig) != 0:
				sel = raw_input("Introduce number of the asignature to delete: ")
				if(int(sel) < 0 or int(sel) > len(asig)):
					print "Does not exist an asignature with that number"
				else: 
					asig.pop(int(sel)-1)
					print "Asignature has been deleted"
		else: 
			if(sel != "0"): 
				print "No valid option"
		menuCRDAsignature()
		sel = raw_input("Introduce an option: ")
	return asig
	
# Create an schedule
def createAllSchedule(allAsig):
	if(len(allAsig) == 0):
		print "There is no asignatures"
		return []
	nAsig = [] # Vector to select and order
	nNRC = [] # Vector of NRC's
	maxAsig = [] # Maximum of possibilities
	maxsum = 0
	maxAllAsig = len(allAsig)
	for i in range(0, maxAllAsig): # how many asignatures
		nAsig.append(1) # At least one asignature for a program
		nNRC.append(uniqueNRC(allAsig[i])) # List of each NRC (because of a same NRC can be in more than one day)
		maxAsig.append(len(nNRC[i])) # How many NRC 
		maxsum = maxsum + maxAsig[i] # Sum of all the NRC that at the end would exist
	maxorder = 0
	posScheduleNRC = []
	arriveEnd = False
	while(maxorder <= maxsum and not arriveEnd):
		asignatures = []
		for i in range(0, maxAllAsig):
			asignatures.append(nNRC[i][nAsig[i]-1]) # The asignatures for a week
		#--- Schedule :s
		week = [["Monday"],["Tuesday"],["Wednesday"],["Thursday"],["Friday"],["Saturday"]]
		sameTime = False
		for j in range(0, len(asignatures)):
			for l in range(0, len(allAsig)):
				for i in range(0, len(allAsig[l])):
					# Set the day with the assignature
					if(allAsig[l][i][0] == asignatures[j]):
						day = 0
						if(allAsig[l][i][6][:2] == "MA"): day = 1
						elif(allAsig[l][i][6][:2] == "MI"): day = 2
						elif(allAsig[l][i][6][:2] == "JU"): day = 3
						elif(allAsig[l][i][6][:2] == "VI"): day = 4
						elif(allAsig[l][i][6][:1] == "S"): day = 5
						assign = []
						for k in range(0,2): # 0 = begin | 1 = end
							timeA = []
							timeA.append(allAsig[l][i][7+k]) # time start and end
							timeA.append(allAsig[l][i][0]) # NRC
							timeA.append(allAsig[l][i][2]) # Alphanumeric
							assign.append(timeA)
						pos = 1						
						for k in range(1, len(week[day])):
							if(week[day][k][0][0] == assign[0][0] and week[day][k][1][0] == assign[1][0]): # Same time
									sameTime = True
									break
							if(week[day][k][0][0] > assign[0][0]): # start time
								if(week[day][k][1][0] <= assign[1][0] or week[day][k][0][0] < assign[1][0]): # end time
									sameTime = True
								break
							else:
								if(week[day][k][1][0] <= assign[0][0]):
									pos = pos + 1
								else:
									sameTime = True
									break
						if(not sameTime):
							week[day].insert(pos,assign) # insert a possible asignature in an specified time (pos)
					if(sameTime): break
				if(sameTime): break
			if(sameTime): break			
		if(not sameTime): # Check that there is not a same asignature at the same time
			posScheduleNRC.append(week) # Posibility for an organized week
		nAsig[0] = nAsig[0] + 1
		if(maxorder == maxsum): arriveEnd = True # End of the cycle
		maxorder = 0
		for i in range(0, maxAllAsig): # it comes back to 0 .. 0 after arrives to the maximum sum
			if(nAsig[i] > maxAsig[i] and i != (maxAllAsig-1)): # next cicle
				nAsig[i] = 1 
				nAsig[i+1] = nAsig[i+1] + 1
			maxorder = maxorder + nAsig[i] # It has to be used when is a new cicle
	return posScheduleNRC

# Print all the possibilities of the schedules
def printSchedule(schedule):
	if(len(schedule) > 0):
		for k in range(0,len(schedule)):	
			print "----------------------------------------------------------"
			print "Option ", k+1
			print "----------------------------------------------------------"
			for i in range(0,len(schedule[k])):
				for j in range(0,len(schedule[k][i])):
					print schedule[k][i][j]
	else:
		print "There is no any option for an schedule with the specified asignatures"
	print "----------------------------------------------------------"
	print "----------------------------------------------------------"
	
# Create an schedule with 0 assignatures previously choosed
def scratchSchedule(tableD):
	allAsig = optionMenuAs(tableD)
	possibleSchedule = createAllSchedule(allAsig)
	printSchedule(possibleSchedule)
	return allAsig	

sel = raw_input("Introduce the direction of the file (example: /home/myname/Documents/doc.xlsx) = ") # direction of the file
if os.path.isfile(sel):
	wb = xlrd.open_workbook(sel) # Open workbook
else:
	sys.exit("No file has been found. The program is going to be closed it")
sheet = wb.sheet_by_index(1) # this is the sheet with all the asignatures
tableD = loadTable(sheet) # load table
#program = loadSomething(tableD, "TGSC",9) # load all asignatures of with program = TGSC
#possibleAs = loadSomething(tableD, selAsig,2) # List of all the possibilities of an asignature
#possibleNRC = uniqueNRC(possibleAs) # List of all the NRC with no repetition
sel = "1"
print "*************"
print "Create an schedule for your schedule time in Uniminuto University"
print "*************"
asignatures = scratchSchedule(tableD)
while(sel != "0"):
	menuSelectSchedule()
	sel = raw_input("Introduce an option: ")
	if(sel == "1"):
		print "*************"
		print "Choose your options again"
		print "*************"
		asignatures = scratchSchedule(tableD)
	elif(sel == "2"):
		asignatures = optionMenuAs(tableD,asignatures)
	elif(sel == "3"):
		possibleSchedule = createAllSchedule(asignatures)
		printSchedule(possibleSchedule)
	else: 
		print "Bye!"
		break
