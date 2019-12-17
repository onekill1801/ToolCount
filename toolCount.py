import xlrd 
import time
import openpyxl

def SampleForm(ws):
	ws.cell(1,1).value = "Event"
	ws.cell(1,2).value = "Detected on"
	ws.cell(1,3).value = "Group"
	ws.cell(1,4).value = "Device"
	ws.cell(1,5).value = "Task"
	ws.cell(1,6).value = "Description"
	ws.cell(1,7).value = "IP address"

def SampleFormOne(ws):
	ws.cell(1,1).value = "Event"
	ws.cell(1,2).value = "Detected on"
	ws.cell(1,3).value = "Group"
	ws.cell(1,4).value = "Device"
	ws.cell(1,5).value = "Task"
	ws.cell(1,6).value = "Description"
	ws.cell(1,7).value = "IP address"
	ws.cell(1,8).value = "IP address Attack"

def Create_Sheet_New(filename,eventList,deviceList):		
	wb = openpyxl.load_workbook(filename)
	for index in range(0,len(eventList)):
		nameSheet = eventList[index]
		if len(nameSheet) > 31:
			nameSheet = nameSheet[0:31]
		wb.create_sheet(nameSheet) # Sheet Name
		ws = wb.worksheets[index+3] 		   # Sheet Index
		listDevice_OneEvent = deviceList[index]
		SampleFormOne(ws)
		for x in range(0,len(listDevice_OneEvent)):
			ws.cell(x+2,1).value = listDevice_OneEvent[x][0]
			ws.cell(x+2,2).value = listDevice_OneEvent[x][1]
			ws.cell(x+2,3).value = listDevice_OneEvent[x][2]
			ws.cell(x+2,4).value = listDevice_OneEvent[x][3]
			ws.cell(x+2,5).value = listDevice_OneEvent[x][4]
			ws.cell(x+2,6).value = listDevice_OneEvent[x][5]
			ws.cell(x+2,7).value = listDevice_OneEvent[x][6]
	wb.save(filename)

def EventCount(filename, eventList):
	loc = (filename) 
	wb = xlrd.open_workbook(loc)
	sheet = wb.sheet_by_index(1) 
	listEventName = []
	for i in range(0,len(eventList)):
		tempList = []
		for r in range(1,sheet.nrows):
			s_value = sheet.cell_value(r,2)
			if s_value == eventList[i] and sheet.cell_value(r,5) != "Managed devices":
				tempList.append([s_value,sheet.cell_value(r,3),sheet.cell_value(r,5),sheet.cell_value(r,6),sheet.cell_value(r,7),sheet.cell_value(r,8),sheet.cell_value(r,9)])
		listEventName.append(tempList)
	return listEventName

def EventList(filename):
	loc = (filename) 
	wb = xlrd.open_workbook(loc)
	sheet = wb.sheet_by_index(1) 
	# print(len(wb.sheet_names()))sheet.nrows
	# print(sheet.nrows)
	listEvent = []
	for r in range(1,sheet.nrows):
		if sheet.cell_value(r,5) != "Managed devices":
			s_value = sheet.cell_value(r,2)
			listEvent.append(s_value)
	listEvent = set(listEvent)
	listEvent = list(listEvent)
	listEvent.sort()
	# print(type(listEvent))
	# print(listEvent)
	# print(len(listEvent))
	return listEvent

def sortEventList(eventList):
	j = 0
	for i in range(0,len(eventList)):
		# 	Network attack detected , Malicious object detected ,Disinfection impossible
		if eventList[i] == "Network attack detected":
			temp = eventList[j]
			eventList[j] = eventList[i]
			eventList[i] = temp
			j=j+1
	for i in range(0,len(eventList)):
		# 	Network attack detected , Malicious object detected ,Disinfection impossible
		if eventList[i] == eventList[i] == "Malicious object detected":
			temp = eventList[j]
			eventList[j] = eventList[i]
			eventList[i] = temp
			j=j+1
	for i in range(0,len(eventList)):
		# 	Network attack detected , Malicious object detected ,Disinfection impossible
		if eventList[i] == eventList[i] == "Disinfection impossible":
			temp = eventList[j]
			eventList[j] = eventList[i]
			eventList[i] = temp
			j=j+1
	return eventList

def main(filename):
	eventList = sortEventList(EventList(filename))
	# print(len(eventList))
	# print(eventList)
	# print(eventList[0])
	deviceList = EventCount(filename,eventList)
	# print(deviceList)
	# print(deviceList[10])
	Create_Sheet_New(filename,eventList,deviceList)

if __name__== "__main__":
	print("Library needs to be installed: xlrd, openpyxl.")
	print("Enter filename (no write .xlsx) : (Only run filetype .xlsx)")
	# fileName=input() 
	fileName = "E:\\VNPT Security Inter\\ALL\\12-17"
	fileName = fileName + ".xlsx"
	start_time = time.time()
	# print(fileName)
	# EventList(fileName)
	main(fileName)
	end_time = time.time()
	print("Time : ", (end_time - start_time), "s")