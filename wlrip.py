####WLrip.py v0.2
###By Barnaby Skeggs
##Parses the Windows Search Index / TextHarvester data store file 'WaitList.dat'
#Licensed under the GPL3 license

##Relevant modules
import csv
import mmap
import struct
import codecs
import re
from datetime import datetime,timedelta
import os
import time
import sys
import argparse
import binascii

#If -c argument, remove spacing from string for cleaner output
def CleanString(String):
	String = re.sub(r'[\x00]',r'', String)
	return(String)

#Function to retreive the size of 1 index record (based on hex integer)
def GetFileSize(Offset):
	mmFileSize = struct.unpack_from('<i', mm, Offset)
	Offset = Offset + 4
	return(mmFileSize[0], Offset)

#Function to retreive 1 index record from WaitList.dat binary file. Uses FileSize retreived from function 'GetFileSize'
def GetBinary(FileSize, Offset):
	mmWorkingBinary = mm[Offset:(FileSize + (Offset))] ################################removed +1
	#Returns binary of one index record for processing
	return(mmWorkingBinary)

	
#Retreives metadata values which appear at the start of the index record
def GetHeader(WorkingBinary):
	#	* = interprited in report	^ = hex displayed in report for community review
	#						 * * * * ^ * ^ *      
	#						 0 1 2 3 4 5 6 7 
	Header = struct.unpack("<I Q I B B B Q", WorkingBinary[:27])
	RecordOffset = 27
	##Creates a hex string of the unknown header values for output in report for community review
	UnknownHex = [(str(binascii.hexlify(WorkingBinary[17:18])))[2:-1], (str(binascii.hexlify(WorkingBinary[19:27])))[2:-1]]
	# Unknown hex value matching:
	# [0] = Header 4
	# [1] = Header 6
	return(Header, RecordOffset, UnknownHex)

#Function to read metadata values that exist before the body values in a binary index record
def GetData(WorkingBinary, RecordOffset, BeforeBody):
	#Set values as an empty string, as not all records will contain every value. For report printing purposes.
	AllNames = ""
	AllAddresses = ""
	AllOther = ""
	Subject = ""
	Location = ""
	Company = ""
	Address = ""
	City = ""
	Country = ""
	Contact = ""
	FirstName = ""
	FullName = ""
	Title = ""
	Surname = ""
	MiddleName = ""
	State = ""
	ContactURL = ""
	#Note: Following while loop checks the size of the binary index record passed into this function is large enough to contain data.
	while (len(WorkingBinary[RecordOffset:])) > 12 and (struct.unpack_from("<B", WorkingBinary, RecordOffset))[0] == 0: #checks the byte value that instructs whether there is more metadata to extract
		PreData = struct.unpack_from("<I I I", WorkingBinary, RecordOffset+1)
		##PreData Values:
		#PreData[0] = Value flag (e.g. 04 00 00 00 == Address, 07 00 00 00 == Name)
		#PreData[1] = Grammar Proofing Type (not used in report)
		#PreData[2] = Value length in characters. Need to multiple by 2 to get binary length, as waitlist.dat is null terminated.
		RecordOffset += 13 #Offsets current position by 12 (3 integers) + 1 (the byte flag that makes the condition for this while loop)
		Data = struct.unpack_from("%ds" % (PreData[2]*2), WorkingBinary, RecordOffset) #Unpacks binary of the value
		DataAscii = codecs.decode(Data[0], 'ascii', 'ignore') #Decodes binary to ascii
		##Optional: If -c argument was entered, this will remove some non-ascii characters to clean up output
		if args.cleanup:
			DataAscii = CleanString(DataAscii)
		RecordOffset += (PreData[2]*2) #Offsets current position by length of the extracted value
		#Stores the ascii text in the appropriate variable. Can append multiple values to one field (semicolon delimited)
		if PreData[0] == 7:
			AllNames += (DataAscii+";")
		elif PreData[0] == 4:
			AllAddresses += (DataAscii+";")
		elif PreData[0] == 31:
			Company = DataAscii
		elif PreData[0] == 27:
			Location = DataAscii
		elif PreData[0] == 11:
			Address = DataAscii
		elif PreData[0] == 12:
			City = DataAscii
		elif PreData[0] == 13:
			Country = DataAscii
		elif PreData[0] == 14:
			Contact = DataAscii
		elif PreData[0] == 15:
			FirstName = DataAscii
		elif PreData[0] == 16:
			FullName = DataAscii
		elif PreData[0] == 17:
			Title = DataAscii
		elif PreData[0] == 18:
			Surname = DataAscii
		elif PreData[0] == 19:
			MiddleName = DataAscii
		elif PreData[0] == 21:
			State = DataAscii
		elif PreData[0] == 22:
			ContactURL = DataAscii
		elif PreData[0] == 6:
			Subject = DataAscii
		else:
			AllOther += (DataAscii+"[Type:{}]".format(PreData[0])) #Captures un-encountered values in the 'other' field, appending them with the value flag integer value for incorporation into future versions
	RecordOffset += 1 #offsets record past the final byte value which exited the while loop
	#Returns values extracted
	if BeforeBody == 1:
		return(AllNames, AllAddresses, Subject, Location, Company, Address, City, State, Country, FirstName, MiddleName, Surname, FullName, Title, Contact, RecordOffset, ContactURL, AllOther)
	else:
		return(Subject, Other, RecordOffset)
##Function to retreive body text of an index record
def GetBody(WorkingBinary, RecordOffset):
	#Set values as an empty string, as not all records will contain every value. For report printing purposes.
	BodyAscii = ""
	BodyType=""
	#Note: The following while loop checks the size of the binary index record passed into this function is large enough to contain data.
	if len(WorkingBinary[RecordOffset:]) > 16:
		PreBody = struct.unpack_from("<I I I I", WorkingBinary, RecordOffset)
		#PreBody values:
		#PreBody[0] = Current body offset (as body can be broken up into multiple parts which are linked together below)
		#PreBody[1] = Body Type flag (e.g. 05 00 00 00 = Email Body, 17 00 00 00 = Contact, 1d 00 00 00 = Document Body
		#PreBody[2] = Grammar Proofing Type (not used in report)
		#PreBody[3] = Value length in characters. Need to multiple by 2 to get binary length, as waitlist.dat is null terminated.
		##Assigned values to ascii text for report
		BodyType = PreBody[1]
		if BodyType == 5:
			BodyType = 'Email'
		elif BodyType == 23:
			BodyType = 'Contact'
		elif BodyType == 29:
			BodyType = 'Document'
		else:			BodyType = "[Type:{}]".format(BodyType) #Captures un-encountered values in the 'other' field, appending them with the value flag integer value for incorporation into future versions
		RecordOffset += 16 # set record after pre body
		Body = struct.unpack_from('%ds' % (PreBody[3]*2), WorkingBinary, RecordOffset) #Reads binary body length
		BodyAscii = codecs.decode(Body[0], 'ascii', 'ignore') #Decodes binary to ascii text
		RecordOffset += (PreBody[3]*2) #Offsets current position by the length of the binary length extracted
	##The following while loop checks a byte value, which isntructs whether there are additional lengths of body to link to the length just extracted.
	##It also checks to ensure there is enough length left in the index record to have data.
	while (len(WorkingBinary[RecordOffset:])) > 16 and (struct.unpack_from('<B', WorkingBinary, RecordOffset))[0] == 1: 
		PreMoreBody = struct.unpack_from("<I I I I", WorkingBinary, RecordOffset+1)
		#PreMoreBody values:
		#PreMoreBody[0] = Current body offset (as body can be broken up into multiple parts which are linked together below)
		#PreMoreBody[1] = Body Type flag (e.g. 05 00 00 00 = Email Body, 17 00 00 00 = Contact, 1d 00 00 00 = Document Body
		#PreMoreBody[2] = Grammar Proofing Type (not used in report)
		#PreMoreBody[3] = Value length in characters. Need to multiple by 2 to get binary length, as waitlist.dat is null terminated.
		##Assigned values to ascii text for report
		RecordOffset += 17 # set record after PreMoreBody and after byte flag read in the while loop
		MoreBody = struct.unpack_from('%ds' % (PreMoreBody[3]*2), WorkingBinary, RecordOffset) #Reads binary body length
		MoreBodyAscii = codecs.decode(MoreBody[0], 'ascii', 'ignore') #Decodes binary to ascii text
		BodyAscii = (BodyAscii+"\n"+MoreBodyAscii) #Appends this length of body text to previous lengths of body text in this index record
		RecordOffset += (PreMoreBody[3]*2)  #Offsets current position by the length of the binary length extracted
	##Optional: If -c argument was entered, this will remove some non-ascii characters to clean up output
	if args.cleanup:
		BodyAscii = CleanString(BodyAscii)
	RecordOffset += 1 #offsets record past the final byte value which exited the while loop
	return(BodyAscii, RecordOffset, BodyType)

def Print(Offset, Header, UnknownHex, Names, Addresses, Subject, BodyType, Body, Location, Company, Address, City, State, Country, FirstName, MiddleName, Surname, FullName, Title, Contact, ContactURL, Number, Other):
	Offset = Offset - 4
	SentFlag = Header[3]
	if SentFlag == 0:
		SentFlag = 'Sent'
	elif SentFlag == 1:
		SentFlag = ""
	else:
		SentFlag = '[Type:{}]'.format(Header[3])	
	Type = Header[5]
	if Type == 1:
		Type = 'Email'
	elif Type == 0:
		Type = 'Non Email'
	else:
		Type = 'Type[{}]'.format(Header[4])
	RecordLength = (str(Header[0]) + 'b')
	DateTime = str(datetime(1601,1,1) + timedelta(microseconds=(Header[1]/10.)))[:19]
	with open(OutputFile, 'a') as csvfile: #write result to csv output file
		fieldnames = ['Offset', 'Record Length', 'Date/Time(UTC)', 'SentFlag', 'Unkn*', 'Type', 'DocID*', 'Subject', 'Recipient Name(s)', 'Recipient Address(es)', 'BodyType', 'Body', 'Location','Company', 'Address', 'City', 'State', 'Country', 'First Name', 'Middle Name', 'Surname', 'Full Name', 'Title', 'Contact', 'URL', 'Other']
		writer = csv.DictWriter(csvfile, lineterminator='\n', fieldnames=fieldnames)	
		writer.writerow({'Offset': Offset, 'Record Length': RecordLength, 'Date/Time(UTC)':DateTime, 'SentFlag':SentFlag, 'Unkn*':UnknownHex[0], 'Type':Type, 'DocID*': UnknownHex[1], 'Subject': Subject, 'Recipient Name(s)': Names,'Recipient Address(es)': Addresses, 'BodyType': BodyType, 'Body':Number, 'Location':Location, 'Company':Company, 'Address':Address, 'City':City, 'State':State, 'Country':Country, 'First Name':FirstName, 'Middle Name':MiddleName, 'Surname':Surname, 'Full Name':FullName, 'Title':Title, 'Contact':Contact, 'URL': ContactURL, 'Other': Other,})
	WriteTextFile(Offset, RecordLength, UnknownHex[1], DateTime, SentFlag, Type, Subject, Names, Addresses, BodyType, Body, Location, Company, Address, City, State, Country, FirstName, MiddleName, Surname, FullName, Title, Contact, ContactURL, Number, Other)
	WriteToXlsx = (Offset, RecordLength, DateTime, SentFlag, UnknownHex[0], Type, UnknownHex[1], Subject, Names, Addresses, BodyType,  "", Location, Company, Address, City, State, Country, FirstName, MiddleName, Surname, FullName, Title, Contact, ContactURL, Other)
	return(Number, WriteToXlsx)

def WriteTextFile(Offset, RecordLength, DocID, DateTime, SentFlag,  Type, Subject, Names, Addresses, BodyType, Body, Location, Company, Address, City, State, Country, FirstName, MiddleName, Surname, FullName, Title, Contact, ContactURL, Number, Other):
	hashes = '###############################'
	TxtName = "{}{}-WLripReport/TxtFiles/{}.txt".format(OutputDir, Now, Number)
	os.makedirs(os.path.dirname(TxtName), exist_ok=True)	
	TxtFile = open(TxtName, "w")
	TxtFile.write(hashes + '\n' + 'Offset: ' + str(Offset) + '\n' + 'Record Length: ' + RecordLength + '\n' +  'DocID*: ' + DocID + '\n' + 'SentFlag: ' + SentFlag + '\n' + 'Subject: ' + Subject + '\n' + 'Type: ' + Type + '\n' + 'Location: ' + Location + '\n' + 'Company: ' + Company + '\n' + 'Body Type: ' + BodyType + '\n' + hashes + '\n' + 'Date Time: ' + DateTime + '\n' + hashes + '\n' + 'Recipient List: ' + Names + '\n' + 'Address List: ' + Addresses + '\n' + hashes + '\n' + Body + '\n' + hashes + '\n' + '(Only populated if record is a contact)' + '\n' + 'Address :'+ Address + '\n' + 'City :' + City + '\n' + 'State :' + State + '\n' + 'Country :' + Country + '\n' + 'First Name: ' + FirstName + '\n' + 'Middle Name: ' + MiddleName + '\n' + 'Surname: ' + Surname + '\n' + 'Full Name: ' + FullName + '\n' + 'Title: ' + Title + '\n' + 'Contact: ' + Contact + '\n' + 'URL: ' + ContactURL + '\n' + hashes + '\n' + '(Captures data fields not currently encountered by the developer)' + '\n' + 'Other Fields: ' + Other + '\n' + hashes + '\n' + hashes + '\n' + "* DocID: Format Unknown - Appears to be a unique ID for the file that was indexed." + '\n' + "Unkn: Unknown value. Has always been 0 in developer's tests. Output for community analysis." + '\n' + "For more information please visit b2dfir.blogspot.com")
	TxtFile.close()
	return(TxtName)
	
def Close(ItemCount):
	print('**********Extraction Complete**********')
	print('Records Extracted: {}'.format(ItemCount))
	print('***************************************')
	if args.xlsx:
		worksheet.set_column('A:A', 8.11)
		worksheet.set_column('B:B', 11.42)
		worksheet.set_column('C:C', 16.21)
		worksheet.set_column('D:D', 7.86)
		worksheet.set_column('E:E', 5.21)
		worksheet.set_column('F:F', 8.16)
		worksheet.set_column('G:G', 16.21)
		worksheet.set_column('H:H', 40)
		worksheet.set_column('I:I', 25)
		worksheet.set_column('J:J', 25)
		worksheet.set_column('K:K', 8.37)
		worksheet.set_column('L:L', 13.32)
		print("Writing xlsx report...")
		workbook.close()
		print("Done")
	return()
	
#####BODY OF PROGRAM#####

#Parse Arguments
parser = argparse.ArgumentParser(description="WLrip v0.2 - By Barnaby Skeggs - Extract indexed records from: %AppData%\Local\Microsoft\InputPersonalization\TextHarvester\WaitList.dat \n Github:https://github.com/B2dfir/wlrip \n Details:https://b2dfir.blogspot.com.au/2016/10/touch-screen-lexicon-forensics.html")
parser.add_argument("-c", "--cleanup", help="Removes spaces from utf-8 strings for a cleaner text output", action="store_true")
parser.add_argument("-x", "--xlsx", help="Write xlsx report with hyperlinks to record extracts (requires XlsxWriter python module)", action="store_true")
parser.add_argument("-k", "--kill", help="[Admin Required] Kills 'Microsoft Windows Search Indexer' process to remove WaitList.dat file lock on a live system", action="store_true")
parser.add_argument("-f", "--file", help="file input for processing", required=True)
parser.add_argument("-o", "--outputdir", help="output directory")
args = parser.parse_args()
FileInput=args.file

if args.xlsx:
	import xlsxwriter
if args.kill:
	os.system("taskkill /F /im SearchIndexer.exe")

#Program Header
print("// WLrip.py // Parses WaitList.dat // By Barnaby Skeggs // b2dfir.blogspot.com")	

#Time used for folder naming
Now = time.strftime("%Y%m%d-%Hh%Mm%Ss")

##Create Output CSV File And Directory
if args.outputdir:
	OutputDir=args.outputdir + "/"
else:
	OutputDir=""
OutputFile = "{}{}-WLripReport/Wlrip_output.csv".format(OutputDir, Now) #Generates output file/folder name
os.makedirs(os.path.dirname(OutputFile), exist_ok=True)	 #Creates output file/folder
with open(OutputFile, 'w') as csvfile: #write an empty CSV file, ready for results to be appended
	fieldnames = ['Offset', 'Record Length', 'Date/Time(UTC)', 'SentFlag', 'Unkn*', 'Type', 'DocID*', 'Subject', 'Recipient Name(s)', 'Recipient Address(es)', 'BodyType', 'Body', 'Location', 'Company', 'Address', 'City', 'State', 'Country', 'First Name', 'Middle Name', 'Surname', 'Full Name', 'Title', 'Contact', 'URL', 'Other']
	writer = csv.DictWriter(csvfile, lineterminator='\n', fieldnames=fieldnames)
	writer.writeheader()

###Create Output XLSX File
if args.xlsx:
	workbook = xlsxwriter.Workbook("{}{}-WLripReport/Wlrip_output.xlsx".format(OutputDir, Now))
	worksheet = workbook.add_worksheet('output')
	row = 0
	col = 0
	xlsheader = ('Offset', 'Record Length', 'Date/Time(UTC)', 'SentFlag', 'Unkn*', 'Type', 'DocID*', 'Subject', 'Recipient Name(s)', 'Recipient Address(es)', 'BodyType', 'Body', 'Location', 'Company', 'Address', 'City', 'State', 'Country', 'First Name', 'Middle Name', 'Surname', 'Full Name', 'Title', 'Contact', 'URL', 'Other')
	for item in xlsheader:
		worksheet.write(row,col, item)
		col += 1
	row = 1

####Begin Parsing
with open(FileInput, 'rb') as f:
	mm = mmap.mmap(f.fileno(), 0, access=mmap.ACCESS_READ) #reads binary file
	
Offset = 5 #Initial record offset
Number = 1 # Number which increments - used to name output text files
ItemCount = 0 #Number used to increment count of files produced
ProgressCount = 100

while Offset < len(mm):
	FileSize, Offset = GetFileSize(Offset) #Gets offset of the first indexed record, and the file size of the record
	if FileSize >= 50 and (Offset + FileSize) < len(mm):
		WorkingBinary = GetBinary(FileSize, Offset) #Assigns the first indexed record (based on offset and file size) to a variable for processing
		Header, RecordOffset, UnknownHex = GetHeader(WorkingBinary) #Reads the first set of values from the indexed record, and increments the record offset for remaining processing
		Names, Addresses, Subject, Location, Company, Address, City, State, Country, FirstName, MiddleName, Surname, FullName, Title, Contact, RecordOffset, ContactURL, Other = GetData(WorkingBinary, RecordOffset, 1) #Retreives data which is structured in the same format from the indexed record
		Body, RecordOffset, BodyType = GetBody(WorkingBinary, RecordOffset) #Retreives Body of file
		if Header[5] == 1: #If Header[5] == 1, then the file is an email and will have a subject. Subject is stored after the body text, so the function is required to be called again
			if Subject == "":
				RecordOffset = RecordOffset - 1
				Subject, Other1, RecordOffset = GetData(WorkingBinary, RecordOffset, 0) #retreives subject of email, and any other additional metadata values after the body text (none have been identified, this primarily for error handling)
				if Other1 != "": 
					Other += Other1#Adds any new metadata fields identified after body text to original 'other' variable. 
		else:
			Subject = "" #sets subject field to blank for non emails
		Number, WriteToXlsx = Print(Offset, Header, UnknownHex, Names, Addresses, Subject, BodyType, Body, Location, Company, Address, City, State, Country, FirstName, MiddleName, Surname, FullName, Title, Contact, ContactURL, Number, Other)
		###Write to XLSX###
		if args.xlsx:	
			col = 0
			for item in (WriteToXlsx):
				worksheet.write(row, col, item)
				col += 1
			row += 1
			worksheet.write_url('L{}'.format(row), "TxtFiles/{}.txt".format(Number))
		###################
		Number += 1
		RecordOffset = 0
		Offset = Offset + FileSize + 1
		ItemCount += 1
	else:
		Offset = Offset + FileSize + 1
	if ItemCount == ProgressCount:
		print(str(ItemCount) + " records processed...")
		ProgressCount += 100
Close(ItemCount)
