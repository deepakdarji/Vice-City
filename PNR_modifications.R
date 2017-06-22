library(openxlsx) # For opening xlsl files in R
library(chron) # For time varaitions
library(data.table) #For Databse changes
library(stringr) # Mainly used for separting strings
library(sqldf) # sqlite library

input_file = readline(prompt = "File Input : ") # Name of the File in the working directory
SUS_NO = readline(prompt = "Suspects Number : ") # Suspected number
Head_row = as.numeric(readline(prompt = "Row containing Header : ")) # Where header starts
raw_sheet = as.numeric(readline(prompt = "Raw PNR sheet no : ")) # Raw sheet no. in the file

#Importing the required file in R

PNR <- read.xlsx(xlsxFile = input_file , sheet = raw_sheet , startRow = Head_row , colNames=TRUE, detectDates=TRUE, skipEmptyRows=FALSE)

# Upadating the column names as per our requirement

colnames(PNR) <- c("PNR_NO","BOOKING_DATE", "JOURNEY_DATE","QUOTA", "USER_ID", "EMAIL", "REGISTERED_NAME", "BOOKING_MOBILE", "PROFILE_MOBILE","OFFICE_PHONE", "ADDRESS","DISTRICT" , "STATE","PINCODE","TRANSACTION_ID", "BANK", "BANK_TRANSACTION_NO", "AMOUNT", "FROM" , "TO" ,"CLASS" , "TRAIN_NO" , "IP_ADDRESS", "PNR_DETAILS")

# Drop the irrelevant fields

drop <- c("PRS AMOUNT","QUOTA")
PNR = PNR[,!(names(PNR) %in% drop)]

############################################################
# This piece of code Take PNR_DETAILS from our main table takes it to a new table and separates the names of the passenger

PASS = as.matrix(PNR$PNR_DETAILS)

PASS = str_split_fixed(PASS, "#",Inf)

PNR = data.frame(PNR , PASS) # Join both data-tables (names and original one)

# Drop PNR_ details as it has no use

dropp <- c("PNR_DETAILS")
PNR = PNR[,!(names(PNR) %in% dropp)]

# Updating the final list of the Column names

colnames(PNR) <- c("PNR_NO","BOOKING_DATE", "JOURNEY_DATE","QUOTA", "USER_ID", "EMAIL", "REGISTERED_NAME", "BOOKING_MOBILE", "PROFILE_MOBILE","OFFICE_PHONE", "ADDRESS","DISTRICT" , "STATE","PINCODE","TRANSACTION_ID", "BANK", "BANK_TRANSACTION_NO", "AMOUNT", "FROM" , "TO" ,"CLASS" , "TRAIN_NO" , "IP_ADDRESS", "PASSENGER1","PASSENGER2","PASSENGER3","PASSENGER4","PASSENGER5","PASSENGER6")

# Output pivot details to a file
outfile = readline(prompt="Output file: ")
sheet = readline(prompt="Sheet name: ")
hs <- createStyle(textDecoration = "BOLD", fontColour = "#FFFFFF", fontSize=12, fontName="Arial Narrow", fgFill = "#4F80BD")
write.xlsx(x=PNR, file=outfile, sheetName = sheet, headerStyle=hs)
