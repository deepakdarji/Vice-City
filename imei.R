library(openxlsx) 
library(chron) 
library(data.table)

# file = "Airtel.xlsx"
# SUS_NO = "7543812518"
# imp_row = 6
# raw_sheet = 1

infile = readline(prompt="Input file: ")
SUS_NO = readline(prompt="Suspect number: ")
imp_row = as.numeric(readline(prompt="Row no. containing headers: "))
raw_sheet = as.numeric(readline(prompt="Raw CDR Sheet number: "))


df = read.xlsx(xlsxFile=infile, sheet=raw_sheet, startRow=imp_row, colNames=TRUE, detectDates=TRUE, skipEmptyRows=FALSE)

df <- df[!is.na(df$`Dur(s)`),] # remove useless rows from bottom

names(df)[names(df) == 'Calling.No'] <- 'Calling_No' 

colnames(df) <- c("Calling_No", "Called_No", "Date", "Time", "Duration", "Cell_1", "Cell_2", "Communication_Type", "IMEI", "IMSI", "Type", "SMSC", "Roam")

drops <- c("Calling_No", "Called_No","Duration", "Cell_1", "Cell_2", "Communication_Type","IMSI", "Type", "SMSC", "Roam") # These are the columns you want that may be removed, change them as per your will
df = df[ , !(names(df) %in% drops)] # drop the columns from the dataframe

df$Time <- times(as.numeric(df$Time)) # Change time from default decimal format to proper HH::MM::SS format

# Microsoft Excel date gives some offset, remove it using this and change format to proper DD/MM/YY format
df$Date <- format(as.Date(as.numeric(df$Date) ,origin = "1899-12-30"), "%d/%m/%y")


dt = as.data.table(df$IMEI) # abstracting out the IMEI column
dN<-dt[, .N ,by = df$IMEI] # creating frequency table
dS<- dN[order(-N)]  # arranging in descending order of the frequencies of IMEI used and storing it
 
#Storing the output in a file.
outfile = readline(prompt="Output file: ")
sheet = readline(prompt="Sheet name: ")
hs <- createStyle(textDecoration = "BOLD", fontColour = "#FFFFFF", fontSize=12, fontName="Arial Narrow", fgFill = "#4F80BD")
write.xlsx(x=dS, file=outfile, sheetName = sheet, headerStyle=hs)


