# Install the libraries using:
# install.packages("<package name>")

# install.package("openxlsx")
# install.package("xlsx")
# install.package("chron")
# install.package("data.table")

# rm(list=ls())

library(openxlsx) # This library is used to open xlsx files directly in R
library(chron) # This library is used for time manipulation
library(data.table) # Used to join rows one after another

# infile = "Airtel.xlsx"
# SUS_NO = "7543812518"
# imp_row = 6
# raw_sheet = 1

infile = readline(prompt="Input file: ")
SUS_NO = readline(prompt="Suspect number: ")
imp_row = as.numeric(readline(prompt="Row no. containing headers: "))
raw_sheet = as.numeric(readline(prompt="Raw CDR Sheet number: "))

# Read particular sheet from excel workbook; adjust startRow as per the number of useless rows in the beginnig, remember to remove empty rows
df = read.xlsx(xlsxFile=infile, sheet=raw_sheet, startRow=imp_row, colNames=TRUE, detectDates=TRUE, skipEmptyRows=FALSE)

########################################################################################################################################################

# Pre-processing

df <- df[!is.na(df$`Dur(s)`),] # remove useless rows from bottom

names(df)[names(df) == 'Calling.No'] <- 'Calling_No' # Do this in case you want to change the header name of a particular column

#Do this in case you want to change the names of all columns (more preferable)
colnames(df) <- c("Calling_No", "Called_No", "Date", "Time", "Duration", "Cell_1", "Cell_2", "Communication_Type", "IMEI", "IMSI", "Type", "SMSC", "Roam")

drops <- c("Type","SMSC") # These are the columns you want that may be removed, change them as per your will
df = df[ , !(names(df) %in% drops)] # drop the columns from the dataframe

# Make hyphens empty in Cell_1 and Cell_2 columns
df$Cell_1[df$Cell_1 == "-"] <- ""
df$Cell_2[df$Cell_2 == "-"] <- ""

df$Time <- times(as.numeric(df$Time)) # Change time from default decimal format to proper HH::MM::SS format

# Microsoft Excel date gives some offset, remove it using this and change format to proper DD/MM/YY format
df$Date <- format(as.Date(as.numeric(df$Date) ,origin = "1899-12-30"), "%d/%m/%y")

wb <- createWorkbook()
addWorksheet(wb, "CDR")
writeData(wb,1,df)
saveWorkbook(wb, file = "CDR.xlsx", overwrite = TRUE)

########################################################################################################################################################

