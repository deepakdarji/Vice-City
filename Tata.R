library(openxlsx) # This library is used to open xlsx files directly in R
library(chron) # This library is used for time manipulation
library(data.table) # Used to join rows one after another

# file = "Tata.xlsx"
# SUS_NO = "9058786086"
# imp_row = 4
# raw_sheet = 1

infile = readline(prompt="Input file: ")
SUS_NO = readline(prompt="Suspect number: ")
imp_row = as.numeric(readline(prompt="Row no. containing headers: "))
raw_sheet = as.numeric(readline(prompt="Raw CDR Sheet number: "))

# Read particular sheet from excel workbook; adjust startRow as per the number of useless rows in the beginnig, remember to remove empty rows
df = read.xlsx(xlsxFile=infile, sheet=raw_sheet, startRow=imp_row, colNames=TRUE, detectDates=TRUE, skipEmptyRows=FALSE)

df <- df[!is.na(df$DURATION),] # remove useless rows from bottom

names(df)[names(df) == 'Calling.No'] <- 'Calling_No' # Do this in case you want to change the header name of a particular column

#Do this in case you want to change the names of all columns (more preferable)
colnames(df) <- c("Calling_No", "Called_No", "Date", "Time", "Duration", "Cell_1", "Cell_2", "Communication_Type",
					"IMEI", "IMSI", "Type", "SMSC", "Roam", "Switch", "LRN")

drops <- c("Type","SMSC","Switch","LRN") # These are the columns you want that may be removed, change them as per your will
df = df[ , !(names(df) %in% drops)] # drop the columns from the dataframe

# Make hyphens empty in Cell_1 and Cell_2 columns
df$Cell_1[df$Cell_1 == "-"] <- ""
df$Cell_2[df$Cell_2 == "-"] <- ""
df$Roam[df$Roam == "-"] <- ""

df$Time <- times(as.numeric(df$Time)) # Change time from default decimal format to proper HH::MM::SS format

# Microsoft Excel date gives some offset, remove it using this and change format to proper DD/MM/YY format
df$Date <- format(as.Date(df$Date), "%d/%m/%y")

########################################################################################################################################################

# Create a new dataframe for pivoting according to B_party

# This function has two parameters -> x: phone numbers; t: communication type, supplied when calling this function
f <- function(x, t)
{
	return (sum( df$Communication_Type[df$Calling_No == x | df$Called_No == x] == t ))
}

J = rbind(as.matrix(df$Calling_No), as.matrix(df$Called_No)) # all numbers that communicated with our suspect
U = as.matrix(unique(J[J != SUS_NO])) # the same in matrix format, to be used in apply function
rm(J) # delete temporary J

# Gets count of each communication type; MARGIN=1, function is applied for each row of matrix
# x1 means first column of X=U
SMS_IN_List = apply(X=U, MARGIN=1, FUN=function(x1) f(x = x1, t="SMT"))
SMS_OUT_List = apply(X=U, MARGIN=1, FUN=function(x1) f(x = x1, t="SMO"))
CALL_IN_List = apply(X=U, MARGIN=1, FUN=function(x1) f(x = x1, t="MTC"))
CALL_OUT_List = apply(X=U, MARGIN=1, FUN=function(x1) f(x = x1, t="MOC"))
# Total communication count between all B-parties
TOTAL_List = SMS_IN_List+SMS_OUT_List+CALL_IN_List+CALL_OUT_List

# Create pivot dataframe: B-Party and communication count details
df_pivot <- data.frame(U,CALL_IN_List,CALL_OUT_List,SMS_IN_List,SMS_OUT_List,TOTAL_List)
colnames(df_pivot) <- c("B_Party", "IN", "OUT", "SMS_IN", "SMS_OUT", "Total") # give column names
df_pivot <- df_pivot[order(-df_pivot$Total),] # Note the (-) sign: sort  in descending order of total communication

# Output pivot details to a file
outfile = readline(prompt="Output file: ")
sheet = readline(prompt="Sheet name: ")
hs <- createStyle(textDecoration = "BOLD", fontColour = "#FFFFFF", fontSize=12, fontName="Arial Narrow", fgFill = "#4F80BD")
write.xlsx(x=df_pivot, file=outfile, sheetName = sheet, headerStyle=hs)
