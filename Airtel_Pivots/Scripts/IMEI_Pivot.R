infile = "CDR.xlsx"
SUS_NO = "7543812518"
imp_row = 1
raw_sheet = 1

# infile = readline(prompt="Input file: ")
# SUS_NO = readline(prompt="Suspect number: ")
# imp_row = as.numeric(readline(prompt="Row no. containing headers: "))
# raw_sheet = as.numeric(readline(prompt="Raw CDR Sheet number: "))

# Read particular sheet from excel workbook; adjust startRow as per the number of useless rows in the beginnig, remember to remove empty rows
D = read.xlsx(xlsxFile=infile, sheet=raw_sheet, startRow=imp_row, colNames=TRUE, detectDates=TRUE, skipEmptyRows=FALSE)

# Create IMEI Pivot

IMEI_pivot <- as.data.frame(table(D$IMEI))
colnames(IMEI_pivot) <- c("IMEI", "Frequency")
IMEI_pivot <- IMEI_pivot[order(-IMEI_pivot$Frequency),]

wb <- createWorkbook()
addWorksheet(wb, "IMEI")
writeData(wb,1,IMEI_pivot)
saveWorkbook(wb, file = "IMEI_Pivot.xlsx", overwrite = TRUE)