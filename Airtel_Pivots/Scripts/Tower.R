infile = "Airtel.xlsx"
imp_row = 1
raw_sheet = 5

# infile = readline(prompt="Input file: ")
# imp_row = as.numeric(readline(prompt="Row no. containing headers: "))
# raw_sheet = as.numeric(readline(prompt="Raw CDR Sheet number: "))

# Read particular sheet from excel workbook; adjust startRow as per the number of useless rows in the beginnig, remember to remove empty rows
df = read.xlsx(xlsxFile=infile, sheet=raw_sheet, startRow=imp_row, colNames=TRUE, detectDates=TRUE, skipEmptyRows=FALSE)

#Do this in case you want to change the names of all columns (more preferable)
# df <- mutate_all(df, funs(toupper))
colnames(df) <- c("A_Party", "B_Party", "Date", "Time", "Duration", "Communication_Type", "IMEI", "Cell_1", "Location", "Town", "Latitude", "Longitude", "Azimuth", "Roam")
drops <- c("A_Party","B_Party", "Date", "Time", "Duration", "Communication_Type", "IMEI", "Roam")
df = df[ , !(names(df) %in% drops)] # drop the columns from the dataframe

# If Town is NA, replace it with Unknown_Town
df$Town[is.na(df$Town)] = "Unknown_Town"
df[is.na(df)] = ""

# Get unique Tower Cell1 IDs and dump tower details to file
tower = df[!duplicated(df$Cell_1),]
wb_temp <- createWorkbook()
addWorksheet(wb_temp, "Tower_Info")
writeData(wb_temp,1,tower)
saveWorkbook(wb_temp, file = "Tower_Info.xlsx", overwrite = TRUE)