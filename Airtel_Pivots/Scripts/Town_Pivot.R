infile = "CDR.xlsx"
SUS_NO = "7543812518"
imp_row = 1
raw_sheet = 1

D = read.xlsx(xlsxFile=infile, sheet=raw_sheet, startRow=imp_row, colNames=TRUE, detectDates=TRUE, skipEmptyRows=FALSE)

names(D)[names(D) == 'Calling_No'] <- 'A_Party'
names(D)[names(D) == 'Called_No'] <- 'B_Party'
D[D$B_Party == SUS_NO, c("A_Party","B_Party")] <- D[D$B_Party == SUS_NO, c("B_Party","A_Party")] # Swap columns so that this B_Party has true B_parties

# Read tower database
tower = read.xlsx(xlsxFile="Tower_Info.xlsx", sheet=1, startRow=1, colNames=TRUE, detectDates=TRUE, skipEmptyRows=FALSE)

#################################################################################################################################################

# Get tower-ID for each record and make a consolidated new dataframe with all details

f <- function(x)
{
	return (tower[tower$Cell_1 == x, ])
}

W = do.call(rbind, lapply(D$Cell_1, f))
rownames(W) = NULL # Reset row names
W = W[, !(colnames(W) %in% c("Cell_1"))] #Drop duplicate column

# Merge two dataframes in one: This is x
x = merge(D, W, by=0)
x = x[order(as.numeric(x$Row.names)),]
x = x[, !(colnames(x) %in% c("Row.names"))]
rownames(x) <- NULL

# Save merged dataframe to workbook
wb <- createWorkbook()
addWorksheet(wb, "Location_Wise")
writeData(wb,1,x)
saveWorkbook(wb, file = "Location_Wise.xlsx", overwrite = TRUE)

##################################################################################################################################################

# Pivot by Town and save to file

town_pivot <- as.data.frame(table(x$Town)) # Note the 'x'
town_pivot <- town_pivot[order(-town_pivot$Freq),] # Sort by descending order
colnames(town_pivot) <- c("Town", "Frequency")

# Save this to workbook
wb <- createWorkbook()
addWorksheet(wb, "Town")
writeData(wb,1,town_pivot)
saveWorkbook(wb, file = "Town_Pivot.xlsx", overwrite = TRUE)