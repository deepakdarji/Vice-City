infile = "CDR.xlsx"
SUS_NO = "7543812518"
imp_row = 1
raw_sheet = 1

D = read.xlsx(xlsxFile=infile, sheet=raw_sheet, startRow=imp_row, colNames=TRUE, detectDates=TRUE, skipEmptyRows=FALSE)

# Create a new dataframe for pivoting according to B_party

# This function has two parameters -> x: phone numbers; t: communication type, supplied when calling this function
f <- function(x, t)
{
	return (sum( D$Communication_Type[D$Calling_No == x | D$Called_No == x] == t ))
}

J = rbind(as.matrix(D$Calling_No), as.matrix(D$Called_No)) # all numbers that communicated with our suspect
U = as.matrix(unique(J[J != SUS_NO])) # the same in matrix format, to be used in apply function
rm(J) # delete temporary J

# Gets count of each communication type; MARGIN=1, function is applied for each row of matrix
# x1 means first column of X=U
SMS_IN_List = apply(X=U, MARGIN=1, FUN=function(x1) f(x = x1, t="SMS_IN"))
SMS_OUT_List = apply(X=U, MARGIN=1, FUN=function(x1) f(x = x1, t="SMT"))
CALL_IN_List = apply(X=U, MARGIN=1, FUN=function(x1) f(x = x1, t="IN"))
CALL_OUT_List = apply(X=U, MARGIN=1, FUN=function(x1) f(x = x1, t="OUT"))
# Total communication count between all B-parties
TOTAL_List = SMS_IN_List+SMS_OUT_List+CALL_IN_List+CALL_OUT_List

# Create pivot dataframe: B-Party and communication count details
B_pivot <- data.frame(U,CALL_IN_List,CALL_OUT_List,SMS_IN_List,SMS_OUT_List,TOTAL_List)
colnames(B_pivot) <- c("B_Party", "IN", "OUT", "SMS_IN", "SMS_OUT", "Total") # give column names
B_pivot <- B_pivot[order(-B_pivot$Total),] # Note the (-) sign: sort  in descending order of total communication

wb <- createWorkbook()
addWorksheet(wb, "B_Party")
writeData(wb,1,B_pivot)
saveWorkbook(wb, file = "B_Pivot.xlsx", overwrite = TRUE)