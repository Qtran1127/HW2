sub stockvalue ()
dim Ticker as String
dim Total_Stock as Double
dim ws as worksheet
for each ws in worksheets
LastRow = ws.Cells (Rows.Count,1).End(xlUp).Row
dim i as long
Total_Stock = 0
dim Summary_Table_Row as Integer
Summary_Table_Row = 2
for i = 2 to LastRow
if Celss(i+1,1).Value <> Cells (i,1).Value then
Ticker = Celss(i,1).Value
Total_Stock = Total_Stock + Cells (i,7).Value
Range ("I" & Summary_Table_Row).Value = Ticker
Range ("L" & Summary_Table_Row).Value = Total_Stock
Summary_Table_Row = Summary_Table_Row + 1
Total_Stock = 0
else
Total_Stock = Total_Stock + Cells(i,7).Value
end if 
next i
next ws
end sub