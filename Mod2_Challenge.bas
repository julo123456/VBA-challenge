Attribute VB_Name = "Module6"
'https://support.microsoft.com/en-us/topic/macro-to-loop-through-all-worksheets-in-a-workbook-feef14e3-97cf-00e2-538b-5da40186e2b0
'Help from AskingLearningAssistance for the code start+1 to lookup the open price.
'Help from AskingLearningAssistance for the ws.code in each called out Cells_Range_Columns

Sub Run_Mod2_For_Each_Sheet()

 ' Declare Current as a worksheet object variable.
         Dim ws As Worksheet

         ' Loop through all of the worksheets in the active workbook.
         For Each ws In Worksheets
     

            ' Insert your code here.
            

'Give columne titles and adjust column widths
ws.Range("$I$1").Value = "Ticker"
ws.Range("$J$1").Value = "Total Stock Volumn"
ws.Range("$K$1").Value = " Yearly Change"
ws.Range("$L$1").Value = "Percent Change"
ws.Range("A:M").ColumnWidth = 20
ws.Range("A:M").HorizontalAlignment = xlCenter

Dim ticker As String


Dim table_row As Long
  table_row = 2

Dim total_stock As Double
total_stock = 0
  
Dim value_open As Double
    value_open = 0
    
Dim value_close As Double
    value_close = 0
    

Dim start As Long
    start = 2
    


Dim LR As Long
LR = ws.Cells(Rows.Count, 1).End(xlUp).Row

'start loop
For I = 2 To LR
If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
ticker = ws.Cells(I, 1).Value



'calculate total stock volumn
total_stock = ws.Cells(I, 7).Value + total_stock

'prepare to calculate yearly change
value_open = ws.Cells(start, 3).Value
value_close = ws.Cells(I, 6).Value
yearly_change = value_close - value_open
Percent_Change = FormatPercent(yearly_change / value_open)


start = I + 1

'give value to desired summary
ws.Range("I" & table_row).Value = ticker
ws.Range("J" & table_row).Value = total_stock
ws.Range("K" & table_row).Value = yearly_change
ws.Range("L" & table_row).Value = Percent_Change


'assign color to yearly_change
If yearly_change < 0 Then
ws.Range("K" & table_row).Interior.color = vbRed

Else
ws.Range("K" & table_row).Interior.color = vbGreen
End If


table_row = table_row + 1
total_stock = 0

Else

total_stock = total_stock + ws.Cells(I, 7).Value



End If

Next I


ws.Range("$N$2").Value = "Greatest % Increase"
ws.Range("$N$3").Value = "Greatest % Decrease"
ws.Range("$N$4").Value = "Greatest Total Volumn"
ws.Range("$O$1").Value = "Ticker"
ws.Range("$P$1").Value = "Value"
ws.Range("N:P").ColumnWidth = 20
ws.Range("N:P").HorizontalAlignment = xlCenter

'calculate greatest % increase, greatest decrease, and max total volumn
Max = WorksheetFunction.Max(ws.Columns("L"))
Min = WorksheetFunction.Min(ws.Columns("L"))
MaxTotal = WorksheetFunction.Max(ws.Columns("J"))
ws.Range("$P$2").Value = FormatPercent(Max)
ws.Range("$P$3").Value = FormatPercent(Min)
ws.Range("$P$4").Value = MaxTotal

'vlookup to get ticker information for Max_Increase, Min_Increase and Max_Total_Volumn
ws.Range("$O$2") = Application.WorksheetFunction.XLookup(Max, ws.Range("L:L"), ws.Range("I:I"))
ws.Range("$O$3") = Application.WorksheetFunction.XLookup(Min, ws.Range("L:L"), ws.Range("I:I"))
ws.Range("$O$4") = Application.WorksheetFunction.XLookup(MaxTotal, ws.Range("J:J"), ws.Range("I:I"))


            ' This line displays the worksheet name in a message box.
            MsgBox ws.Name
         
         Next ws


End Sub

