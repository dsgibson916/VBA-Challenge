Sub Stock():
'declare Variables
Dim ticker As Integer
Dim Year_open As Double
Dim Year_close As Double
Dim Yearly_Change As Double
Dim Percent As Double
Dim Stock_Volume As Double
Dim ws As Worksheet
 

For Each ws In Worksheets


Starter = 2
lastrow = Cells(Rows.Count, "A").End(xlUp).Row

'assign for loop
    For I = 2 To lastrow

'search the cells for different values
If (ws.Cells(I, 1).Value <> ws.Cells(I + 1, 1)) Then
   ws.Cells(Starter, 10).Value = ws.Cells(I, 1).Value
 'get data for Yearly Change
    

 Year_close = ws.Cells(I, 6).Value
 Yearly_Change = Year_close - Year_open
 ws.Cells(Starter, 11).Value = Yearly_Change
 'MsgBox (Year_close & " " & Year_open)
        
        'Get data for percent change
        Percent = (Yearly_Change / Year_open) * 100
        ws.Cells(Starter, 12).Value = Percent
        'Define Stock value
Stock_Volume = Stock_Volume + ws.Cells(I, 7).Value
ws.Cells(Starter, 13).Value = Stock_Volume
   Starter = Starter + 1
   
ElseIf (ws.Cells(I, 1).Value <> ws.Cells(I - 1, 1)) Then
Year_open = ws.Cells(I, 3).Value

Stock_Volume = 0
Stock_Volume = Stock_Volume + Cells(I, 7).Value
'MsgBox (Year_open)

ElseIf ws.Cells(I, 1).Value = ws.Cells(I + 1, 1) Then

Stock_Volume = Stock_Volume + ws.Cells(I, 7).Value

End If

If ws.Cells(I, 12).Value > 0 Then
ws.Cells(I, 12).Interior.ColorIndex = 4
Else: ws.Cells(I, 12).Interior.ColorIndex = 3

End If



Next I

Next ws
End Sub 