Sub CreateSchedule()
    Dim ws As Worksheet
    Dim startDate As Date
    Dim endDate As Date
    Dim currentDate As Date
    Dim nextRow As Long
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "Schedule"
    
    ' Set the start and end dates
    startDate = DateSerial(2024, 1, 1) ' January 1, 2024
    endDate = DateSerial(2024, 3, 31) ' March 31, 2024
    
    ' Initialize the current date
    currentDate = startDate
    
    ' Add headers to the sheet
    ws.Range("A1").Value = "Date"
    ws.Range("B1").Value = "Day"
    
    ' Start populating the schedule
    nextRow = 2 ' Start from row 2
    
    Do While currentDate <= endDate
        ' Check if the current day is Monday, Wednesday, or Friday
        If Weekday(currentDate, vbMonday) = 1 Or Weekday(currentDate, vbMonday) = 3 Or Weekday(currentDate, vbMonday) = 5 Then
            ' Add date to column A
            ws.Cells(nextRow, 1).Value = currentDate
            
            ' Add day to column B
            ws.Cells(nextRow, 2).Value = Format(currentDate, "dddd")
            
            ' Move to the next row
            nextRow = nextRow + 1
        End If
        
        ' Move to the next day
        currentDate = currentDate + 1
    Loop
End Sub
