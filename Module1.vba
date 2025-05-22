Option Explicit

Function SumConsecutiveCells(rng As Range) As Double
    Dim cell As Range
    Dim sum As Double
    
    sum = 0
    For Each cell In rng
        If IsEmpty(cell.Value) Then
            Exit For
        Else
            sum = sum + cell.Value
        End If
    Next cell
    
    SumConsecutiveCells = sum
End Function

Sub CalculateMonthlyExpenditures()
    Dim ws As Worksheet
    Dim dataRange As Range
    Dim cell As Range
    Dim startDate As Date
    Dim currentExpenditure As Double
    Dim monthIndex As Long
    Dim months() As String
    Dim expenditures() As Double
    Dim data() As Variant
    Dim i As Long
    Dim chartObj As ChartObject

    ' Set the worksheet and data range
    Set ws = ThisWorkbook.Sheets("Sheet1")
    Set dataRange = ws.Range("C9:C1000")

    ' Initialize variables
    startDate = DateSerial(2024, 10, 1)
    currentExpenditure = 0
    monthIndex = 0

    ' Loop through the data range
    For Each cell In dataRange
        If IsEmpty(cell.Value) Then
            If currentExpenditure > 0 Then
                ReDim Preserve months(monthIndex)
                ReDim Preserve expenditures(monthIndex)
                months(monthIndex) = Format(startDate, "mmmm yyyy")
                expenditures(monthIndex) = currentExpenditure
                startDate = DateAdd("m", 1, startDate)
                currentExpenditure = 0
                monthIndex = monthIndex + 1
            End If
        Else
            currentExpenditure = currentExpenditure + cell.Value
        End If
    Next cell

    ' Add the last month's expenditure if there's any remaining
    If currentExpenditure <> 0 Then
        ReDim Preserve months(monthIndex)
        ReDim Preserve expenditures(monthIndex)
        months(monthIndex) = Format(startDate, "mmmm yyyy")
        expenditures(monthIndex) = currentExpenditure
    End If

    ' Prepare data with months and expenditures
    ReDim data(1 To monthIndex + 1, 1 To 2)
    data(1, 1) = "Month"
    data(1, 2) = "Expenditure (CAD)"
    For i = 1 To monthIndex
        data(i + 1, 1) = months(i - 1)
        data(i + 1, 2) = expenditures(i - 1)
    Next i

    ' Write the updated data starting at K16
    ws.Range("K16:L" & 16 + UBound(data, 1) - 1).Value = data

    ' Delete any existing chart at N16
    For Each chartObj In ws.ChartObjects
        If chartObj.TopLeftCell.Address = ws.Range("N16").Address Then
            chartObj.Delete
        End If
    Next chartObj

    ' Create the bar chart at N16
    Dim chart As chart
    Set chart = ws.ChartObjects.Add(Left:=ws.Range("N16").Left, Top:=ws.Range("N16").Top, Width:=375, Height:=225).chart
    chart.ChartType = xlColumnClustered
    chart.SetSourceData Source:=ws.Range("K16:L" & 16 + UBound(data, 1) - 1)
    chart.ChartTitle.Text = "Electric Car Expenditures per Month (CAD)"
    chart.Axes(xlCategory, xlPrimary).HasTitle = True
    chart.Axes(xlCategory, xlPrimary).AxisTitle.Text = "Month"
    chart.Axes(xlValue, xlPrimary).HasTitle = True
    chart.Axes(xlValue, xlPrimary).AxisTitle.Text = "Expenditure (CAD)"
End Sub
Sub AppendData()
    Dim ws As Worksheet
    Dim lastRow As Long
    
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Change "Sheet1" to your actual sheet name
    
    ' Find the last filled row in column A
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    ' Copy values from K6:O6 to respective columns
    ws.Cells(lastRow, 1).Value = ws.Range("K6").Value  ' Date in Column A
    ws.Cells(lastRow, 3).Value = ws.Range("L6").Value  ' Cost in Column C
    ws.Cells(lastRow, 4).Value = ws.Range("M6").Value  ' Km in Column D
    ws.Cells(lastRow, 5).Value = ws.Range("N6").Value  ' kWh in Column E
    ws.Cells(lastRow, 7).Value = ws.Range("O6").Value  ' Where in Column G
    
    ' Add formula in column B based on column D and cell R2
    ws.Cells(lastRow, 2).Formula = "=D" & lastRow & "/$R$2"
    
    ' Add formula in column F to sum previous row's F value and current row's C value
    If lastRow > 2 Then
        ws.Cells(lastRow, 6).Formula = "=F" & lastRow - 1 & "+C" & lastRow
    Else
        ws.Cells(lastRow, 6).Formula = "=C" & lastRow ' First row just takes column C value
    End If
    
    MsgBox "Data has been appended successfully!", vbInformation, "Success"
End Sub