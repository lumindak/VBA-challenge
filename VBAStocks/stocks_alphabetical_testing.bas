Attribute VB_Name = "Module1"


Sub stock_data() 'begin stock_data

Dim Tickers() As String 'array for tickers
Dim n_rows As Long
Dim j As Long
Dim temp1 As Long
Dim yearlychange As Double
Dim begin_value As Double
Dim end_value As Double
Dim tot_volume As Double
Dim maxincrease As String
Dim maxdecrease As String
Dim maxtotalvolume As String
Dim split_maxincrease
Dim split_maxdecrease
Dim split_maxtotalvolume
Dim sheets(7) As String

'start looping the sheets
'sheets = Array("A", "B", "C", "D", "E", "F", "P")
sheets(0) = "A"
sheets(1) = "B"
sheets(2) = "C"
sheets(3) = "D"
sheets(4) = "E"
sheets(5) = "F"
sheets(6) = "P"

For sheet = 0 To 6
    Dim s As String
    s = sheets(sheet)
    
    n_rows = Findrows(2, 1, s) ' finds the no.of rows
    Worksheets(s).Activate
    Worksheets(s).Columns(10).ClearContents  'clear column 10
    Worksheets(s).Columns(11).ClearContents
    Worksheets(s).Columns(11).Interior.Color = xlNone
    Worksheets(s).Columns(12).ClearContents
    Worksheets(s).Columns(13).ClearContents
    Worksheets(s).Columns(14).ClearContents

    'initialiuze values

    j = 2
    temp1 = 0
    tot_volume = 0

    'headers
    Cells(1, 10).Value = Cells(1, 1).Value
    Cells(1, 11).Value = "Yearly Change"
    Cells(1, 12).Value = "percent. Change"
    Cells(1, 13).Value = "total volume"


    For i = 2 To n_rows + 1 'begin looping the rows

        Dim temp As String

        If (Cells(i, 1).Value <> Cells(i - 1, 1).Value) Then
        begin_value = Cells(i, 3).Value
        tot_volume = Cells(i, 7).Value   'initial volume
        End If

        If (Cells(i, 1).Value <> Cells(i + 1, 1).Value) Then
            end_value = Cells(i, 6).Value
        End If

        temp = Cells(i, 1).Value


        If (Cells(i + 1, 1).Value = temp) Then
            tot_volume = tot_volume + Cells(i, 7).Value
    
        Else
    
            Cells(j, 10).Value = temp
            yearlychange = end_value - begin_value
            Cells(j, 11).Value = yearlychange
            
            If (begin_value <> 0) Then
                Cells(j, 12).Value = Str(yearlychange * 100 / begin_value) + "%"
            Else
                Cells(j, 12).Value = "0"
            End If
            
            tot_volume = tot_volume + Cells(i, 7).Value
            Cells(j, 13).Value = tot_volume 'total volume
   
            If (yearlychange < 0) Then
                Cells(j, 11).Interior.ColorIndex = 3 'red  color
            End If
    
            If (yearlychange > 0) Then
                Cells(j, 11).Interior.ColorIndex = 4 'green color
            End If
   
            j = j + 1
    
        End If

    Next 'end of looping the rows

    'MsgBox (j)

    maxincrease = max_increase(2, 12, s)
    maxdecrease = max_decrease(2, 12, s)
    maxtotalvolume = max_volume(2, 13, s)

    Range("P4").Value() = "Ticker"
    Range("Q4").Value() = "Value"
    split_maxincrease = Split(maxincrease)
    Range("O5").Value() = "Greatest % Increase"
    Range("P5").Value() = split_maxincrease(2)
    Range("Q5").Value() = split_maxincrease(1) + "%"

    split_maxdecrease = Split(maxdecrease)
    Range("O6").Value() = "Greatest % decrease"
    Range("P6").Value() = split_maxdecrease(1)
    Range("Q6").Value() = split_maxdecrease(0) + "%"

    split_maxtotalvolume = Split(maxtotalvolume)
    Range("O7").Value() = "Greatest Total Volume"
    Range("P7").Value() = split_maxtotalvolume(2)
    Range("Q7").Value() = split_maxtotalvolume(1)

    MsgBox ("sheet  " + s + " is done")
    
Next 'end of looping the sheets

    End Sub 'end of stock_data



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''This functions calculate how many non-emepty data cells''
''''''''''ailable in a row', given the column and sheet'''''''''''''''''''''''''''''
'''''''''Inputs Findrows(starting row, column, sheet)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function Findrows(i As Long, j As Integer, k As String) As Long
'i is the first row to start with
'j is the column to number
'k is the sheet number
Dim No_of_rows As Long 'no of non empty rows- long uses 4bytes
                                       'normal integer with 2bytes is not sufficient
Dim sheet As String
Worksheets(k).Activate
'MsgBox (k)
No_of_rows = 0

Do While Cells(i, j).Value() <> ""
    No_of_rows = No_of_rows + 1
    i = i + 1
Loop

Findrows = No_of_rows
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''This function returns the greatest increase
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function max_increase(initial_row As Long, column As Long, k As String) As String
Dim increase As Double
Dim ticker As String
Worksheets(k).Activate
increase = -1
ticker = "error"

Do While Cells(initial_row, column).Value() <> ""

    If (Cells(initial_row, column).Value() > increase) Then
        increase = Cells(initial_row, column).Value()
        ticker = Cells(initial_row, column - 2).Value()
        
    End If
    
    initial_row = initial_row + 1
Loop

max_increase = Str(increase * 100) + " " + ticker
End Function



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''This function returns the greatest decrease
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function max_decrease(initial_row As Long, column As Long, k As String) As String
Dim decrease As Double
Dim ticker As String
Worksheets(k).Activate
decrease = 1
ticker = "error"

Do While Cells(initial_row, column).Value() <> ""

    If (Cells(initial_row, column).Value() < decrease) Then
        decrease = Cells(initial_row, column).Value()
        ticker = Cells(initial_row, column - 2).Value()
        
    End If
    
    initial_row = initial_row + 1
Loop

max_decrease = Str(decrease * 100) + " " + ticker
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''This function returns the max volume
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function max_volume(initial_row As Long, column As Long, k As String) As String
Dim volume As Double
Dim ticker As String
Worksheets(k).Activate
volume = 0
ticker = "error"

Do While Cells(initial_row, column).Value() <> ""

    If (Cells(initial_row, column).Value() > volume) Then
        volume = Cells(initial_row, column).Value()
        ticker = Cells(initial_row, column - 3).Value()
        
    End If
    
    initial_row = initial_row + 1
Loop

max_volume = Str(volume) + " " + ticker
End Function










