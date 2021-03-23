Sub ticker_count()

'VARIABLES DECLARATION
' Set an initial variable for holding the ticker name
Dim Ticker_name As String
'Set an initial variable for holding the ticker value change
Dim Ticker_value_change As Double
'Set an initial variable to support the value change comparisson and calculation
Dim Ticker_initial_value As Double
Dim Ticker_final_value As Double
'Set an initial variable for the loop counter that will identify the initial cell for value comparisson
Dim Loop_counter As Integer
'Set an initial variable for holding the ticker percentage change
Dim Ticker_percentage_change As Double
' Set an initial variable for holding the ticker volume
Dim Ticker_volume As LongLong
Ticker_volume = 0

'SET PROCESS FACILITATORS
'Keep track of the location for each ticker name
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2
Loop_counter = 0
'Identify the last row with information to support the loop operation
Dim ult As Long
ult = Cells(Rows.Count, 1).End(xlUp).Row

'STAR THE PROCESS
'Loop through all ticker names
    For i = 2 To ult
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        'LOGICAL PROCESS
            'Set Ticker name
            Ticker_name = Cells(i, 1).Value
            'Set Ticker value change
            Ticker_initial_value = Cells((i - (Loop_counter)), 3).Value
            If Ticker_initial_value = 0 Then
                Ticker_initial_value = 0 + 0.001
            End If
            Ticker_value_change = Cells(i, 6).Value - Ticker_initial_value
            'Set Ticker percntage change
            If Cells(i, 6).Value = 0 Then
            Cells(i, 6).Value = Cells(i, 6).Value + 0.001
            End If
            Ticker_percentage_change = 1 - (Cells(i, 6).Value / Ticker_initial_value)
             'Set Ticker volume
            Ticker_volume = Ticker_volume + Cells(i, 7).Value
        'PRINTING VALUES
            'Print the Ticker name in the summary table
            Range("J" & Summary_Table_Row).Value = Ticker_name
            'Print the Ticker volume in the summary table
            Range("M" & Summary_Table_Row).Value = Ticker_volume
            'Print the Ticker value change
            Range("K" & Summary_Table_Row).Value = Ticker_value_change
            'Formating cells color of Ticker value change
            If Ticker_percentage_change >= 0 Then
                Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
               Else
               Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
            End If
            'Print the Ticker percentage change
            Range("L" & Summary_Table_Row).Value = Ticker_percentage_change
            'Formating cells for percentage change
            Range("L:L").NumberFormat = "0.00%"
        'SEQUENCE TO FILL THE TABLE
            'Add one to the summary table row'
            Summary_Table_Row = Summary_Table_Row + 1
            Ticker_volume = 0
            Ticker_initial_value = 0
            Ticker_value_change = 0
            Loop_counter = 0
        Else
            Ticker_volume = Ticker_volume + Cells(i, 7).Value
            Loop_counter = Loop_counter + 1
        End If
    Next i
    
End Sub