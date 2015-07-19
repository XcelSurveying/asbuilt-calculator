Attribute VB_Name = "mPrint"
Sub PrintToPrinter()
    
    Dim col(1) As String
    Dim row(1) As String
    
' Initilize Print Area
    ' start cell
    col(0) = "A"
    row(0) = "6"
    ' end cell
    col(1) = "AA"
    row(1) = mPrint.lastrowSheet(2, 26) + 1
    
' Set Print Area
    Worksheets("Summary Point Report").PageSetup.PrintArea = col(0) & row(0) & ":" & col(1) & row(1)
    
' Show Print Dialog
    Application.Dialogs(xlDialogPrint).Show
    
End Sub


Function lastrowSheet(i As Integer, j As Integer) As Integer
' # Determines the integer value of lastrow of data in ActiveSheet column
' # between column i and j
' #
' # INPUT: i as integer # 1st column to start lastrow search
' #        j as Integer # last column to finish lastrow search
' #
' # OUTPUT: lastrowSheet as Integer # Lastrow between columns i and j
    
Dim lastrowCol As Integer, iCol As Integer
        
' Initilize variables
    lastrowSheet = 0
    
' ________________________________________________________________
' ############## Find the lastrow used in the sheet  #############
    For iCol = i To j
        
        With ActiveSheet
            lastrowCol = .Cells(.Rows.Count, iCol).End(xlUp).row
        End With

        ' updates the sheet last column if the current row is longer
        If lastrowCol > lastrowSheet Then
            lastrowSheet = lastrowCol
        End If
    Next iCol
' ________________________________________________________________

End Function
