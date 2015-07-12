Attribute VB_Name = "Module1"
Function FileLocked(strFileName As String) As Boolean
 '****http://support.microsoft.com/kb/209189*******
 '*************************************************
   
   On Error Resume Next
   ' If the file is already opened by another process,
   ' and the specified type of access is not allowed,
   ' the Open operation fails and an error occurs.
   Open strFileName For Binary Access Read Write Lock Read Write As #1
   Close #1
   ' If an error occurs, the document is currently open.
   If Err.Number <> 0 Then
      ' Display the error number and description.
      'MsgBox "Error #" & Str(Err.Number) & " - " & Err.Description
      FileLocked = True
      Err.Clear
   End If
End Function
 


Sub GetFileLocationStakedCSV()
''****************File and Path *******************


Dim StartPos As Long
Dim CurrentPos As Long
Dim LastPos As Long
Dim Found As Integer
Dim CurrentText As String

Sheets("highlight staked").Activate


If IsEmpty(Sheets("highlight staked").Cells(2, 2)) = True Then
MyPath = Sheets("highlight staked").Cells(2, 5).Value
Else
  If Len(Dir(Sheets("highlight staked").Cells(2, 2).Value)) > 0 Then   ' check id folder last used exists
  MyPath = Sheets("highlight staked").Cells(2, 2).Value
  Else
  MyPath = Sheets("highlight staked").Cells(2, 5).Value
  End If
End If

ChDrive MyPath
ChDir MyPath
 
Application.ScreenUpdating = False

''find file
fileToOpen = Application _
    .GetOpenFilename("csv (*.csv), *.csv")
    Sheets("highlight staked").Activate
    Range("B1") = fileToOpen ''full path
     
StartPos = 0
CurrentText = Range("B1")

''find last backslash position for path/filename separation
Do While Found = 0
CurrentPos = InStr(StartPos + 1, CurrentText, "\")
StartPos = CurrentPos
Range("F1") = StartPos
   If InStr(StartPos + 1, CurrentText, "\") = 0 Then
   Found = 1
   End If
Loop
     
Range("B2") = Left(CurrentText, StartPos)
     


End Sub
Sub getstakeddata()

Sheets("highlight staked").Range("A10:Z65000").ClearContents

fileToOpen = Sheets("highlight staked").Cells(1, 2)
If fileToOpen = False Then
Exit Sub
End If

Set FSO = New FileSystemObject
Set FSOFile = FSO.OpenTextFile(fileToOpen, 1, False)
writehere = 10
   
    
   Do While Not FSOFile.AtEndOfStream
       textin = FSOFile.ReadLine ''current line stored here
       Sheets("highlight staked").Cells(writehere, 1) = textin
       writehere = writehere + 1
       
         
       
   Loop

  lastrow = Sheets("highlight staked").Cells(65000, 1).End(xlUp).Row
    

      
    Range("A10:A" & lastrow).Select
    Selection.TextToColumns Destination:=Range("A10"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=True, Space:=False, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1)), _
        TrailingMinusNumbers:=True
       
       
       
 
End Sub

Private Sub removesuffix()

       
    
    Sheets("highlight staked").Activate
    suff1 = "_StkdPt"
    suff2 = "StkdPt_"
    suff3 = "Stkd"

    lastrow = Range("B65500").End(xlUp).Row    ''find last shot

    For g = 10 To lastrow
        If IsEmpty(Cells(g, 1)) = False Then
            pointid = Cells(g, 1)
            If InStr(1, pointid, suff1, vbTextCompare) > 0 Then
                pointidnew = Replace(pointid, suff1, "", vbTextCompare)
                Cells(g, 1) = pointidnew
                GoTo resumehere1
            End If
            If InStr(1, pointid, suff2, vbTextCompare) > 0 Then
                pointidnew = Replace(pointid, suff2, "", vbTextCompare)
                Cells(g, 1) = pointidnew
                GoTo resumehere1
            End If
            If InStr(1, pointid, suff3, vbTextCompare) > 0 Then
                pointidnew = Replace(pointid, suff3, "", vbTextCompare)
                Cells(g, 1) = pointidnew
                GoTo resumehere1
            End If
        End If
    
resumehere1:
Next g



   For g = 10 To lastrow 'find leading zeros after replacing suffix/prefix
        If IsEmpty(Cells(g, 1)) = False Then
          pointid = Cells(g, 1)
          
          For j = 1 To Len(pointid)
           If Mid(pointid, j, 1) = "0" Then 'found first non-zero
           Else
           Exit For
           End If
          Next j
         
          Cells(g, 1) = Mid(pointid, j, (Len(pointid)) + 1 - j)
        
        End If
    Next g


End Sub

Sub GetFileLocationHighlightXLSX()
''****************File and Path *******************


Dim StartPos As Long
Dim CurrentPos As Long
Dim LastPos As Long
Dim Found As Integer
Dim CurrentText As String
Dim strFileName As String

Application.Cursor = xlDefault
Application.ScreenUpdating = True
Application.DisplayStatusBar = ""

Sheets("highlight staked").Activate


If IsEmpty(Sheets("highlight staked").Cells(4, 2)) = True Then
MyPath = Sheets("highlight staked").Cells(4, 5).Value
Else
  If Len(Dir(Sheets("highlight staked").Cells(4, 2).Value)) > 0 Then   ' check id folder last used exists
  MyPath = Sheets("highlight staked").Cells(4, 2).Value
  Else
  MyPath = Sheets("highlight staked").Cells(4, 5).Value
  End If
End If



ChDrive MyPath
ChDir MyPath
 

''find file
fileToOpen = Application _
    .GetOpenFilename("xlsm (*.xlsm), *.xlsm")
    Sheets("highlight staked").Activate
    Range("B3") = fileToOpen ''full path
     
StartPos = 0
CurrentText = Range("B3")

''find last backslash position for path/filename separation
Do While Found = 0
CurrentPos = InStr(StartPos + 1, CurrentText, "\")
StartPos = CurrentPos
Range("F3") = StartPos
   If InStr(StartPos + 1, CurrentText, "\") = 0 Then
   Found = 1
   End If
Loop
     
Range("B4") = Left(CurrentText, StartPos)
     

 'remember path, filename and sheetname of first excel file
 controlpath = Sheets("highlight staked").Cells(2, 18)
 controlsheet = ActiveSheet.Name
 Sheets("highlight staked").Cells(2, 19) = ActiveSheet.Name
 controlfile = Sheets("highlight staked").Cells(2, 17)
  
 'get path, filename and sheetname of second excel file
 newpath = Sheets("highlight staked").Cells(3, 2)
 newfilename = Sheets("highlight staked").Cells(4, 17)
 
 '****http://support.microsoft.com/kb/209189*******
 '*************************************************
  strFileName = Sheets("highlight staked").Cells(3, 2).Value
   ' Call function to test file lock.
   If Not FileLocked(strFileName) Then
   Else
   MsgBox ("File opened on other PC")
   Exit Sub
   End If
 '****http://support.microsoft.com/kb/209189*******
 '*************************************************
 
 Application.Workbooks.Open (newpath)
  openedsheet = ActiveSheet.Name
 

Windows(controlfile).Activate ' back to first

Sheets("highlight staked").Cells(4, 19) = openedsheet





End Sub
Sub xlsxsearchhighlight()

Dim WS_Count As Integer
Dim I As Integer


Application.ScreenUpdating = False
Application.Cursor = xlWait
Application.DisplayStatusBar = True



foundvalue = 0

openMASTER = Sheets("highlight staked").Cells(4, 17)
openCSV = Sheets("highlight staked").Cells(2, 17)


Windows(openMASTER).Activate  'electr. master file

' Set WS_Count equal to the number of worksheets in the active workbook.
WS_Count = ActiveWorkbook.Worksheets.Count

Windows(openCSV).Activate 'back to csv data

lastrowCSV = Sheets("highlight staked").Range("A65500").End(xlUp).Row ''find last shot


For j = 10 To lastrowCSV ' for all entries imported from csv check

stakedname = Sheets("highlight staked").Cells(j, 1) 'get current pointname

   For k = 1 To WS_Count ' check all sheets in Master file
    lastrowMASTER = Workbooks(openMASTER).Sheets(k).Cells(65000, 2).End(xlUp).Row  'last Easting found in row
    Application.StatusBar = "Sheet: " & Workbooks(openMASTER).Worksheets(k).Name & "   Number of Rows: " & lastrowMASTER

            For l = 2 To lastrowMASTER 'go through all entries in current sheet
               existingnames = Workbooks(openMASTER).Sheets(k).Cells(l, 2) 'get current coords

                     zzzz = StrComp(stakedname, existingnames, vbTextCompare)

                    If l = 216 Then
                    fuckthis = 1
                    End If
                    If zzzz = 0 Then 'same name

                      LNK1 = Sheets("highlight staked").Cells(3, 2)
                      LNK2 = "'" & Workbooks(openMASTER).Worksheets(k).Name & "'!" & Cells(l, 7).Address
                       With Worksheets(1)
                       .Hyperlinks.Add Anchor:=.Range("H" & j), _
                        Address:=LNK1, _
                        SubAddress:=LNK2, _
                        ScreenTip:="Location in Master File", _
                        TextToDisplay:="Link to Master"
                       End With

                        Workbooks(openMASTER).Worksheets(k).Rows(l).Interior.Pattern = xlSolid
                        Workbooks(openMASTER).Worksheets(k).Rows(l).Interior.PatternColorIndex = xlAutomatic
                        Workbooks(openMASTER).Worksheets(k).Rows(l).Interior.Color = 12611584 ' some sort of red
                        '.Color = 10040319   some sort of red
                        '.Color = 12611584     blue
                        Workbooks(openMASTER).Worksheets(k).Rows(l).Interior.TintAndShade = 0
                        Workbooks(openMASTER).Worksheets(k).Rows(l).Interior.PatternTintAndShade = 0
                        Workbooks(openMASTER).Worksheets(k).Rows(l).Interior.PatternTintAndShade = 0

                        'Windows("Elec Setout Requests - Stub Up MASTER.xlsm").Activate
                        'Workbooks(openMASTER).Worksheets(k).Cells(l, 1) = "X"
                     Else
                      Sheets("highlight staked").Cells(j, 9) = "Point not found "
                     End If



            Next l

   Next k

Next j

Application.Cursor = xlDefault

Application.ScreenUpdating = True
Application.DisplayStatusBar = ""

End Sub

Sub sortbyexisting()

  lastrow = Sheets("main").Cells(65000, 2).End(xlUp).Row

For o = 10 To lastrow


If IsEmpty(Sheets("main").Cells(o, 9)) = True Then
Cells(o, 10) = 0
Else
Cells(o, 10) = 1
End If

Next o



    ActiveWorkbook.Worksheets("main").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("main").Sort.SortFields.Add Key:=Range("J10:J" & lastrow), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("main").Sort
        .SetRange Range("A10:S" & lastrow)
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With











End Sub



















'*******************************************************************************************
'******************** bring in Engineers data



Sub GetFileLocationCSV()
''****************File and Path *******************


Dim StartPos As Long
Dim CurrentPos As Long
Dim LastPos As Long
Dim Found As Integer
Dim CurrentText As String

Sheets("main").Activate


If IsEmpty(Sheets("main").Cells(2, 2)) = True Then
MyPath = Sheets("main").Cells(2, 5).Value
Else
  If Len(Dir(Sheets("main").Cells(2, 2).Value)) > 0 Then   ' check id folder last used exists
  MyPath = Sheets("main").Cells(2, 2).Value
  Else
  MyPath = Sheets("main").Cells(2, 5).Value
  End If
End If

ChDrive MyPath
ChDir MyPath
 
Application.ScreenUpdating = False

''find file
fileToOpen = Application _
    .GetOpenFilename("csv (*.csv), *.csv")
    Sheets("main").Activate
    Range("B1") = fileToOpen ''full path
     
StartPos = 0
CurrentText = Range("B1")

''find last backslash position for path/filename separation
Do While Found = 0
CurrentPos = InStr(StartPos + 1, CurrentText, "\")
StartPos = CurrentPos
Range("F1") = StartPos
   If InStr(StartPos + 1, CurrentText, "\") = 0 Then
   Found = 1
   End If
Loop
     
Range("B2") = Left(CurrentText, StartPos)
     
Worksheets("main").TextBox4.Value = Sheets("main").Cells(1, 9).Value



End Sub



Sub GetFileLocationXLSX()
''****************File and Path *******************


Dim StartPos As Long
Dim CurrentPos As Long
Dim LastPos As Long
Dim Found As Integer
Dim CurrentText As String

Sheets("main").Activate


If IsEmpty(Sheets("main").Cells(4, 2)) = True Then
MyPath = Sheets("main").Cells(4, 5).Value
Else
  If Len(Dir(Sheets("main").Cells(4, 2).Value)) > 0 Then   ' check id folder last used exists
  MyPath = Sheets("main").Cells(4, 2).Value
  Else
  MyPath = Sheets("main").Cells(4, 5).Value
  End If
End If

ChDrive MyPath
ChDir MyPath
 

''find file
fileToOpen = Application _
    .GetOpenFilename("xlsm (*.xlsm), *.xlsm")
    Sheets("main").Activate
    Range("B3") = fileToOpen ''full path
     
StartPos = 0
CurrentText = Range("B3")

''find last backslash position for path/filename separation
Do While Found = 0
CurrentPos = InStr(StartPos + 1, CurrentText, "\")
StartPos = CurrentPos
Range("F3") = StartPos
   If InStr(StartPos + 1, CurrentText, "\") = 0 Then
   Found = 1
   End If
Loop
     
Range("B4") = Left(CurrentText, StartPos)
     


End Sub




Sub getsetoutdata()

Sheets("main").Range("A10:Z65000").ClearContents

fileToOpen = Sheets("main").Cells(1, 2)
If fileToOpen = False Then
Exit Sub
End If

Set FSO = New FileSystemObject
Set FSOFile = FSO.OpenTextFile(fileToOpen, 1, False)
writehere = 10
   
    
   Do While Not FSOFile.AtEndOfStream
       textin = FSOFile.ReadLine ''current line stored here
       Sheets("main").Cells(writehere, 1) = textin
       writehere = writehere + 1
       
         
       
   Loop

  lastrow = Sheets("main").Cells(65000, 1).End(xlUp).Row
    

      
    Range("A10:A" & lastrow).Select
    Selection.TextToColumns Destination:=Range("A10"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=True, Space:=False, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1)), _
        TrailingMinusNumbers:=True
       
       
       
 
End Sub


Sub getmasterfile()
Dim strFileName As String

 'remember path, filename and sheetname of first excel file
 controlpath = Sheets("main").Cells(2, 18)
 controlsheet = ActiveSheet.Name
 Sheets("main").Cells(2, 19) = ActiveSheet.Name
 controlfile = Sheets("main").Cells(2, 17)
  
 'get path, filename and sheetname of second excel file
 newpath = Sheets("main").Cells(3, 2)
 newfilename = Sheets("main").Cells(4, 17)
 
 '****http://support.microsoft.com/kb/209189*******
 '*************************************************
  strFileName = Sheets("main").Cells(3, 2)
   ' Call function to test file lock.
   If Not FileLocked(strFileName) Then
   Else
   MsgBox ("File opened on other PC")
   Exit Sub
   End If
 '****http://support.microsoft.com/kb/209189*******
 '*************************************************
 
 Application.Workbooks.Open (newpath)
  openedsheet = ActiveSheet.Name
 

Windows(controlfile).Activate ' back to first

Sheets("main").Cells(4, 19) = openedsheet


End Sub


Sub xlsxsearch()
 
Dim WS_Count As Integer
Dim I As Integer

Application.ScreenUpdating = False
Application.Cursor = xlWait
Application.DisplayStatusBar = True
 
Range("H10:I65000").ClearContents


maxdeviation = Sheets("main").TextBox1.Value * 1#
foundvalue = 0

openMASTER = Sheets("main").Cells(4, 17)
openCSV = Sheets("main").Cells(2, 17)


Windows(openMASTER).Activate  'electr. master file

' Set WS_Count equal to the number of worksheets in the active workbook.
WS_Count = ActiveWorkbook.Worksheets.Count

Windows(openCSV).Activate 'back to csv data

lastrowCSV = Sheets("main").Range("B65500").End(xlUp).Row ''find last shot

    
For j = 10 To lastrowCSV ' for all entries imported from csv check
    
EastCSV = Sheets("main").Cells(j, 2) 'get current coords
NorthCSV = Sheets("main").Cells(j, 3)
    
    
   For k = 1 To WS_Count ' check all sheets in Master file
    lastrowMASTER = Workbooks(openMASTER).Sheets(k).Cells(65000, 2).End(xlUp).Row  'last Easting found in row
         Application.StatusBar = "Checking Sheet: " & Workbooks(openMASTER).Worksheets(k).Name & "    Number of Lines:  " & lastrowMASTER
         
              For l = 2 To lastrowMASTER 'go through all entries in current sheet
               EastMASTER = Workbooks(openMASTER).Sheets(k).Cells(l, 3) 'get current coords
               NorthMASTER = Workbooks(openMASTER).Sheets(k).Cells(l, 4) 'get current coords
                   
                   slopedist = Sqr((EastCSV - EastMASTER) ^ 2 + (NorthCSV - NorthMASTER) ^ 2) ' distance between both points
                    If slopedist < maxdeviation Then 'within tolerance
                      LNK1 = Sheets("main").Cells(3, 2)
                      LNK2 = "'" & Workbooks(openMASTER).Worksheets(k).Name & "'!" & Cells(l, 7).Address
                       With Worksheets(1)
                       .Hyperlinks.Add Anchor:=.Range("H" & j), _
                        Address:=LNK1, _
                        SubAddress:=LNK2, _
                        ScreenTip:="Location in Master File", _
                        TextToDisplay:="Link to Master"
                       End With
                      Sheets("main").Cells(j, 9) = "Distance between Points: " & Format(slopedist, "0.000") 'write a text in column 8
                      foundvalue = 1
                      Exit For
                    End If
                
            Next l
      
      
     If foundvalue = 1 Then ' if we found our value in the loop through all rows we can stop looking through all worksheets as well
      foundvalue = 0 'reset for next point
      Exit For
     End If
       
   Next k
    
    
    
Next j
    
Application.Cursor = xlDefault

Application.ScreenUpdating = True
Application.StatusBar = ""

End Sub


Sub stripcodefromstring()
Dim inputstring As String
Dim searchstring As String



  lastrow = Sheets("main").Cells(65000, 2).End(xlUp).Row

For m = 10 To lastrow

inputstring = Cells(m, 1)
searchstring = "-"

If InStr(1, inputstring, searchstring, vbTextCompare) = 0 Then
'do nothin if the search string wasn't found
Else
    
    ''find last searchstring position for string separation
      CurrentPos = Len(inputstring)
      
      Do While Found = 0
        xxx = Mid(inputstring, CurrentPos, 1)
        If xxx = searchstring Then
         Found = 1
         Cells(m, 8) = CurrentPos + 1
         Cells(m, 5) = Mid(inputstring, CurrentPos + 1, Len(inputstring) - CurrentPos)
         Cells(m, 8) = ""
         Cells(m, 1) = ""
         End If
        CurrentPos = CurrentPos - 1
      Loop
  Found = 0

End If

Next m




End Sub

Sub numbers()

  lastrow = Sheets("main").Cells(65000, 2).End(xlUp).Row
nextnum = (Sheets("main").TextBox2.Value)




For n = 10 To lastrow

numlength = Len(Sheets("main").TextBox2.Value)
numstring = nextnum
If numlength < 2 Then
numstring = "0000" & nextnum
GoTo resumehere
End If
If numlength < 3 Then
numstring = "000" & nextnum
GoTo resumehere
End If
If numlength < 4 Then
numstring = "00" & nextnum
GoTo resumehere
End If
If numlength < 5 Then
numstring = "0" & nextnum
GoTo resumehere
End If

resumehere:

Cells(n, 1) = Sheets("main").TextBox3.Value & numstring
nextnum = nextnum + 1



Next n




End Sub



