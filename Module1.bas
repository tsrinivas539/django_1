Attribute VB_Name = "Module1"
Sub MyButton()

Application.AskToUpdateLinks = False
Application.EnableEvents = False
Application.DisplayAlerts = False



Dim lastrow As Long, lastcol As Long, last As Long, newlast As Long, flast As Long, dlast As Long, nlast As Long
Dim rng As Range, Rng1 As Range
Dim t As Integer
Dim m As Integer, n As Integer

'Captures username of system running macro
Dim name As String
name = Environ$("UserName")

Dim dashname As String
dashname = Left(ThisWorkbook.name, Len(ThisWorkbook.name) - 5)

Dim a As Workbook
Dim file As String

Dim berror As Integer, serror As Integer, verror As Integer, derror As Integer
berror = 0
serror = 0
verror = 0
derror = 0
aerror = 0
aval = 0

Dim w As Integer, y As Integer, z As Integer, s As Integer, u As Integer
w = 0
y = 0
z = 0
s = 0
u = 0

'declares empty arrays tocapture no. of errors
Dim verr(10000) As String
Dim berr(10000) As String
Dim serr(10000) As String
Dim derr(10000) As String
Dim aerr(10000) As String

Dim extn As Variant, x As Variant

extn = Array("*.xlsx", "*.xlsm") 'setting file extension

For t = 0 To 1

file = Dir("C:\Users\" & name & "\Desktop\ASIN Uploads\" & extn(t)) 'combining the folder path with extension

Do While file <> "" 'loop begins, and continues until there are no files left

'checking file status against tracker
Dim trk As Workbook
Set trk = Workbooks.Open("\\ant\dept-as\Hyd11\Localization\Exclusions\Asin Exclusions - New Workflow\OPS\Ops Tracker\" & dashname & "_ASIN Tracker.xlsm")

nlast = trk.Sheets("Assign").Cells(trk.Sheets("Assign").Rows.count, 1).End(xlUp).Row

For i = 2 To nlast
If trk.Sheets("Assign").Cells(i, 1) = name And trk.Sheets("Assign").Cells(i, 3) = file Then
If trk.Sheets("Assign").Cells(i, 8) = "Assigned" Then
trk.Save
trk.Close
Else:
trk.Save
trk.Close
GoTo finish
End If
Else: End If
Next i


Application.EnableEvents = False
Set a = Workbooks.Open(filename:="C:\Users\" & name & "\Desktop\ASIN Uploads\" & file)
Application.EnableEvents = True

a.Sheets("Sheet1").Unprotect

Dim part As String
part = a.Sheets("Sheet1").Cells(2, 2)


For i = 1 To a.Worksheets.count 'check if Sheet2 exists, else adds it
If a.Worksheets(i).name = "Sheet2" Then
    exists = True
    End If
Next i
If Not exists Then
    a.Worksheets.Add.name = "Sheet2"
End If

a.Sheets("Sheet2").Cells.Clear 'empties contents of Sheet 2

a.Sheets("Sheet1").Activate

a.Sheets("Sheet1").Rows.Hidden = False 'unhide all rows and columns
a.Sheets("Sheet1").Columns.Hidden = False

lastrow = a.Sheets("Sheet1").Cells(a.Sheets("Sheet1").Rows.count, 1).End(xlUp).Row
lastcol = 40

'Set variables to capture current date and time
Dim end_t As Variant, end_d As Variant
end_t = TimeValue(Now)
end_d = DateValue(Now)


a.Sheets("Sheet1").Range("J2:J" & lastrow).Interior.ColorIndex = 2

'date check
Dim dcheck As Long
dcheck = a.Sheets("Sheet1").Cells(a.Sheets("Sheet1").Rows.count, "R").End(xlUp).Row
If a.Sheets("Sheet1").Cells(1, 18) = "Date" And dcheck < 2 Then
MsgBox "You see to have entered the dates in the wrong column. Please enter them in column 'R' and attempt the upload again."
GoTo finish
Else: End If


'Find column number with unique identifier
For i = 1 To lastcol
If a.Sheets("Sheet1").Cells(1, i) = "ID" Then
m = i
Else: End If
Next

'Check to ensure file has unique identifier column which defines each record
If m = 0 Then
MsgBox ("Please ensure that your file has a unique identifier column before running the code")
GoTo complete
Else: End If

    Dim xyz As Integer

    a.Sheets("Sheet2").Cells.Clear 'empties contents of Sheet 2

    Set rng = a.Sheets("Sheet1").Range(Cells(1, 1), Cells(lastrow, lastcol))
    xyz = Application.WorksheetFunction.CountIf(rng, end_d)
    rng.AutoFilter Field:=18, Criteria1:="<>", Operator:=xlAnd, Criteria2:=end_d
    Set Rng1 = a.Sheets("Sheet1").Range(Cells(2, 1), Cells(lastrow, lastcol))
    If xyz = 0 Then  'check if records containing today's date are present
    GoTo final
    Else: Rng1.SpecialCells(xlCellTypeVisible).Copy
    End If
    a.Sheets("Sheet2").Range("A1").PasteSpecial 'only records containing today's date are copied into blank sheet
    Application.CutCopyMode = False
    rng.AutoFilter


last = a.Sheets("Sheet2").Cells(a.Sheets("Sheet2").Rows.count, 1).End(xlUp).Row

Dim fname, sname As String
fname = a.FullName
sname = a.name

'contingency in case autofilter fails
For j = 1 To last
If a.Sheets("Sheet2").Cells(j, 18) = "" Then
a.Sheets("Sheet2").Rows(j).Delete
Else: End If
Next

last = a.Sheets("Sheet2").Cells(a.Sheets("Sheet2").Rows.count, 1).End(xlUp).Row

If last = 1 Then
a.Sheets("Sheet2").Cells(1, 19) = name 'fills name of ops specialist in column
Else
a.Sheets("Sheet2").Cells(1, 19) = name 'fills name of ops specialist in column
a.Sheets("Sheet2").Range("S1:S" & last).FillDown
End If

With a.Sheets("Sheet2")

For i = 1 To last

If .Cells(i, 10) = "3" And .Cells(i, 11) <> " " And IsEmpty(.Cells(i, 11)) = False Then
verror = verror + 1
verr(w) = a.Sheets("Sheet2").Cells(i, m)
w = w + 1
Else: End If

If .Cells(i, 10) = "1" And IsEmpty(.Cells(i, 11)) = True Then
derror = derror + 1
derr(s) = a.Sheets("Sheet2").Cells(i, m)
s = s + 1
Else: End If

If .Cells(i, 10) = "" Or IsEmpty(.Cells(i, 10)) = True Then
berror = berror + 1
berr(u) = a.Sheets("Sheet2").Cells(i, m)
u = u + 1
Else: End If

Next i

End With

If w > 0 Then

For j = 2 To lastrow
For k = 0 To w
If a.Sheets("Sheet1").Cells(j, m) = verr(k) Then
a.Sheets("Sheet1").Cells(j, 10).Interior.ColorIndex = 3
Else: End If
Next k
Next j

Else: End If

If s > 0 Then

For j = 2 To lastrow
For k = 0 To s
If a.Sheets("Sheet1").Cells(j, m) = derr(k) Then
a.Sheets("Sheet1").Cells(j, 10).Interior.ColorIndex = 3
Else: End If
Next k
Next j

Else: End If

If u > 0 Then

For j = 2 To lastrow
For k = 0 To u
If a.Sheets("Sheet1").Cells(j, m) = berr(k) Then
a.Sheets("Sheet1").Cells(j, 10).Interior.ColorIndex = 3
Else: End If
Next k
Next j

Else: End If


'throws error message to ops
If verror > 0 Then
MsgBox "You have filled a reason beside an ASIN with 3 classification. Please correct the same."
GoTo final1
Else: End If

If derror > 0 Then
MsgBox "You have not filled a reason beside an ASIN with 1 classification. Please correct the same."
GoTo final1
Else: End If

If berror > 0 Then
MsgBox "You seem to have missed marking a classification beside an ASIN. Please correct the same."
GoTo final1
Else: End If

'Open data dump
Dim b As Workbook
Set b = Workbooks.Open("\\ant\dept-as\Hyd11\Localization\Exclusions\Asin Exclusions - New Workflow\OPS\Ops associate-wise dumps\" & dashname & "_ASIN Dump.xlsm")

b.Sheets("Sheet1").Unprotect "Data1104"

newlast = b.Sheets("Sheet1").Cells(b.Sheets("Sheet1").Rows.count, 1).End(xlUp).Row

If b.Sheets("Sheet1").Cells(newlast, 2) = part Then

'Find count of duplicate values from dump
Dim dup As Integer
dup = 0

For i = 1 To last
For j = 2 To newlast
If Int(a.Sheets("Sheet2").Cells(i, m)) = Int(b.Sheets("Sheet1").Cells(j, m)) And a.Sheets("Sheet2").Cells(i, 2) = b.Sheets("Sheet1").Cells(j, 2) Then
dup = dup + 1
Else: End If
Next j
Next i



newlast = newlast + 1

On Error GoTo errline

With a.Sheets("Sheet2").UsedRange

  b.Sheets("Sheet1").Range("A" & newlast).Resize( _
        .Rows.count, .Columns.count) = .Value
End With

End If

errline:

On Error GoTo 0

If b.Sheets("Sheet1").Cells(newlast - 1, 2) = part Then

If newlast <= 3500 Then

For i = 1 To last
For j = 2 To newlast
If Int(a.Sheets("Sheet2").Cells(i, m)) = Int(b.Sheets("Sheet1").Cells(j, m)) And a.Sheets("Sheet2").Cells(i, 2) = b.Sheets("Sheet1").Cells(j, 2) Then
If Int(b.Sheets("Sheet1").Cells(j, 18)) <> Int(end_d) Then
b.Sheets("Sheet1").Rows(j).Delete
Else: End If
Else: End If
Next j
Next i

ElseIf newlast > 3500 Then

For i = 1 To last
For j = (newlast - 2000) To newlast
If Int(a.Sheets("Sheet2").Cells(i, m)) = Int(b.Sheets("Sheet1").Cells(j, m)) And a.Sheets("Sheet2").Cells(i, 2) = b.Sheets("Sheet1").Cells(j, 2) Then
If Int(b.Sheets("Sheet1").Cells(j, 18)) <> Int(end_d) Then
b.Sheets("Sheet1").Rows(j).Delete
Else: End If
Else: End If
Next j
Next i

Else: End If

Else: End If

b.Sheets("Sheet1").UsedRange.RemoveDuplicates Columns:=Array(2, m), Header:=xlYes 'removes duplicate records from test

b.Sheets("Sheet1").Protect "Data1104"

b.Close savechanges:=True 'saves and closes collated file

a.Sheets("Sheet2").Delete

'Opens tracker to note details of upload
Dim c As Workbook
Application.EnableEvents = False
Set c = Workbooks.Open("\\ant\dept-as\Hyd11\Localization\Exclusions\Asin Exclusions - New Workflow\OPS\Ops Tracker\" & dashname & "_ASIN Tracker.xlsm")
Application.EnableEvents = True

c.Sheets("Upload").Unprotect "Prod1104"
c.Sheets("Assign").Unprotect "Prod1104"
c.Sheets("File Record").Unprotect "Prod1104"
c.Sheets("Sheet3").Unprotect "Prod1104"

slast = c.Sheets("File Record").Cells(c.Sheets("File Record").Rows.count, 1).End(xlUp).Row

For j = slast To 2 Step -1

If file = c.Sheets("File Record").Cells(j, 3) Then
start_t = c.Sheets("File Record").Cells(j, 5)
GoTo st_time
Else: End If

Next j

st_time:

flast = c.Sheets("Upload").Cells(c.Sheets("Upload").Rows.count, 1).End(xlUp).Row

If flast <> 0 Then
flast = flast + 1
Else: End If

finishtime = TimeValue(Now())

'Calculates duration of macro runtime
dur = (finishtime - end_t)

c.Sheets("Upload").Cells(flast, 1) = name 'ops userid captured
c.Sheets("Upload").Cells(flast, 2) = fname 'filename captured
c.Sheets("Upload").Cells(flast, 3) = sname 'filename captured
c.Sheets("Upload").Cells(flast, 4) = (last - dup) 'count of unique records uploaded captured
c.Sheets("Upload").Cells(flast, 5) = end_d 'upload date captured
c.Sheets("Upload").Cells(flast, 6) = Format(start_t, "hh:mm:ss") 'time at which ops started working on file captured
c.Sheets("Upload").Cells(flast, 7) = Format(end_t, "hh:mm:ss") 'upload time captured
c.Sheets("Upload").Cells(flast, 8) = end_d
c.Sheets("Upload").Cells(flast, 9) = 0 'count of unique records uploaded captured
c.Sheets("Upload").Cells(flast, 10) = Format(dur, "hh:mm:ss") 'macro runtime duration captured

c.Save

Dim sumval As Integer

flast = flast + 1

For j = 2 To flast
If InStr(c.Sheets("Upload").Cells(j, 3), part) > 0 Then
sumval = c.Sheets("Upload").Cells(j, 4) + sumval
Else: End If
Next j

alast = c.Sheets("Assign").Cells(c.Sheets("Assign").Rows.count, 1).End(xlUp).Row

'Check if all values have been uploaded and change file status accordingly
For k = 2 To alast
totval = c.Sheets("Assign").Cells(k, 4)
If InStr(c.Sheets("Assign").Cells(k, 3), part) > 0 And sumval >= totval Then
c.Sheets("Assign").Cells(k, 8) = "QC Pending"

c.Sheets("Upload").Protect "Prod1104"
c.Sheets("Assign").Protect "Prod1104"
c.Sheets("File Record").Protect "Prod1104"
c.Sheets("Sheet3").Protect "Prod1104"

c.Save
c.Close

a.Sheets("Sheet1").Cells(1, 19) = "Ops ID"

a.Sheets("Sheet1").Range("XFA145000") = name

a.Sheets("Sheet1").Range("S2:S" & lastrow).ClearContents

a.Sheets("Sheet1").Protect "OpsDone1104", DrawingObjects:=False, Contents:=True, Scenarios:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True, AllowFiltering:=True, AllowFormattingCells:=True

a.Save
a.Close

'Moves completed file to location on shared drive for next step of process
srcfile = "C:\Users\" & name & "\Desktop\ASIN Uploads\" & file
FileCopy srcfile, "\\ant\dept-as\Hyd11\Localization\Exclusions\Asin Exclusions - New Workflow\OPS\QC Pending\" & file

MsgBox ("You've successfully completed and uploaded your file!")

GoTo complete

Else: End If
Next k

c.Sheets("Upload").Protect "Prod1104"
c.Sheets("Assign").Protect "Prod1104"
c.Sheets("File Record").Protect "Prod1104"
c.Sheets("Sheet3").Protect "Prod1104"

c.Save
c.Close

Application.DisplayAlerts = False 'deletes temp sheet 2
'a.Sheets("Sheet2").Delete

finish:
a.Save 'saves and closes the file
a.Close

file = Dir 'gets next file

Loop
Next

MsgBox "You have successfully uploaded " & (last - dup) & " values!"

Application.DisplayAlerts = False

GoTo complete

final:

MsgBox "You seem to have clicked on the wrong button. Please try again."

Application.DisplayAlerts = False 'deletes temp sheet 2

final1:

a.Sheets("Sheet2").Delete

a.Save 'saves and closes the file
a.Close

GoTo complete

Err_Log:

ThisWorkbook.Sheets("Error Log").Unprotect "SpecOps1104"

elast = ThisWorkbook.Sheets("Error Log").Cells(ThisWorkbook.Sheets("Error Log").Rows.count, 1).End(xlUp).Row

If elast <> 0 Then
elast = elast + 1
Else: End If

ThisWorkbook.Sheets("Error Log").Cells(elast, 1) = name
ThisWorkbook.Sheets("Error Log").Cells(elast, 2) = file
ThisWorkbook.Sheets("Error Log").Cells(elast, 3) = err.Number
ThisWorkbook.Sheets("Error Log").Cells(elast, 4) = err.Description
ThisWorkbook.Sheets("Error Log").Cells(elast, 5) = err_row
ThisWorkbook.Sheets("Error Log").Cells(elast, 6) = Application.VBE.ActiveCodePane.CodeModule.name
ThisWorkbook.Sheets("Error Log").Cells(elast, 7) = DateValue(Now())
ThisWorkbook.Sheets("Error Log").Cells(elast, 8) = TimeValue(Now())

ThisWorkbook.Sheets("Error Log").Protect "SpecOps1104", DrawingObjects:=False, Contents:=True, Scenarios:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True, AllowFiltering:=True, AllowFormattingCells:=True

complete:

Application.DisplayAlerts = True
Application.EnableEvents = True

End Sub
