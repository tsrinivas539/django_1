Sub myfileTowork()

        Server_Name = "IN-HY11-PBKMKX7" ' Enter your server name here
        Database_Name = "daily_task_1" ' Enter your database name here
        Emp_ID_SR = "root" ' enter your user ID here
        Password1 = "Studios2017" ' Enter your Password1 here

        Set cn = New ADODB.Connection
        cn.Open "Driver={MySQL ODBC 5.3 Unicode Driver};Server=" & Server_Name & ";Database=" & Database_Name & _
        ";Uid=" & Emp_ID_SR & ";Pwd=" & Password1 & ";"

Application.DisplayAlerts = False

ThisWorkbook.Sheets("Language").Unprotect "SpecOps1104"

Dim name As String
name = Environ$("Username")

Dim dashname As String
dashname = Left(ThisWorkbook.name, Len(ThisWorkbook.name) - 5)

Dim i As Integer

'qc tracker check
Dim ctrk As Workbook
'Set ctrk = Workbooks.Open("\\ant\dept-as\Hyd11\Localization\Exclusions\1_Srinivas\Asin Exclusions - New Workflow\QC\Audit Tracker\" & dashname & "_ASIN QC Tracker.xlsm")
Set ctrk = Workbooks.Open("\\ant\dept-as\HYD11\Localization\Exclusions\1_Srinivas\Asin Exclusions - New Workflow\QC\Audit Tracker\" & dashname & "_ASIN QC Tracker.xlsm")

Dim cnewlast As Long
cnewlast = ctrk.Sheets("Assign").Cells(ctrk.Sheets("Assign").Rows.count, 1).End(xlUp).Row

For i = 2 To cnewlast
On Error Resume Next
If ctrk.Sheets("Assign").Cells(i, 1) = name Then
If ctrk.Sheets("Assign").Cells(i, 13) = "QC Assigned" Then
ctrk.Close
MsgBox ("You seem to have a QC file assigned to you. Please complete the same before downloading an ops file!")
GoTo final
Else: End If
Else: End If
Next

Dim cnewlast2 As Long
cnewlast2 = ctrk.Sheets("Upload").Cells(ctrk.Sheets("Upload").Rows.count, 1).End(xlUp).Row

rec_up_time_QC = ctrk.Sheets("Upload").Cells(cnewlast2, 7).Value
rec_up_date_QC = ctrk.Sheets("Upload").Cells(cnewlast2, 5).Value

ctrk.Close

'Opens ops tracker
Dim trk As Workbook
'Set trk = Workbooks.Open("\\ant\dept-as\Hyd11\Localization\Exclusions\1_Srinivas\Asin Exclusions - New Workflow\OPS\Ops Tracker\" & dashname & "_ASIN Tracker.xlsm")
Set trk = Workbooks.Open("\\ant\dept-as\HYD11\Localization\Exclusions\1_Srinivas\Asin Exclusions - New Workflow\OPS\Ops Tracker\" & dashname & "_ASIN Tracker.xlsm")

Dim newlast As Long
newlast = trk.Sheets("Assign").Cells(trk.Sheets("Assign").Rows.count, 1).End(xlUp).Row


'checks to ensure there are no pending files
For i = 2 To newlast
On Error Resume Next
If trk.Sheets("Assign").Cells(i, 1) = name Then
If trk.Sheets("Assign").Cells(i, 8) <> "QC Pending" Then
trk.Save
'trk.Close
MsgBox ("Please complete the pending Ops file before downloading another one!")
GoTo final
Else: End If
Else: End If
Next

Dim newlast2 As Long
newlast2 = trk.Sheets("Upload").Cells(trk.Sheets("Upload").Rows.count, 1).End(xlUp).Row

rec_up_time_OPs = trk.Sheets("Upload").Cells(newlast2, 7).Value
rec_up_date_OPs = trk.Sheets("Upload").Cells(newlast2, 5).Value

trk.Close



Dim dtrk As Workbook
Set dtrk = Workbooks.Open("\\ant\dept-as\Hyd11\Localization\Exclusions\1_Srinivas\Asin Exclusions - New Workflow\TBM\TBM_Trackers\" & dashname & "_ASIN TBM Tracker.xlsm")

newlast = dtrk.Sheets("Assign").Cells(trk.Sheets("Assign").Rows.count, 1).End(xlUp).Row


'checks to ensure there are no pending files
For i = 2 To newlast
On Error Resume Next
If dtrk.Sheets("Assign").Cells(i, 1) = name Then
If dtrk.Sheets("Assign").Cells(i, 8) <> "QC Pending" Then
dtrk.Close
MsgBox ("You seem to have a manual Eyeball file assigned to you. Please complete the same before downloading another one!")
GoTo final
Else: End If
Else: End If
Next

dtrk.Close


Dim langlast As Long
langlast = ThisWorkbook.Sheets("Language").Cells(ThisWorkbook.Sheets("Language").Rows.count, 1).End(xlUp).Row

j = 0

For j = 2 To langlast
    With ThisWorkbook.Sheets("Language")
    ' check to give only ops files if the requesting person is not eligible for QC
        If .Cells(j, 2) = name And .Cells(j, 4) = "" Then

            'If Dir("\\ant\dept-as\Hyd11\Localization\Exclusions\1_Sachin\AEW\Ops\Assigning\UK\") <> "" Then
            'Need to replace this with the count of eligible files that needs to be assigned from Database table for Ops_UK table
                ' trying to lock the tables to eliminate the incident of 2 requests to the datatable to find the count
                Set rs = Nothing
                Set rs = New ADODB.Recordset
                lock_tbl = "LOCK TABLE asin_exclusion.Ops_Assgn_Man_UK WRITE;"
                rs.Open lock_tbl, cn


            Set rs = Nothing
            Set rs = New ADODB.Recordset
            cnt_File_Ops_Uk = "select count(File_name) as f_cnt from asin_exclusion.Ops_Assgn_Man_UK where Transaction_Date is null and Login_ID is null;"
            rs.Open cnt_File_Ops_Uk, cn

            file_cnt_uk = rs![f_cnt]

                Set rs = Nothing
                Set rs = New ADODB.Recordset
                ulk_tbls = "UNLOCK TABLES;"
                rs.Open ulk_tbls, cn

            If file_cnt_uk > 0 Then

           ' If Dir("\\ant\dept-as\HYD11\Localization\Exclusions\1_Srinivas\AEW\Ops\Assigning\UK") <> "" Then

                Call GetMyFile_Ops
                GoTo final
            Else
                MsgBox ("There are no UK Ops files available to download!!")
            End If
        End If
    End With

Next


For j = 2 To langlast

With ThisWorkbook.Sheets("Language")
' check to give Automated QC files if the requesting person is eligible for QC (A)
    If .Cells(j, 2) = name And .Cells(j, 4) = "A" Then

        'Need to replace this with the count of eligible files that needs to be assigned from Database table for QC_UK table
        'If Dir("\\ant\dept-as\Hyd11\Localization\Exclusions\1_Sachin\AEW\QC\QC Assigning_automated\UK\") <> "" Then

        Set rs = Nothing
        Set rs = New ADODB.Recordset
        cnt_File_QC_Aut_Uk = "select count(File_name) as f_cnt from asin_exclusion.qc_assgn_auto_uk where Transaction_Date is null and Login_ID is null;"
        rs.Open cnt_File_QC_Aut_Uk, cn

        file_cnt_QC_Aut_uk = rs![f_cnt]

        If file_cnt_QC_Aut_uk > 0 Then

       ' If Dir("\\ant\dept-as\HYD11\Localization\Exclusions\1_Srinivas\AEW\QC\QC Assigning_automated\UK\") <> "" Then

        Call GetMyFile_QC
        GoTo final
        Else
        MsgBox ("There are no automated files available to download!")
        ThisWorkbook.Save
        End If
    End If
End With

Next

'calculates time difference between recent ops upload file and qc upload file with current time
curtime = TimeValue(Now)
curdate = DateValue(Now)

opsdiffd = curdate - rec_up_date_OPs
opsdifft = Minute(curtime - rec_up_time_OPs)
qcdiffd = (curdate - rec_up_date_QC)
qcdifft = Minute(curtime - rec_up_time_QC)


If opsdiffd = qcdiffd Then

    If rec_up_time_OPs > rec_up_time_QC Then

        recfile = "OPs"
        Else
        recfile = "QC"
    End If



ElseIf opsdiffd > qcdiffd Then
recfile = "QC"
Else: recfile = "OPs"
End If


If recfile = "QC" Then
'download ops file for the associate
Call GetMyFile_Ops
GoTo final
End If

If recfile = "OPs" Then
'download QC file for the associate
Call GetMyFile_QC
GoTo final
End If

final:

    Application.AskToUpdateLinks = True
    Application.DisplayAlerts = True

End Sub
