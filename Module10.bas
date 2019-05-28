Attribute VB_Name = "Module10"
Sub GetMyFile_Ops()

       ' ThisWorkbook.Save

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

'Opens ops tracker
Dim trk As Workbook
'Set trk = Workbooks.Open("\\ant\dept-as\Hyd11\Localization\Exclusions\1_Srinivas\Asin Exclusions - New Workflow\OPS\Ops Tracker\" & dashname & "_ASIN Tracker.xlsm")
Set trk = Workbooks.Open("\\ant\dept-as\HYD11\Localization\Exclusions\1_Srinivas\Asin Exclusions - New Workflow\OPS\Ops Tracker\" & dashname & "_ASIN Tracker.xlsm")

Dim newlast As Long
newlast = trk.Sheets("Assign").Cells(trk.Sheets("Assign").Rows.count, 1).End(xlUp).Row

'defining language order for assigning
Dim langlast As Long
langlast = ThisWorkbook.Sheets("Language").Cells(ThisWorkbook.Sheets("Language").Rows.count, 1).End(xlUp).Row

Dim src_path As String
Dim dest_path As String
Dim path As String
Dim fldr As FileDialog
Dim srcfile As String

  Date_1 = Date
  dest_path = "C:\Users\" & name & "\Desktop\ASIN Uploads\"
  If dest_path = "" Then GoTo final
    
Dim now_day As String
now_day = UCase(Left(WeekdayName(Weekday(Now)), 3))

Dim assigned_t As Variant
Dim assigned_d As Variant

Dim x As String, d As String
Dim wk As Workbook

Dim lang() As String

For j = 2 To langlast

         If ThisWorkbook.Sheets("Language").Cells(j, 2) = name Then
         If ThisWorkbook.Sheets("Language").Cells(j, 3) = "UK" Then
         lang = Split("UK,DE,ES,FR,IT", ",")
        
        ThisWorkbook.Sheets("Language").Protect "SpecOps1104"
        w = Application.WorksheetFunction.RandBetween(1, 10)
  
            Set rs = Nothing
            Set rs = New ADODB.Recordset
            cnt_File_Ops_Uk = "select count(File_name) as f_cnt from asin_exclusion.Ops_Assgn_Man_UK where Transaction_Date is null and Login_ID is null;"
            rs.Open cnt_File_Ops_Uk, cn
            
            file_cnt_uk = rs![f_cnt]
            
            If file_cnt_uk > 0 Then
                        
                    assigned_t = TimeValue(Now)
                    assigned_d = DateValue(Now)
                    
                    
                    ' Getting the file name from the Table
                Set rs = Nothing
                Set rs = New ADODB.Recordset
                lock_tbl = "LOCK TABLE asin_exclusion.Ops_Assgn_Man_UK WRITE;"
                rs.Open lock_tbl, cn
              
                Set rs = Nothing
                Set rs = New ADODB.Recordset
                Get_fn = "select File_name as fn, ID from asin_exclusion.Ops_Assgn_Man_UK where Transaction_Date is null and Login_ID is null limit 1"
                rs.Open Get_fn, cn
              
                file_name = rs![fn]
                ID_1 = rs![ID]
              
                Set rs = Nothing
                Set rs = New ADODB.Recordset
                Update_ID = "update asin_exclusion.Ops_Assgn_Man_UK set Transaction_Date = STR_TO_DATE('" & Date_1 & "', '%m/%d/%Y'), Login_ID = '" & name & "' where ID ='" & ID_1 & "';"
                rs.Open Update_ID, cn
                                
                Set rs = Nothing
                Set rs = New ADODB.Recordset
                ulk_tbls = "UNLOCK TABLES;"
                rs.Open ulk_tbls, cn
                    
                        Application.EnableEvents = False
                        Application.Wait (Now() + TimeValue("00:00:0" & w))
                            
                        ' exctracting the file that we got above into the desktop of the requester
                         
                        'Move the file
                        str7ZipPath = "\\ant\dept-as\HYD11\Localization\Exclusions\Source\7z.exe"
                        src_zip = "\\ant\dept-as\HYD11\Localization\Exclusions\Source\Ops_Asgn_UK.7z"
                        strPassword = "Test1234"
                        'dest_fol = "C:\Users\" & name & "\Desktop\Download" - using dest_path which is defined above

                        File_to_Move = file_name

                        cmd_ex_1_file = str7ZipPath & " e " & src_zip & " -p" & strPassword & " -o""" & dest_path & """ """ & File_to_Move & """"
                        'Debug.Print cmd_ex_1_file
                        Shell cmd_ex_1_file, vbHide

                        
                        Application.Wait (Now() + TimeValue("00:00:15"))
                        Set wk = Workbooks.Open(filename:=dest_path & file_name, ReadOnly:=True, UpdateLinks:=False)

'
                        lastrow = wk.Sheets("Sheet1").Cells(wk.Sheets("Sheet1").Rows.count, 1).End(xlUp).Row
                        
                        lastcol = wk.Sheets("Sheet1").Cells(1, wk.Sheets("Sheet1").Columns.count).End(xlToLeft).Column
                    
                        Application.EnableEvents = True
                        
                        'Moving file to ops' local drive
                        filename = wk.FullName
                        fname = wk.name
                        'wk.Save
                        wk.Close False
                        'Name path & d As dest_path & d
                        GoTo finish

            Else:
            GoTo finish
            End If
        
        ElseIf ThisWorkbook.Sheets("Language").Cells(j, 3) = "ES" Then
        lang = Split("ES,UK,FR,IT,DE", ",")
        ElseIf ThisWorkbook.Sheets("Language").Cells(j, 3) = "FR" Then
        lang = Split("FR,UK,IT,DE,ES", ",")
        ElseIf ThisWorkbook.Sheets("Language").Cells(j, 3) = "IT" Then
        lang = Split("IT,UK,FR,ES,DE", ",")
        ElseIf ThisWorkbook.Sheets("Language").Cells(j, 3) = "DE" Then
        lang = Split("DE,UK,FR,IT,ES", ",")
        Else: End If
        Else: End If
        
Next j

ThisWorkbook.Sheets("Language").Protect "SpecOps1104"

'x = "*.xlsm"
'src_path = "\\ant\dept-as\Hyd11\Localization\Exclusions\1_Sachin\AEW\Ops\Assigning\"
'src_path = "\\ant\dept-as\Hyd11\Localization\Exclusions\1_Srinivas\AEW\Ops\Assigning\"
'd = Dir(src_path & lang(0) & "\" & x)
'path = src_path & lang(0) & "\"

 w = Application.WorksheetFunction.RandBetween(1, 10)

For mp = 0 To 4

    Mp_pref = lang(mp)

    
            Set rs = Nothing
            Set rs = New ADODB.Recordset
            cnt_File_Ops = "select count(File_name) as f_cnt from asin_exclusion.Ops_Assgn_Man_" & Mp_pref & " where Transaction_Date is null and Login_ID is null;"
            rs.Open cnt_File_Ops, cn
            
            file_cnt = rs![f_cnt]
            
            If file_cnt > 0 Then
                        
                    assigned_t = TimeValue(Now)
                    assigned_d = DateValue(Now)
            
                    ' Getting the file name from the Table
                Set rs = Nothing
                Set rs = New ADODB.Recordset
                lock_tbl = "LOCK TABLE asin_exclusion.Ops_Assgn_Man_" & Mp_pref & " WRITE;"
                rs.Open lock_tbl, cn
              
                Set rs = Nothing
                Set rs = New ADODB.Recordset
                Get_fn = "select File_name as fn, ID from asin_exclusion.Ops_Assgn_Man_" & Mp_pref & " where Transaction_Date is null and Login_ID is null limit 1"
                rs.Open Get_fn, cn
              
                file_name = rs![fn]
                ID_1 = rs![ID]
              
                Set rs = Nothing
                Set rs = New ADODB.Recordset
                Update_ID = "update asin_exclusion.Ops_Assgn_Man_" & Mp_pref & " set Transaction_Date = STR_TO_DATE('" & Date_1 & "', '%m/%d/%Y'), Login_ID = '" & name & "' where ID ='" & ID_1 & "';"
                rs.Open Update_ID, cn
                                
                Set rs = Nothing
                Set rs = New ADODB.Recordset
                ulk_tbls = "UNLOCK TABLES;"
                rs.Open ulk_tbls, cn
                    
                        Application.EnableEvents = False
                        Application.Wait (Now() + TimeValue("00:00:0" & w))
                            
                        ' exctracting the file that we got above into the desktop of the requester
                         
                        'Move the file
                        str7ZipPath = "\\ant\dept-as\HYD11\Localization\Exclusions\Source\7z.exe"
                        src_zip = "\\ant\dept-as\HYD11\Localization\Exclusions\Source\Ops_Asgn_" & Mp_pref & ".7z"
                        strPassword = "Test1234"
                        'dest_fol = "C:\Users\" & name & "\Desktop\Download" - using dest_path which is defined above

                        File_to_Move = file_name

                        cmd_ex_1_file = str7ZipPath & " e " & src_zip & " -p" & strPassword & " -o""" & dest_path & """ """ & File_to_Move & """"
                        'Debug.Print cmd_ex_1_file
                        Shell cmd_ex_1_file, vbHide

                        
                        
                        Application.Wait (Now() + TimeValue("00:00:15"))
                        Set wk = Workbooks.Open(filename:=dest_path & file_name, ReadOnly:=True, UpdateLinks:=False)

                       lastrow = wk.Sheets("Sheet1").Cells(wk.Sheets("Sheet1").Rows.count, 1).End(xlUp).Row
                        
                        lastcol = wk.Sheets("Sheet1").Cells(1, wk.Sheets("Sheet1").Columns.count).End(xlToLeft).Column
                        Application.EnableEvents = True
                        'Moving file to ops' local drive
                        filename = wk.FullName
                        fname = wk.name
                        'wk.Save
                        wk.Close False
                        'Name path & d As dest_path & d
                        GoTo finish
    
           End If

Next mp
 
finish:

'Message in case there are no files available
If lastrow = 0 Then
Call GetMyFile_QC
trk.Save
trk.Close
GoTo final
Else: End If


If newlast <> 0 Then
newlast = newlast + 1
Else: End If

trk.Sheets("Upload").Unprotect "Prod1104"
trk.Sheets("Assign").Unprotect "Prod1104"
trk.Sheets("File Record").Unprotect "Prod1104"
trk.Sheets("Sheet3").Unprotect "Prod1104"

trk.Sheets("Assign").Cells(newlast, 1) = name
'trk.Sheets("Assign").Cells(newlast, 2) = filename 'Right(filename, InStrRev(filename, "\Test\"))
trk.Sheets("Assign").Cells(newlast, 3) = fname 'Right(filename, InStrRev(filename, "\Test\"))
trk.Sheets("Assign").Cells(newlast, 4) = (lastrow - 1)
trk.Sheets("Assign").Cells(newlast, 5) = assigned_d
trk.Sheets("Assign").Cells(newlast, 6) = Format(assigned_t, "hh:mm:ss")
trk.Sheets("Assign").Cells(newlast, 7) = 0
trk.Sheets("Assign").Cells(newlast, 8) = "Assigned"


trk.Sheets("Upload").Protect "Prod1104"
trk.Sheets("Assign").Protect "Prod1104"
trk.Sheets("File Record").Protect "Prod1104"
trk.Sheets("Sheet3").Protect "Prod1104"

trk.Save
trk.Close

MsgBox "Congrats! You've got your new OPs file named " & file_name & " for the day!"

final:

Application.DisplayAlerts = True

End Sub
