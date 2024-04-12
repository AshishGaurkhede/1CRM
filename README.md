# 1CRM'**************************************** [ Variable Declaration ] **************************************
Dim vStepName
Dim vStepDetail
Dim vStatus
Dim vPOLength
Dim vInvoiceLength

'Global Constant
Const cnstIEPATH = "c:\Program Files\Internet Explorer\iexplore.exe"

'Clean up Switch needs to move in the Environment file 
Const RUNCLENUP = True

'Exist or Wait for Constants
Const intCnstWaitLevel0=1
Const intCnstWaitLevel1=2
Const intCnstWaitLevel2=5
Const intCnstWaitLevel3=10
Const intCnstWaitLevel4=3

' ******************************************[ Functions ] *********************************************

'*********************Send Enter Key**************************
 Function PressEnter()
    Set WshShell = CreateObject("WScript.Shell")
    Wait(1)
    WshShell.SendKeys "{ENTER}"
    Set WshShell = Nothing
End Function

  Function fLaunch
   vApplication = datatable.Value("Application", "Global")
   vUrl = datatable.Value("PageUrl", "Global")
   SystemUtil.Run vApplication, vUrl
  End Function
  
  '******************Folder Management**************************
    Function fFolder()
    ssName = "UFT"
    fPath ="C:\Temp"&"\"&ssName
	Set obj = Createobject("Scripting.FileSystemObject")
	If obj.FolderExists(fPath) = True Then
	 obj.DeleteFolder(fPath)
	 Else 
	 obj.CreateFolder fPath
	End If 
	 obj.CreateFolder fPath
	Set obj = nothing
 End Function
 
  '*************************** Reporter ****************************
Function fReporter (vStatus, vStepName, vStepDetail)
  CurrentDate = Date
  CurrentTime = "~"& Time
  CD = Replace(CurrentDate, "/", "_")
  CT = Replace(CurrentTime, ":", "_")
  CurrentDT = CD + CT
  ssName = "UFT"
  fPath ="C:\Temp"&"\"&ssName
  imgpath = fPath &"\"&"ScreenSnap-" & CurrentDT&".png"
  Desktop.CaptureBitmap imgpath, True
  Reporter.ReportEvent vStatus, vStepName, vStepDetail  , imgpath
  End Function
  
  '*************************** Invoice Extractor ****************************
Function fInvoice_Extractor(vInvoice)
	viewposition = InStr(vInvoice, "—"+" El Paso Electric")
	Invoice_Number = Mid(vInvoice,1,viewposition-2)
	DataTable.Value("Invoice_Id","Global")= Invoice_Number
End Function

'************************* PDF Path Validator *****************

Function fPath_Validator(vPDFpath)

	Set fso = CreateObject("Scripting.FileSystemObject")
	
   	If (fso.FileExists(vPDFpath)) Then
      		Call fReporter (micpass, "Validate Invoice PDF Path ", "Invoice PDF Path validated")
  	Else
      		Call fReporter (micpass, "Validate Invoice PDF Path", "Invoice PDF Path Not Found")
      		ExitTest
  	End If
 	Set obj = nothing
 	
 End Function

'##################################[ 1 ]#############################################################################
'# Function Name: fUtility_ReportPass
'# Description: It  Displays Pass Steps Report with ScreenShot
'# Parameters: 2
'# Return Value: String (Capture Screenshot filepath)
'# strStatus = fUtility_ReportPass("login", Login Successfully)
'####################################################################################################################

Function fUtility_ReportPass(p_strStepname,p_strStepdetails)
        'On Error Resume Next
'       Desktop.CaptureBitmap "TestPassed.bmp",True
'       Reporter.ReportEvent micPass,p_strStepname,p_strStepdetails
        fileName = fUtility_TimeStamp()
        Reporter.ReportEvent micPass, p_strStepname,p_strStepdetails, Reporter.ReportPath & "\" & fileName &".png"
        fUtility_ReportPass = Reporter.ReportPath & "\" & fileName &".png"
       'On Error goto 0
End Function

'##################################[ 2 ]###################################################################################
'# Function Name: fUtility_ReportFail
'# Description: It  Displays Fail Steps Report  with ScreenShot
'# Parameters: 2
'# Return Value: String (Capture Screenshot filepath)
'#
'# strStatus = fUtility_ReportFail("login", "Login Failed")
'#
'#####################################################################################################################

    Function fUtility_ReportFail(p_strStepname,p_strStepdetails)
       'On Error Resume Next
        fileName = fUtility_TimeStamp()
        Reporter.ReportEvent micFail,p_strStepname,p_strStepdetails, Reporter.ReportPath & "\" & fileName &".png"
        fUtility_ReportFail = Reporter.ReportPath & "\" & fileName &".png"
      ' On Error goto 0
  End Function

'##################################[ 3 ]###############################################################################
'# Function Name: fUtility_ReportDone
'# Description: It  Displays Done Steps Report  with ScreenShot
'# Parameters: 2
'# Return Value: String (Capture Screenshot filepath)
'#
'# strStatus = fUtility_ReportDone("login", "Login Done")
'#
'#####################################################################################################################

    Function fUtility_ReportDone(p_strStepname,p_strStepdetails)
       'On Error Resume Next
        fileName = fUtility_TimeStamp()
        Reporter.ReportEvent micDone,p_strStepname,p_strStepdetails, Reporter.ReportPath & "\" & fileName &".png"
        fUtility_ReportDone = Reporter.ReportPath & "\" & fileName &".png"
      ' On Error goto 0
  End Function

'###################################[ 4 ]#############################################################################
'# Function Name: fUtility_ReportWarning
'# Description: It  Displays Done Steps Report  with ScreenShot
'# Parameters: 2
'# Return Value: String (Capture Screenshot filepath)
'# strStatus = fUtility_ReportWarning("login", "Login Success with a warning")
'#
'#####################################################################################################################

    Function fUtility_ReportWarning(p_strStepname,p_strStepdetails)
       'On Error Resume Next
        fileName = fUtility_TimeStamp()
        Reporter.ReportEvent micWarning,p_strStepname,p_strStepdetails, Reporter.ReportPath & "\" & fileName &".png"
        fUtility_ReportWarning = Reporter.ReportPath & "\" & fileName &".png"
      ' On Error goto 0
  End Function
  
'#####################################[ 5 ]############################################################################
'# Function Name: fUtility_TimeStamp
'# Description: It is used to generate unique numbers by timestamp
'# Parameters: 
'# Return Value: N/A
'#
'#####################################################################################################################
Function fUtility_TimeStamp
    
    VarTimeStamp=Now
    VarTimeStamp =Replace(VarTimeStamp,":","")
    VarTimeStamp =Replace(VarTimeStamp,"/","")
    VarTimeStamp = Trim(VarTimeStamp)
    Desktop.CaptureBitmap Reporter.ReportPath & "\" & VarTimeStamp &".png",True
    fUtility_TimeStamp = VarTimeStamp
End Function
'#####################################################################################################################
Function fGetErrorMessage(StepDescription, ExpResult)
	Set objMsg = Description.Create
 
	 objMsg("class").Value = "errorMessage"
	 objMsg("html tag").Value = "SPAN"
	 
	 Set objPageMsg = Browser("Aqua Finance - Create").Page("Aqua Finance - Create").ChildObjects(objMsg)
	 
	 If objPageMsg.Count = 0 Then
	 	Exit Function
	 End If
	 
	 vErrorMsg = objPageMsg(0).GetROProperty("innertext")
	 'fGetErrorMessage = vErrorMsg
	  
	 If vErrorMsg <> "" Then
	 	Call fUtility_ReportFail ("Error Message: ", vErrorMsg)
	 	fExcel_InsertReport vExlPath, "Error", StepDescription, ExpResult,"Error Message : ", vErrorMsg, "Fail"
		ExitTestIteration
		'ExitActionIteration
	End If
	
End Function
'#####################################################################################################################

Function fErrorMessage(ByRef objField, ByRef vExlPath, ByRef StepDescription, ByRef ExpResult)

vMsg = objField.GetROProperty("class")

If Instr(vMsg,"hasError")>0 Then
	Set objMsg = Description.Create
 
	 objMsg("class").Value = "errorMessage"
	 objMsg("html tag").Value = "SPAN"
	 
	 Set objPageMsg = Browser("Aqua Finance - Create").Page("Aqua Finance - Create").ChildObjects(objMsg)
	 
	 If objPageMsg.Count = 0 Then
	 	Exit Function
	 End If
	 
	 vErrorMsg = objPageMsg(0).GetROProperty("innertext")
	  
	 If vErrorMsg <> "" Then
	 	Call fUtility_ReportFail ("Error Message : ", vErrorMsg)
	 	fExcel_InsertReport vExlPath, "Error", StepDescription, ExpResult,"Error Message : "& vErrorMsg, "Fail"
	 	Browser("Aqua Finance - Create").Page("Aqua Finance - Create").WebButton("Cancel").Click
		If Browser("Aqua Finance - Create").Page("Aqua Finance - Create").WebButton("Yes, Exit").Exist(2) Then
			Browser("Aqua Finance - Create").Page("Aqua Finance - Create").WebButton("Yes, Exit").Click
		End If
		
		'ExitActionIteration
		ExitTestIteration
	End If

End If

End Function

'#####################################[ 8 ]############################################################################
'# Function Name: OptionExistsInDropDown
'# Description: It is used to find options in dropdown list
'# Parameters: 2
'# Return Value: Boolean
'#####################################################################################################################

Function fOptionExistsInDropDown(ByRef objDropDown, ByRef strOption)
    Dim blnExist, arrDropdownOption, i
    blnExist = 0
    
    'arrDropdownOption = GetDropdownOptions(objDropDown)
    arrDropdownOption = Split(objDropDown.GetROProperty("all items"),";")
     
    For i = LBound(arrDropdownOption) To UBound(arrDropdownOption) Step 1
         If lcase(strOption) = lcase(arrDropdownOption(i)) Then
             strOption=arrDropdownOption(i)
             blnExist = 1
             Exit For
         End If
     Next
     
     If blnExist = 0 Then
         If UBound(arrDropdownOption) > 0 Then
             Reporter.ReportEvent micWarning, "Invalid DropDown Selection", "Value not found " & vbcrlf &  _ 
                chr(39) & strOption & chr(39) & vbcrlf & _
                "available options in the dropdown" & vbcrlf & _
                Join(arrDropdownOption, vbcrlf)
        Else
            Reporter.ReportEvent micWarning, "Invalid DropDown Selection", "Value not found " & vbcrlf &  _ 
                    chr(39) & strOption & chr(39) & vbcrlf & _
                    "Empty dropdown list"
         End If         
     End If    
    fOptionExistsInDropDown = blnExist
End Function

'#####################################[ 9 ]############################################################################
'# Function Name: waitForObject
'# Description: It Wait for Objects to appear
'# Parameters: 1
'# Return Value: Boolean
'#####################################################################################################################

Function fwaitForObject(ByRef myObject)
    Dim counter
    counter = 1
    Do
        Wait 1
        counter = counter + 1
        If counter = 30 Then
            fwaitForObject = False
            Exit Do
        End If
    'Loop While myObject.Exist
    Loop until myObject.Exist
    fwaitForObject = True
End Function
    
Function fExcel_CreateExcelReport(ByVal strExlPath)

    '****************Variable declaration***************

    Dim strDate     'Date
    Dim strExlReportPath    'Excel Sheet Path
    Dim blnFlag        'Boolean variable as a flag
    Dim str_Fields        'Array contains All fields to be in the Report
    Dim objFS    'File system object
    Dim obj_Exl    'Excel object
    Dim obj_sheet    'Excel Sheet
    Dim i    'Iterator
   
    '******Defining values***********************
   
    strDate = Month(Now) & "\" & Day(Now) & "\" & Year(Now)
    str_Fields = Array("Step_No.","Step_Description","Date","TimeStamp","Expected_Result","Actual_Result","Status")

    strExlReportPath = strExlPath

    fExcel_CreateExcelReport = True

    If Len(strExlReportPath) > 0 Then
         Set obj_Exl = createobject("excel.application")
         obj_Exl.Application.Visible = False
         obj_Exl.Workbooks.Add
         obj_Exl.DisplayAlerts = false
         'wait 2
         obj_Exl.ActiveWorkbook.SaveAs(strExlReportPath)
         
         Environment("oExcel") = obj_Exl 
         Set oExcelWB = obj_Exl.ActiveWorkbook
         Environment("oExcelWB") = oExcelWB
         Set obj_sheet = obj_Exl.ActiveWorkbook.Worksheets.Add
        'obj_sheet.Name = "Application"&Environment("TestIteration") 
	        
         obj_sheet.Name = "TC"&Environment("TestIteration") &"_"&DataTable.Value("Test Case ID","Global")
	'obj_sheet.Name = "TC"&Environment("TestIteration") &"_"&"PO Creation"


         Set obj_sheet = obj_Exl.ActiveWorkbook.Worksheets(obj_sheet.Name )
         
               obj_sheet.cells(1,1).value = "Test Name : "&Environment("TestName")
               obj_sheet.Range("A1","G1").Merge
               obj_sheet.cells(2,1).value = "Test Case ID : "&DataTable.Value("Test Case ID","Global")
               obj_sheet.Range("A2","G2").Merge
               obj_sheet.cells(3,1).value = "Test Description : "&DataTable("p_TestDescription", dtGlobalSheet)
               obj_sheet.Range("A3","G3").Merge
               obj_sheet.cells(4,1).value = "Host Name : "&Environment("LocalHostName")
               obj_sheet.Range("A4","G4").Merge
'               obj_sheet.cells(5,1).value = "Application URL : "&Datatable("p_AppURL", "Main")
		obj_sheet.cells(5,1).value = "Application URL : "& "https://demo.1crmcloud.com/login.php?login_module=Home&login_action=index"
               obj_sheet.Range("A5","G5").Merge
               obj_sheet.cells(6,1).value = "User Name : "&Ucase(Environment("UserName"))
               obj_sheet.Range("A6","G6").Merge
'               obj_sheet.cells(7,1).value = "Test Results Path : "&Datatable("p_ExlPath", "Main")
		obj_sheet.cells(7,1).value = "Test Results Path : "& "C:\Report"
               obj_sheet.Range("A7","G7").Merge
               obj_sheet.cells(8,1).value = "Execution Date : "&now
               obj_sheet.Range("A8","G8").Merge
               obj_sheet.Range("A1","G8").Borders.LineStyle = 1
               obj_sheet.Range("A1","G8").Font.Bold = True
       
         RowCount = obj_sheet.UsedRange.Rows.count
         RowCount = RowCount + 2

          For i = 0 To Ubound(str_Fields)
           obj_sheet.cells(RowCount,i+1).value = str_Fields(i)
           obj_sheet.Cells(RowCount, i+1).Borders.LineStyle = 1
           obj_sheet.Cells(RowCount, i+1).HorizontalAlignment =  -4108
           obj_sheet.Cells(RowCount, i+1).Interior.ColorIndex = 19
           obj_sheet.Cells(RowCount, i+1).Font.Bold = True
          Next
          
		obj_sheet.cells(1,1).ColumnWidth = 9
	       	obj_sheet.cells(1,2).ColumnWidth = 35
	       	obj_sheet.cells(1,3).ColumnWidth = 10
	       	obj_sheet.cells(1,4).ColumnWidth = 12
	       	obj_sheet.cells(1,5).ColumnWidth = 40
		obj_sheet.cells(1,6).ColumnWidth = 55
		obj_sheet.cells(1,7).ColumnWidth = 10
		
		'Set oSheet2 = oExcelWB.Worksheets("Sheet1")
		Set oSheet2 = obj_Exl.ActiveWorkbook.Worksheets.Add
		oSheet2.name = "Report Summary"
		oSheet2.Move oExcelWB.Worksheets(1)
		Environment("StatusSheet") =  oSheet2
		oSheet2.cells(1,1).value = "Test Case No."
		oSheet2.cells(1,1).ColumnWidth = 12
		oSheet2.cells(1,2).value = "Test Case ID"
		oSheet2.cells(1,2).ColumnWidth = 30
		oSheet2.cells(1,3).value = "Test Description"
		oSheet2.cells(1,3).ColumnWidth = 120
		oSheet2.cells(1,4).value = "Status"
		oSheet2.cells(1,4).ColumnWidth = 12
		
		oSheet2.Range("A1","D1").Borders.LineStyle = 1
	       oSheet2.Range("A1","D1").HorizontalAlignment =  -4108
	       oSheet2.Range("A1","D1").Interior.ColorIndex = 19
	       oSheet2.Range("A1","D1").Font.Bold = True
	       
		RowCount2 = oSheet2.UsedRange.Rows.count
	       RowCount2 = RowCount2 + 1
	       oSheet2.cells(RowCount2,1).value = "TC "&Environment("TestIteration")
		oSheet2.cells(RowCount2,2).value = DataTable.Value("Test Case ID",dtGlobalSheet)
		oSheet2.cells(RowCount2,3).value = DataTable("p_TestDescription", dtGlobalSheet)
		oExcelWB.Save
	'         obj_sheet.Rows.Autofit
	'         'obj_sheet.Columns.Autofit
	'         obj_Exl.ActiveWorkbook.Save
	'         obj_Exl.ActiveWorkbook.Close
	'         obj_Exl.Quit
	'         Set obj_Exl = nothing
    Else
        fExcel_CreateExcelReport = False
        Exit Function   
    End IF
End Function

'#####################################################################################################################

Function fGetExlReportPath(Byval vExlReportPath)

'For Excel Reporting-----------------------------------------
    'Dim vExlReportPath
    Dim vFoldername , vSubFolder
    Dim vFileName
    Dim vExlPath
    Dim objFSO
   
    Set objFSO = CreateObject("Scripting.FileSystemObject")
   
    If Len(vExlReportPath)=0 Then
        vExlReportPath = Environment.Value("SystemTempDir")
    End if
    'creating date wise folder 
    vFoldername = vExlReportPath & "\Report_" & Replace(Cstr(Date),"/","-")
   
    'Folder to be created If the folder doenot exst then create the folder
   
    If objFSO.FolderExists(vFoldername) = false Then
        objFSO.CreateFolder (vFoldername)
    End If
   
'    vSubFolder   = vFoldername&"\"&Environment("TestName")
'    If objFSO.FolderExists(vSubFolder) = false Then
'         objFSO.CreateFolder(vSubFolder) ' creating Sub folder
'    End If
   
'    vFileName   = "ORB-Stage-"&Environment("TestName")& "_" & Replace(Cstr(Time),":","-") &"_Report.xls"
    vFileName   = Environment("TestName")& "_" & Replace(Cstr(Time),":","-") &"_Report.xlsx"
    
    vExlPath    = vFoldername &"\"& vFileName
    fGetExlReportPath = vExlPath   
    
End Function

'#####################################################################################################################
    
Function fExcel_InsertReport(ByVal str_ExcelPath,ByVal StepNo,ByVal StepDesc,ByVal ExpResult,ByVal ActResult,ByVal Status)
    Dim str_TestName
    Dim str_Date
    Dim str_Time
    Dim RowCount
    Dim obj_Exl
    Dim obj_sheet

    str_Date = Month(Now) & "/" & Day(Now) & "/" & Year(Now)
    str_Time = Time
   
    fExcel_InsertReport = True
   
'    If Len(str_ExcelPath) > 0 Then
'        Set obj_Exl = createobject("exc   el.application")
'        obj_Exl.DisplayAlerts = false
'        obj_Exl.Workbooks.Open str_ExcelPath
'        obj_Exl.Application.Visible = False
     If IsObject(str_ExcelPath) Then
       		
     	obj_sheetName = "TC"&Environment("TestIteration") &"_"&DataTable.Value("Test Case ID","Global")
'	 obj_sheetName = "TC"&Environment("TestIteration") &"_"&"1CRM"
	 
        'set obj_sheet = obj_Exl.ActiveWorkbook.Worksheets("Application"&Environment("TestItertaion"))
        Set objWB = str_ExcelPath
        set obj_sheet = objWB.Worksheets(obj_sheetName)
        RowCount = obj_sheet.UsedRange.Rows.count
        RowCount = RowCount + 1
       
        Set objRange = obj_sheet.Range("A"&RowCount,"G"&RowCount)
        objRange.WrapText = True
    
        obj_sheet.cells(RowCount,1).value = StepNo & RowCount-10
        'obj_sheet.Range("A"&RowCount).ColumnWidth = 9
       
        obj_sheet.cells(RowCount,2).value = StepDesc
        'obj_sheet.Range("B"&RowCount).ColumnWidth =35
       
        obj_sheet.cells(RowCount,3).value = str_Date       
        obj_sheet.cells(RowCount,4).value = str_Time
        'obj_sheet.Range("C"&RowCount,"D"&RowCount).ColumnWidth =14

        obj_sheet.cells(RowCount,5).value = ExpResult  
	'obj_sheet.Range("E"&RowCount).ColumnWidth =40        
        obj_sheet.cells(RowCount,6).value = ActResult
        'obj_sheet.Range("F"&RowCount).ColumnWidth =55
       
        obj_sheet.cells(RowCount,7).value = Status
        'obj_sheet.Range("G"&RowCount).ColumnWidth =10
       
       Set oSheet2 = Environment("StatusSheet") 
	RowCount2 = oSheet2.UsedRange.Rows.count
       'RowCount2 = RowCount2 + 1
       oSheet2.cells(RowCount2,4).value = Status
       
        If strComp(Ucase(Status),"FAIL") = 0 Then
            'objRange.Interior.ColorIndex = 3
            obj_sheet.Cells(RowCount,7).Interior.ColorIndex = 3
            oSheet2.cells(RowCount2,4).Interior.ColorIndex = 3
        End IF
       
        If strComp(Ucase(Status),"PASS") = 0  Then
            'objRange.Interior.ColorIndex = 43
            obj_sheet.Cells(RowCount,7).Interior.ColorIndex = 43
            oSheet2.cells(RowCount2,4).Interior.ColorIndex = 43
        End If
           
'        If strComp(Ucase(Status),"WARNING") = 0 Then
'            'objRange.Interior.ColorIndex = 43
'            obj_sheet.Cells(RowCount,7).Interior.ColorIndex = 46
'        End If
       
        objRange.Borders.LineStyle = 1
        objRange.HorizontalAlignment = -4131
        objRange.VerticalAlignment = -4108
       
       oSheet2.Range("A"&RowCount2,"D"&RowCount2).Borders.LineStyle = 1
       oSheet2.Range("A"&RowCount2,"D"&RowCount2).HorizontalAlignment = -4131
       oSheet2.Range("A"&RowCount2,"D"&RowCount2).VerticalAlignment = -4108
       
 	objWB.Save              
'        obj_sheet.ActiveWorkbook.Save
'        obj_sheet.ActiveWorkbook.Close
'        obj_sheet.Quit
'        
'	  Set obj_sheet = nothing        
'        Set obj_sheet = nothing
    Else
        fExcel_InsertReport = False
        Exit Function
    End IF

End Function

'#####################################################################################################################

Function fExcel_AddSheet(str_ExcelPath)

	'****************Variable declaration***************
    Dim strDate     'Date
    Dim strExlReportPath    'Excel Sheet Path
    Dim blnFlag        'Boolean variable as flag
    Dim str_Fields        'Array contains All fields to be in the Report
    Dim objFS    'File system object
    Dim obj_Exl    'Excel object
    Dim obj_sheet    'Excel Sheet
    Dim i    'Iterator
   
    '******Defining values***********************
   
    strDate = Month(Now) & "\" & Day(Now) & "\" & Year(Now)
    str_Fields = Array("Step_No.","Step_Description","Date","TimeStamp","Expected_Result","Actual_Result","Status")

    'strExlReportPath = str_ExcelPath
    'fExcel_CreateExcelReport = True

'    If Len(strExlReportPath) > 0 Then
'  	Set obj_Exl = createobject("excel.application")
'       obj_Exl.Application.Visible = False
'         
'       Set obj_WB = obj_Exl.Workbooks.Open(strExlReportPath)
       'Set obj_sheet = obj_WB.Worksheets.Add(obj_WB.Worksheets(obj_WB.Worksheets.Count))
       'obj_sheet.Name = "Application"&Environment("TestIteration") 
       If IsObject(str_ExcelPath) Then
	        Set objWB = str_ExcelPath
	        Set obj_sheet = objWB.Worksheets.Add
	        
	        obj_sheet.Name = "TC"&Environment("TestIteration") &"_"&DataTable.Value("Test Case ID","Global")
	        objWB.sheets(obj_sheet.Name).Move objWB.Worksheets(objWB.Worksheets.Count)
	       
	        Set obj_sheet = objWB.Worksheets(obj_sheet.Name)
               obj_sheet.cells(1,1).value = "Test Name : "&Environment("TestName")
               obj_sheet.Range("A1","G1").Merge
               obj_sheet.cells(2,1).value = "Test Case ID : "&DataTable.Value("Test Case ID","Global")
               obj_sheet.Range("A2","G2").Merge
               obj_sheet.cells(3,1).value = "Test Description : "&DataTable("p_TestDescription", dtGlobalSheet)
               obj_sheet.Range("A3","G3").Merge
               obj_sheet.cells(4,1).value = "Host Name : "&Environment("LocalHostName")
               obj_sheet.Range("A4","G4").Merge
               obj_sheet.cells(5,1).value = "Application URL : "&Datatable("p_AppURL", "Main")
               obj_sheet.Range("A5","G5").Merge
               obj_sheet.cells(6,1).value = "User Name : "&Ucase(Environment("UserName"))
               obj_sheet.Range("A6","G6").Merge
               obj_sheet.cells(7,1).value = "Test Results Path : "&Datatable("p_ExlPath", "Main")
               obj_sheet.Range("A7","G7").Merge
               obj_sheet.cells(8,1).value = "Execution Date : "&now
               obj_sheet.Range("A8","G8").Merge
               obj_sheet.Range("A1","G8").Borders.LineStyle = 1
               obj_sheet.Range("A1","G8").Font.Bold = True
             
	         RowCount = obj_sheet.UsedRange.Rows.count
	         RowCount = RowCount + 2
	
	          For i = 0 To Ubound(str_Fields)
	           obj_sheet.cells(RowCount,i+1).value = str_Fields(i)
	           obj_sheet.Cells(RowCount, i+1).Borders.LineStyle = 1
	           obj_sheet.Cells(RowCount, i+1).HorizontalAlignment =  -4108
	           obj_sheet.Cells(RowCount, i+1).Interior.ColorIndex = 19
	           obj_sheet.Cells(RowCount, i+1).Font.Bold = True
	          Next

		obj_sheet.cells(1,1).ColumnWidth = 9
	       	obj_sheet.cells(1,2).ColumnWidth = 35
	       	obj_sheet.cells(1,3).ColumnWidth = 10
	       	obj_sheet.cells(1,4).ColumnWidth = 12
	       	obj_sheet.cells(1,5).ColumnWidth = 40
		obj_sheet.cells(1,6).ColumnWidth = 55
		obj_sheet.cells(1,7).ColumnWidth = 10
		
		Set oSheet2 = Environment("StatusSheet") 
		'oSheet2.Move objWB.Worksheets(1)
		RowCount2 = oSheet2.UsedRange.Rows.count
	       RowCount2 = RowCount2 + 1
	       oSheet2.cells(RowCount2,1).value = "TC "&Environment("TestIteration")
		oSheet2.cells(RowCount2,2).value = DataTable.Value("Test Case ID",dtGlobalSheet)
		oSheet2.cells(RowCount2,3).value = DataTable("p_TestDescription", dtGlobalSheet)
		objWB.Save
		
         'obj_sheet.Rows.Autofit
         'obj_sheet.Columns.Autofit
         'obj_Exl.ActiveWorkbook.Close
         'obj_Exl.Quit

'	  Set obj_sheet = nothing    
'         Set obj_Exl = nothing
     
    End IF
End Function
'#####################################################################################################################
Function fCloseExcel

	Set obj_Exl = Environment("oExcel")
	Set objWB = Environment("oExcelWB")
	'Set oSheet2 = Environment("StatusSheet") 
	'oSheet2.Move objWB.Worksheets(1)
	objWB.Worksheets(1).Activate
	objWB.Close
	obj_Exl.Quit
	'Set oSheet2 = nothing
	Set objWB = nothing
	Set obj_Exl = nothing
	
End Function
'#####################################################################################################################

Function RandomString
'    Dim str
'    Const LETTERS = "abcdefghijklmnopqrstuvwxyz"
'    For i = 1 to strSize
'        str = str & Mid( LETTERS, RandomNumber( 1, Len( LETTERS ) ), 1 )
'    Next
 
    character = Array("a", "b", "c", "d", "e", "f", "g", "h", "i", "j")
    unique = fConstantLenght(Datepart("m",Date),1)&fConstantLenght(Datepart("d",Date),1)&Replace(FormatDateTime(Time,4),":","")&DatePart("s",now)
    'unique = fConstantLenght(Datepart("m",Date),1)&fConstantLenght(Datepart("d",Date),1)&fConstantLenght( Replace(Timer,".",""),6)
    vNewStr = ""
    For i = 1 to Len(unique) Step 1 'i is the counter variable and it is incremented by 1
            myStr = CInt(Mid(unique, i, 1))
            vNewStr = vNewStr & character(myStr)
    Next

    RandomString = vNewStr

End Function

'#####################################################################################################################
Function fConvertDate(vEnterDate)
	
	If vEnterDate <> "" Then
		Dim vDate
		Dim vMonth
		Dim vYear
		Dim vConverted_Date
		
		vDate = Datepart("d",vEnterDate)
		vMonth = Datepart("m",vEnterDate)
		vYear = Datepart("yyyy",vEnterDate)
		
		If len(vDate) = 1 then
			vDate = "0" & vDate
		Else
			vDate = vDate
		End If
		
		If len(vMonth) = 1 then
			vMonth = "0" & vMonth
		Else
			vMonth = vMonth
		End If
		
		vConverted_Date = vMonth & "/"&vDate & "/" & vYear
		
		fConvertDate = vConverted_Date
	End If
	
End Function
'#####################################################################################################################
Function fConstantLenght(vInputStr,vStrLenght)
'This function is used for add "0" to the string
	If Len(vInputStr) = Cint(vStrLenght) Then
		vInputStr = "0"&vInputStr
	End If
	fConstantLenght = vInputStr
End Function

Function fDigitLenght(vInputStr,vStrLenght)
	'This function is used to add digits to the string
	Do While Len(vInputStr) < Cint(vStrLenght)
		vMax = 9
		vMin = 0
		Randomize
		vInputStr = vInputStr + CStr(Int((vMax-vMin+1)*Rnd+vMin))
	Loop

	fDigitLenght = vInputStr
End Function
'#####################################################################################################################
Function SelectCheckbox(vCheckbox)
    Set objExcelWB = Environment("oExcelWB")
    Set ObjDesc= Description.Create
    ObjDesc("html tag").value="INPUT"
    ObjDesc("type").value="checkbox"
    ObjDesc("value").value= vCheckbox

    Set allCheckBox = Browser("Aqua Finance - Create").Page("Aqua Finance - Create").ChildObjects(ObjDesc)
    
    If allCheckBox.Count()=0 Then
'    	fUtility_ReportPass "Select the checkbox "&vCheckbox, "Selected the checkbox "&vCheckbox
'	fExcel_InsertReport objExcelWB, "Step ","Select the checkbox '"&vCheckbox&"'", "Should be able to select the checkbox '"&vCheckbox&"'","Checkbox '"&vCheckbox&"' not found", "Fail"
    Else
    	allCheckBox(0).Click
    End If
'    allCheckBox(0).Click

End Function
'#####################################################################################################################

Function fAllFieldErrorMessage( )

	Set objMsg = Description.Create
 
	 objMsg("class").Value = "errorMessage"
	 objMsg("html tag").Value = "SPAN"
	 
	 Set objPageMsg = Browser("Aqua Finance - Create").Page("Aqua Finance - Create").ChildObjects(objMsg)
	 myCount=objPageMsg.Count

	 Dim arrFieldError() 
	 ReDim arrFieldError(myCount-1)
	 ' arrFieldError() = Array(myCount)
	 For i = 0 To Cint(myCount-1) Step 1
	 		arrFieldError(i)=objPageMsg(i).GetROProperty("innertext")
			'fExcel_InsertReport vExlPath, "Step ","All Missing Fields", "All Missing Field details should be displayed","The Following Missing Field details are displayed : "&vFieldError(i), "Fail"
	 Next
	 
	 
	 If objPageMsg.Count = 0 Then
	 	Exit Function
	 End If
	 
'	 vErrorMsg = objPageMsg(0).GetROProperty("innertext")
	  
'	 If vErrorMsg <> "" Then
'	 	Call fUtility_ReportFail ("Error Message : ", vErrorMsg)
'	 	fExcel_InsertReport vExlPath, "Error", StepDescription, ExpResult,"Error Message : "& vErrorMsg, "Fail"
'	 	Browser("Aqua Finance - Create").Page("Aqua Finance - Create").WebButton("Cancel").Click
'		Wait 1
'		Browser("Aqua Finance - Create").Page("Aqua Finance - Create").WebButton("Yes, Exit").Click
'		
'		'ExitActionIteration
'		ExitTestIteration
'	End If
 'End If

	fAllFieldErrorMessage = arrFieldError
End Function

'*************************Screenshot function for the Report****************************

Function CaptureImage()
    Dim Date_Time
    Dim Myfile
    Date_Time=Now() 
    Myfile= Date_Time&".png"
    Myfile = Replace(Myfile,"/","-") 
    Myfile = Replace(Myfile,":","-")
    Myfile= "C:\Report\Screen shorts"&Myfile 
    Desktop.CaptureBitmap Myfile, True
End Function

'*******************Create Folder for screenshorts*******************
Function CreateFolderDemo
   Dim fso, f
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set f = fso.CreateFolder("C:\Report\Screen shorts\"+ Date)
   CreateFolderDemo = f.Path
End Function

'****************Close Chrome Browser******************

Function CloseChrome
	Dim objshell
	Set objshell=CreateObject("WScript.Shell")
	'objshell.Run "TASKKILL /F /IM "& ProgramName(chrome.exe)
	
	Set objshell=nothing
End Function

'#####################################################################################################################

Public Function CloseBrowser()
    Dim shell
    Set shell = CreateObject("WScript.Shell")
    shell.SendKeys "^w" 'New line
    set shell = Nothing
    Wait 1
End Function
'#####################################################################################################################


