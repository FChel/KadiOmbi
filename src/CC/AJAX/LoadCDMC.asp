<!-- #Include file=../../ADOVBS.inc -->
<!-- #include file="../CAPSFunctions.asp" -->
<%

'On Error Resume Next
'Description:	Loads the CDMC File from the UploadCDMC.asp page
'Author:		Michael Giacomin
'Date:			February 2021

    'If the session has expired then send the user back to the Default/login page
	If IsEmpty(Session("UserID")) Then Response.Redirect("../Timeout.asp")
	
'Instantiate Common Page Variables.
Dim objCon
Dim objCmd
Dim objCmd1
Dim objRS

Dim strDeleteCheck
Dim dteBatchDate

'Open Database Connection
Set objCon = Server.CreateObject("ADODB.Connection")
Set objCmd = Server.CreateObject("ADODB.Command")
Set objCmd1 = Server.CreateObject("ADODB.Command")
Set objRS = Server.CreateObject("ADODB.Recordset")

objCon.Open Session("DBConnection")

'If the local load has been clicked then call the procedure to load the network file rather than uploading it
If Request.QueryString("Action")="SaveFileLocal" Then
	
	Call StartLoad()

End If


Sub StartLoad()

'On Error Resume Next 
Dim objExcelCon
Dim strUploadPath

Dim errors
Dim lineNo
Dim strFileName
Dim lngFileID

Dim objStartFolder
Dim strFileNameDefault
Dim strUpdatedBy
Dim objFSO
Dim colFiles
Dim objFolder
Dim objFile

Dim strFileDateTime

errors = "" 
lineNo = 1    
strUpdatedBy = Session("UserID")


If Request.QueryString("Action")="SaveFileLocal" Then

		strDeleteCheck = Request.QueryString("Delete")
				
		Set objFSO = CreateObject("Scripting.FileSystemObject")

			'objStartFolder = "D:\Inetpub\WWWRoot\CAPS\ASPNew\Admin\CAPSAdmin\Attachments\CDMC"
			objStartFolder = GetSystemAdmin("ServerFilePath") & "\Admin\CAPSAdmin\Attachments\CDMC\"		

			Set objFolder = objFSO.GetFolder(objStartFolder)
			Set colFiles = objFolder.Files

			'Get the System Parameter for the fileName
			strFileNameDefault = GetSystemAdmin("CDMCCardlistFileName")
			
			If IsNull(strFileNameDefault) or strFileNameDefault = "" Then

				Response.write "<div class=""content-body""><section id=""basic-alerts""><div class=""alert alert-danger alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
						"<span aria-hidden=""true"">&times;</span></button>" & _
						"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
						"<span>NOT LOADED! There is no System Parameter for CDMC File Names (""CDMCCardlistFileName""). See System Admin.</span></div></div></div>"
					Exit Sub
			End If

			For Each objFile in colFiles			

				If Left(objFile.Name,11) = Trim(strFileNameDefault) Then
					strFileName = objFile.Name
					filePath = objStartFolder & "" & strFileName
				End If
				
			Next

			If IsNull(strFileName) or strFileName = "" Then

				Response.write "<div class=""content-body""><section id=""basic-alerts""><div class=""alert alert-danger alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
						"<span aria-hidden=""true"">&times;</span></button>" & _
						"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
						"<span>" & strFileNameDefault & " NOT LOADED! There is no CDMC File in the Server Folder to Load! Copy the 'CMS_AllData.csv' file to the location " & objStartFolder & "</span></div></div></div>"
					Exit Sub
			End If
			
'---Start Excel driver upload

			'Check to see if the same FileSeqNum for the same FileType has already been loaded
			lngFileID = GetFileLoadID("CDMC","",strFileName)
			
			If lngFileID = "" Then
				'Get the next fileID Number for the ANZCardlist File
				lngFileID = GetNextCDMCFileID
			Else
				'If the checkbox to overwrite is checked then load the data, otherwise do not load
				'If strDeleteCheck = "on" Then
				If lngFileID = 1 or strDeleteCheck = "on" Then
					'Delete any existing CS From Diners Records
					objCon.Execute "TRUNCATE TABLE tblCAPSCDMC"'"DELETE FROM tblCAPSCDMC WHERE [FileID] = " & lngFileID & ""
				Else
					Response.write "<div class=""content-body""><section id=""basic-alerts""><div class=""alert alert-danger alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
						"<span aria-hidden=""true"">&times;</span></button>" & _
						"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
						"<span>NOT LOADED! The CDMC File """ & strFileName & """ has already been loaded! <a href=""UploadCDMC.asp?Reload=1&FileSeqNum=" & lngFileID & """ Style=""font-weight:bold; color:white;""> Check the 'Overwrite Existing Batch' box and load again to overwrite...</a></span></div></div></div>"
					Exit Sub
				End If
			End If
			
			'Call the relevant procedure depending on whether the file is .xls or .txt
			If Right(filePath,3) = "csv" or Right(filePath,3) = "txt" then
			
				'After uploading, Read excel file
				Set objExcelCon = Server.CreateObject("ADODB.connection")     
				
				objExcelCon.Open "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" & objStartFolder & ";Extensions=asc,csv,tab,txt;ColNameHeader=Yes;"
				
				'objExcelCon.Open "DBQ=" & filePath & "; DRIVER={Microsoft Excel Driver (*.xls)};" 
				'objExcelCon.Open "Driver={Microsoft Excel Driver (*.xls)};DriverId=790;Dbq=" & filePath & ";DefaultDir=c:\Apps\CAPS2\ASP2\Admin\CAPSAdmin\Attachments;" 
		'		objExcelCon.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strUploadPath & "\" & ";Extended Properties=""text;HDR=Yes;FMT=Delimited"";"
				'objExcelCon.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & filePath & ";Extended Properties=""text;HDR=Yes;FMT=Delimited(~)"";"
				
				'-----This line below was working with a server with OLEDB drivers ****
				'objExcelCon.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strUploadPath & "\" & ";Extended Properties=""text;"";"
				
				
				
				'objExcelCon.Open strUploadPath & "\CAPSCSFromDiners.dsn"
				
				'Write the SQL Query 
				objRS.open "SELECT * FROM [" & strFileName & "]", objExcelCon  
		    
				ReadExcel		    
				
				'Check for errors
				If(errors<>"") Then
					'Print the errors and return
					Response.Write "<font face=arial size=1><b>File not uploaded. Please correct the following errors and load again</b> <br> "& errors &"</font><br>"         
				Else  
					'If the no errors found in the ReadExcel method, then start uploading the records in the database       	    	    
					UploadExcel strDeleteCheck,lngFileID, strFileName,filePath,strFileDateTime
					'UploadExcel Uploader.Form("chkDelete"),lngFileID, strFileName,filePath,strFileDateTime
					'Response.Write "<b>ANZ Cardlist File Sucessfully Uploaded!!<b><br>"
				End If
				
				'Close the recordset/connection 
				objRS.Close 
				objExcelCon.Close 
		    
				'Move the file to the Loaded folder
				Set objFSO = CreateObject("Scripting.FileSystemObject")
				'Set outPut = objFSO.CreateTextFile("c:\\output.txt", true);
				'Set objTextFile = objFSO.OpenTextFile (strFileNamePath, ForReading)

				objStartFolder = GetSystemAdmin("ServerFilePath") & "\Admin\CAPSAdmin\Attachments\CDMC\"		
				
				'Change the file name to add the date so it does not overwrite existing files in the same folder
				strFileName = Left(strFileName,Len(strFileName)-4) & Year(now()) & PadDigits(Month(now()),2) & Day(now()) & Hour(now()) & Minute(now()) & Right(strFileName,4)
				
				'Set objTextFile = Nothing
				'objFSO.MoveFile filePath,objStartFolder & "Loaded\" & strFileName 
		
					Response.write "<div class=""content-body""><section id=""basic-alerts""><div class=""alert alert-success alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
						"<span aria-hidden=""true"">&times;</span></button>" & _
						"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
						"<span>CDMC File """ & strFileName & """ load COMPLETE!</span></div></div></div>"

		
			Else
				'ANZ file is xls as provided by ANZ and  not a text file, so do not enable text load
				'ReadText(filePath)
			End If
			
	    'End If
	'End of Start load for Client side loading of files (which has been merged with server file load procedure at th etop of this process)
	'End IF

'---End New Excel Section	
End if

End Sub

'Function for validating the excel file values
Sub ReadExcel 

Dim errors
Dim strCardNumber
Dim strEID
Dim lineNo

    'First loop to check all the values
    Do until objRS.EOF 
        'Read each record present in the Excel file and check for the validation
 
        'strCardNumber = objRS("Card Number") 'Should be integer and Not Null values       
        'If IsNull(strFileSeqNum) Then
        '    errors = errors & "Error in line no. " & lineNo & ": Card Number should NOT be an Empty/Null value <br>" 
        'End if

		'response.write "fields=" & objRS(0) & "," '& objRS(1) & "," & objRS(2) & "," & objRS(3) & "," & objRS(4) & "," & objRS(5) & "," & objRS(6) & ","
		
		strEID = objRS("EmployeeNumber") 'Should be integer and Not Null values       
        If IsNull(strEID) Then
            errors = errors & "Error in line no. " & lineNo & ": Employee ID should NOT be an Empty/Null value <br>" 
        End if
       
	   'strSurname = objRS("Cardholder Name") 'Should be integer and Not Null values       
        'If IsNull(strSurname) Then
        '    errors = errors & "Error in line no. " & lineNo & ": Cardholder Name should NOT be an Empty/Null value <br>" 
        'End if
                          
            lineNo = lineNo + 1            
            
        objRS.movenext 
    
        If IsNull(objRS("EmployeeNumber")) AND IsNull(objRS("EmployeeNumber")) Then       
        
            Exit Sub

        End If
    
    Loop        
    
End Sub	

'Function for validating the excel file values
Sub UploadExcel(chkDelete,lngFileID, strFileName,strFilePath,strFileDateTime)

Dim x

Dim lngCDMCID
Dim strGroupName
Dim strDivisionName
Dim strBranchName
Dim strDepartmentName
Dim strDepartmentNumber
Dim strCostCentre
Dim strEmployeeID
Dim strEmployeeType
Dim strFirstname
Dim strSurname
Dim strTitle
Dim strEmail_Address
Dim strTelephoneNumber
Dim strMobileNumber
Dim strDateofBirth
Dim strGender
Dim strActualRankLvl
Dim strSite
Dim strUnit
Dim strReportsTo
Dim strDCD_PostalAddress
Dim straddressline1
Dim straddressline2
Dim straddressline3
Dim straddressline4
Dim straddressline5
Dim straddressline6
Dim strPostalAddress_Unit
Dim strPostalAddress_ClientLocation
Dim strPostalAddress_DeliveryLocation
Dim strPostalAddress_City
Dim strPostalAddress_State
Dim strPostalAddress_PostCode
Dim strPostalAddress_Country
Dim strDCDProtectedIdentity
Dim strIsValidPostal
Dim strOutAddr1
Dim strOutAddr2
Dim strOutAddr3
Dim strOutSuburb
Dim strOutState
Dim strOutPostCode
Dim strPostalMessage
Dim strhasChanged
Dim strDCD_WorkAddress
Dim strClientLocation
Dim strStreetAddress
Dim strCity
Dim strState
Dim strPostCode
Dim strFormalFirstName
Dim strFormalLastName
Dim strFormalMiddleName
Dim strOutTitle
Dim strOutDinersWorkPhone
Dim strOutDinersMobilePhone
Dim strOutANZPhone
Dim strOutDinersAddress1
Dim strOutDinersAddress2
Dim strRemoveCountdown
Dim strFirstUpdated
Dim strLastUpdated
Dim strActive
Dim strUpdatedBy
Dim strDateUpdated
Dim strFileID
Dim strLoaded
Dim strDateLoaded

Dim strFileSeqNum
Dim lngFileLoadID

Dim objStartFolder
Dim objFSO

    objRS.MoveFirst()
    
    Do until objRS.EOF 
        
		x = x + 1
		
		lngCDMCID = 0'objRS("CDMCID") 
		strGroupName = objRS("GroupName") 
		strDivisionName = objRS("DivisionName") 
		strBranchName = objRS("BranchName") 
		strDepartmentName = objRS("DepartmentName") 
		strDepartmentNumber = objRS("DepartmentNumber") 
		strCostCentre = objRS("CostCentre") 
		strEmployeeID = objRS("EmployeeNumber") 
		strEmployeeType = objRS("EmployeeType") 
		strFirstname = objRS("GivenName") 
		strSurname = objRS("Surname") 
		strTitle = objRS("Title") 
		strEmail_Address = objRS("Email_Address") 
		strTelephoneNumber = objRS("TelephoneNumber") 
		strMobileNumber = objRS("MobileNumber") 
		strDateofBirth = objRS("DateofBirth") 
		strGender = objRS("Gender") 
		strActualRankLvl = objRS("ActualRankLv1")
		strSite = objRS("Site")
		strUnit = objRS("Unit")
		strReportsTo = objRS("ReportsTo")
		strDCD_PostalAddress = objRS("DCD_PostalAddress")		
		straddressline1 = objRS("addressline1") 
		straddressline2 = objRS("addressline2") 
		straddressline3 = objRS("addressline3") 
		straddressline4 = objRS("addressline4") 
		straddressline5 = objRS("addressline5") 
		straddressline6 = objRS("addressline6") 
		strPostalAddress_Unit = objRS("PostalAddress_Unit") 
		strPostalAddress_ClientLocation = objRS("PostalAddress_ClientLocation") 
		strPostalAddress_DeliveryLocation = objRS("PostalAddress_DeliveryLocation") 
		strPostalAddress_City = objRS("PostalAddress_City") 
		strPostalAddress_State = objRS("PostalAddress_State") 
		strPostalAddress_PostCode = objRS("PostalAddress_PostCode") 
		strPostalAddress_Country = objRS("PostalAddress_Country") 
		strDCDProtectedIdentity = objRS("DCDProtectedIdentity") 
		strIsValidPostal = ""'""'objRS("IsValidPostal") 
		strOutAddr1 = ""'objRS("OutAddr1") 
		strOutAddr2 = ""'objRS("OutAddr2") 
		strOutAddr3 = ""'objRS("OutAddr3") 
		strOutSuburb = ""'objRS("OutSuburb") 
		strOutState = ""'objRS("OutState") 
		strOutPostCode = ""'objRS("OutPostCode") 
		strPostalMessage = ""'objRS("PostalMessage") 
		strhasChanged = "N"'objRS("hasChanged") 
		strDCD_WorkAddress = objRS("DCD_WorkAddress") 
		strClientLocation = objRS("ClientLocation") 
		strStreetAddress = objRS("StreetAddress") 
		strCity = objRS("City") 
		strState = objRS("State") 
		strPostCode = objRS("PostCode") 
		strFormalFirstName = objRS("FirstName") 
		strFormalLastName = objRS("LastName") 
		strFormalMiddleName = objRS("MiddleName") 
		strOutTitle = ""'objRS("OutTitle") 
		strOutDinersWorkPhone = ""'objRS("OutDinersWorkPhone") 
		strOutDinersMobilePhone = ""'objRS("OutDinersMobilePhone") 
		strOutANZPhone = ""'objRS("OutANZPhone") 
		strOutDinersAddress1 = ""'objRS("OutDinersAddress1") 
		strOutDinersAddress2 = ""'objRS("OutDinersAddress2") 
		strRemoveCountdown = 0'objRS("RemoveCountdown") 
		strFirstUpdated = "NULL"'objRS("FirstUpdated") 
		strLastUpdated = "NULL"'objRS("LastUpdated") 
		strActive = ""'objRS("Active") 
		strDateLoaded = Now()'objRS("DateLoaded") 
		strUpdatedBy = Session("UserID")'objRS("UpdatedBy") 
		strFileID = lngFileID'objRS("FileID") 
		strLoaded = "N"'objRS("Loaded") 


        'The first 2 rows of the ANZ Cardlist Excel file has an image (ANZ logo) and header which are to be ignored
		'If x < 3 Then
			
			
		'Else
            
            'Call the procedure to save the record to SQL
			
			'Response.Write  "exec spGeneralExpensesSave ="& strCSFromDinersID & "," & strFileDateTime & "," & strFileSeqNum & "," & x
			
			'response.write "</br>" & lngCDMCID & "," & strGroupName & "," & strDivisionName & "," & strBranchName & "," & strEmployeeID & "," & strEmployeeType
			
			'response.write "</br>" & lngCDMCID & "," & strEmployeeID & "," & strDateofBirth
			
            SaveRecord lngCDMCID,strGroupName,strDivisionName,strBranchName,strDepartmentName,strDepartmentNumber,strCostCentre,strEmployeeID,strEmployeeType, _
					strFirstname,strSurname,strTitle,strEmail_Address,strTelephoneNumber,strMobileNumber,strDateofBirth,strGender,strActualRankLvl,strSite,strUnit,strReportsTo,strDCD_PostalAddress, _
					straddressline1,straddressline2,straddressline3,straddressline4,straddressline5,straddressline6,strPostalAddress_Unit,strPostalAddress_ClientLocation, _
					strPostalAddress_DeliveryLocation,strPostalAddress_City,strPostalAddress_State,strPostalAddress_PostCode,strPostalAddress_Country,strDCDProtectedIdentity, _
					strIsValidPostal,strOutAddr1,strOutAddr2,strOutAddr3,strOutSuburb,strOutState,strOutPostCode,strPostalMessage,strhasChanged,strDCD_WorkAddress,strClientLocation, _
					strStreetAddress,strCity,strState,strPostCode,strFormalFirstName,strFormalLastName,strFormalMiddleName,strOutTitle,strOutDinersWorkPhone,strOutDinersMobilePhone, _
					strOutANZPhone,strOutDinersAddress1,strOutDinersAddress2,strRemoveCountdown,strFirstUpdated,strLastUpdated,strActive,strUpdatedBy,strDateUpdated,strFileID,strLoaded, x
            
        
		'Session("CDMCLoadProgress") = x
		
        'End If
	If x > 1000 then exit sub
		objRS.movenext 
		
		If IsNull(strEmployeeID) Then               
		   
		   Exit Sub
		   
		End If
        
    Loop 
    
	If x > 1 Then
		
		'The fileDate and Number are only in the header row
		strFileDateTime = MediumDate(Now())
		'strFileSeqNum = Mid(strLine,11,12)
		
		strFileSeqNum = GetLastFileLoadID("CDMC",strFileName)
		'The FileID is passed in
		'lngFileID = GetFileLoadID("CDMC",strFileSeqNum,strFileName)
		
		'If the file has records in it then call the save File Load Procedure to save summary details (CAPSFunctions.asp)
'		lngFileLoadID = SaveFileLoadID ("CDMC",strFileName, strFilePath,x,0,0,0,0,0,0,0,strFileDateTime,lngFileID,"Imported",Session("UserID"),"N")
		
		'Call the procedure to update summary information for the file just loaded (CAPSFunctions.asp)
'		Call UpdateFileLoadSummary ("CDMC",lngFileID, strFileName, lngFileLoadID)
		'response.write "UpdateFileLoadSummary (""CDMCFile""," & lngFileID & "," & strFileName & "," & lngFileLoadID & ")"
		
'		Response.write "<div class=""content-body""><section id=""basic-alerts""><div class=""alert alert-success alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
'						"<span aria-hidden=""true"">&times;</span></button>" & _
'						"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
'						"<span>CDMC File """ & strFileName & """ load COMPLETE!</span></div></div></div>"
		
		'If the file has records in it then call the save File Load Procedure to save summary details (CAPSFunctions.asp)
		lngFileLoadID = SaveFileLoadID ("CDMC",strFileName,strFilePath,x,0,0,0,0,0,0,0,strFileDateTime,strFileSeqNum,"Imported",Session("UserID"),"N")
		'response.write "SaveFileLoadID (CDMC," & strFileName & "," & strFileNamePath & "," & x & "," & strFileDateTime & "," & strFileSeqNum
		'Call the procedure to update summary information for the file just loaded (CAPSFunctions.asp)
		Call UpdateFileLoadSummary ("CDMC",strFileSeqNum, strFileName,lngFileLoadID)
		'response.write "UpdateFileLoadSummary (""CSFRomDiners""," & strFileSeqNum & "," & lngFileLoadID & ")"
		
					
	End If
	
End Sub	


Sub SaveRecord(lngCDMCID,strGroupName,strDivisionName,strBranchName,strDepartmentName,strDepartmentNumber,strCostCentre,strEmployeeID,strEmployeeType, _
			strFirstname,strSurname,strTitle,strEmail_Address,strTelephoneNumber,strMobileNumber,strDateofBirth,strGender,strActualRankLvl,strSite,strUnit,strReportsTo,strDCD_PostalAddress, _
			straddressline1,straddressline2,straddressline3,straddressline4,straddressline5,straddressline6,strPostalAddress_Unit,strPostalAddress_ClientLocation, _
			strPostalAddress_DeliveryLocation,strPostalAddress_City,strPostalAddress_State,strPostalAddress_PostCode,strPostalAddress_Country,strDCDProtectedIdentity, _
			strIsValidPostal,strOutAddr1,strOutAddr2,strOutAddr3,strOutSuburb,strOutState,strOutPostCode,strPostalMessage,strhasChanged,strDCD_WorkAddress,strClientLocation, _
			strStreetAddress,strCity,strState,strPostCode,strFormalFirstName,strFormalLastName,strFormalMiddleName,strOutTitle,strOutDinersWorkPhone,strOutDinersMobilePhone, _
			strOutANZPhone,strOutDinersAddress1,strOutDinersAddress2,strRemoveCountdown,strFirstUpdated,strLastUpdated,strActive,strUpdatedBy,strDateUpdated,strFileID,strLoaded, x)

Dim intRecord

	'Make sure there are no long field names (where some data exists with long data)
	If IsNull(strFormalMiddleName) or strFormalMiddleName = "" Then
		strFormalMiddleName = ""
	Else
		If Len(strFormalMiddleName) > 50 Then strFormalMiddleName = Left(strFormalMiddleName,50)
	End If
	
  	With objCmd
  	
  	    'If the procedure has akready run then don't create the parameter objects again (more than once)
  	    If x = 1 then
			.CommandType = 4
			.CommandText = "spCAPSCDMCSave"
			
			.Parameters.Append objCmd.CreateParameter("CDMCID", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("GroupName", adVarWChar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("DivisionName", adVarWChar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("BranchName", adVarWChar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("DepartmentName", adVarWChar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("DepartmentNumber", adVarWChar, adParamInput,10)
			.Parameters.Append objCmd.CreateParameter("CostCentre", adVarWChar, adParamInput,10)
			.Parameters.Append objCmd.CreateParameter("EmployeeID", adVarWChar, adParamInput,20)
			.Parameters.Append objCmd.CreateParameter("EmployeeType", adVarWChar, adParamInput,30)
			.Parameters.Append objCmd.CreateParameter("Firstname", adVarWChar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("Surname", adVarWChar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("Title", adVarWChar, adParamInput,30)
			.Parameters.Append objCmd.CreateParameter("Email_Address", adVarWChar, adParamInput,100)
			.Parameters.Append objCmd.CreateParameter("TelephoneNumber", adVarWChar, adParamInput,30)
			.Parameters.Append objCmd.CreateParameter("MobileNumber", adVarWChar, adParamInput,30)
			.Parameters.Append objCmd.CreateParameter("DateofBirth", adVarWChar, adParamInput,10)
			.Parameters.Append objCmd.CreateParameter("Gender", adVarWChar, adParamInput,20)
			.Parameters.Append objCmd.CreateParameter("ActualRankLvl", adVarWChar, adParamInput,30)
			.Parameters.Append objCmd.CreateParameter("Site", adVarWChar, adParamInput,100)
			.Parameters.Append objCmd.CreateParameter("Unit", adVarWChar, adParamInput,200)
			.Parameters.Append objCmd.CreateParameter("ReportsTo", adVarWChar, adParamInput,100)
			.Parameters.Append objCmd.CreateParameter("DCD_PostalAddress", adVarWChar, adParamInput,500)
			.Parameters.Append objCmd.CreateParameter("addressline1", adVarWChar, adParamInput,100)
			.Parameters.Append objCmd.CreateParameter("addressline2", adVarWChar, adParamInput,100)
			.Parameters.Append objCmd.CreateParameter("addressline3", adVarWChar, adParamInput,100)
			.Parameters.Append objCmd.CreateParameter("addressline4", adVarWChar, adParamInput,100)
			.Parameters.Append objCmd.CreateParameter("addressline5", adVarWChar, adParamInput,100)
			.Parameters.Append objCmd.CreateParameter("addressline6", adVarWChar, adParamInput,100)
			.Parameters.Append objCmd.CreateParameter("PostalAddress_Unit", adVarWChar, adParamInput,255)
			.Parameters.Append objCmd.CreateParameter("PostalAddress_ClientLocation", adVarWChar, adParamInput,255)
			.Parameters.Append objCmd.CreateParameter("PostalAddress_DeliveryLocation", adVarWChar, adParamInput,255)
			.Parameters.Append objCmd.CreateParameter("PostalAddress_City", adVarWChar, adParamInput,255)
			.Parameters.Append objCmd.CreateParameter("PostalAddress_State", adVarWChar, adParamInput,255)
			.Parameters.Append objCmd.CreateParameter("PostalAddress_PostCode", adVarWChar, adParamInput,255)
			.Parameters.Append objCmd.CreateParameter("PostalAddress_Country", adVarWChar, adParamInput,255)
			.Parameters.Append objCmd.CreateParameter("ClientLocation", adVarWChar, adParamInput,100)
			.Parameters.Append objCmd.CreateParameter("StreetAddress", adVarWChar, adParamInput,100)
			.Parameters.Append objCmd.CreateParameter("City", adVarWChar, adParamInput,100)
			.Parameters.Append objCmd.CreateParameter("State", adVarWChar, adParamInput,100)
			.Parameters.Append objCmd.CreateParameter("PostCode", adVarWChar, adParamInput,20)
			.Parameters.Append objCmd.CreateParameter("DCDProtectedIdentity", adVarWChar, adParamInput,10)
			.Parameters.Append objCmd.CreateParameter("IsValidPostal", adVarWChar, adParamInput,3)
			.Parameters.Append objCmd.CreateParameter("OutAddr1", adVarWChar, adParamInput,40)
			.Parameters.Append objCmd.CreateParameter("OutAddr2", adVarWChar, adParamInput,40)
			.Parameters.Append objCmd.CreateParameter("OutAddr3", adVarWChar, adParamInput,40)
			.Parameters.Append objCmd.CreateParameter("OutSuburb", adVarWChar, adParamInput,22)
			.Parameters.Append objCmd.CreateParameter("OutState", adVarWChar, adParamInput,3)
			.Parameters.Append objCmd.CreateParameter("OutPostCode", adVarWChar, adParamInput,4)
			.Parameters.Append objCmd.CreateParameter("PostalMessage", adVarWChar, adParamInput,40)
			.Parameters.Append objCmd.CreateParameter("hasChanged", adVarWChar, adParamInput,3)
			.Parameters.Append objCmd.CreateParameter("FormalFirstName", adVarWChar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("FormalLastName", adVarWChar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("FormalMiddleName", adVarWChar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("OutTitle", adVarWChar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("OutDinersWorkPhone", adVarWChar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("OutDinersMobilePhone", adVarWChar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("OutANZPhone", adVarWChar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("OutDinersAddress1", adVarWChar, adParamInput,40)
			.Parameters.Append objCmd.CreateParameter("OutDinersAddress2", adVarWChar, adParamInput,40)
			.Parameters.Append objCmd.CreateParameter("RemoveCountdown", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("FirstUpdated", adDate, adParamInput)
			.Parameters.Append objCmd.CreateParameter("LastUpdated", adDate, adParamInput)
			.Parameters.Append objCmd.CreateParameter("Active", adChar, adParamInput,1)

			.Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("FileID", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("Loaded", adChar, adParamInput,1)    
			.Parameters.Append objCmd.CreateParameter("CDMCIDOutput", adInteger, adParamOutput)				
            
        End If

		'response.write "lngCDMCID=" & lngCDMCID
			.Parameters("CDMCID") = lngCDMCID
			.Parameters("GroupName") = strGroupName
			.Parameters("DivisionName") = strDivisionName
			.Parameters("BranchName") = strBranchName
			.Parameters("DepartmentName") = strDepartmentName
			.Parameters("DepartmentNumber") = strDepartmentNumber
			.Parameters("CostCentre") = strCostCentre
			.Parameters("EmployeeID") = strEmployeeID
			.Parameters("EmployeeType") = strEmployeeType
			.Parameters("Firstname") = strFirstname
			.Parameters("Surname") = strSurname
			.Parameters("Title") = strTitle
			.Parameters("Email_Address") = strEmail_Address
			.Parameters("TelephoneNumber") = Left(strTelephoneNumber,30)
			.Parameters("MobileNumber") = Left(strMobileNumber,30)
			'Response.Write strDateofBirth
			.Parameters("DateofBirth") = strDateofBirth
			.Parameters("Gender") = strGender
			.Parameters("ActualRankLvl") = strActualRankLvl
			.Parameters("Site") = strSite
			.Parameters("Unit") = strUnit
			.Parameters("ReportsTo") = strReportsTo
			.Parameters("DCD_PostalAddress") = strDCD_PostalAddress
			.Parameters("addressline1") = straddressline1
			.Parameters("addressline2") = straddressline2
			.Parameters("addressline3") = straddressline3
			.Parameters("addressline4") = straddressline4
			.Parameters("addressline5") = straddressline5
			.Parameters("addressline6") = straddressline6
			.Parameters("PostalAddress_Unit") = strPostalAddress_Unit
			.Parameters("PostalAddress_ClientLocation") = strPostalAddress_ClientLocation
			.Parameters("PostalAddress_DeliveryLocation") = strPostalAddress_DeliveryLocation
			.Parameters("PostalAddress_City") = strPostalAddress_City
			.Parameters("PostalAddress_State") = strPostalAddress_State
			.Parameters("PostalAddress_PostCode") = strPostalAddress_PostCode
			.Parameters("PostalAddress_Country") = strPostalAddress_Country
			.Parameters("ClientLocation") = strClientLocation
			.Parameters("StreetAddress") = strStreetAddress
			.Parameters("City") = strCity
			.Parameters("State") = strState
			.Parameters("PostCode") = strPostCode
			.Parameters("DCDProtectedIdentity") = strDCDProtectedIdentity
			.Parameters("IsValidPostal") = strIsValidPostal
			.Parameters("OutAddr1") = strOutAddr1
			.Parameters("OutAddr2") = strOutAddr2
			.Parameters("OutAddr3") = strOutAddr3
			.Parameters("OutSuburb") = strOutSuburb
			.Parameters("OutState") = strOutState
			.Parameters("OutPostCode") = strOutPostCode
			.Parameters("PostalMessage") = strPostalMessage
			.Parameters("hasChanged") = strhasChanged
			.Parameters("FormalFirstName") = strFormalFirstName
			.Parameters("FormalLastName") = strFormalLastName
			.Parameters("FormalMiddleName") = strFormalMiddleName
			.Parameters("OutTitle") = strOutTitle
			.Parameters("OutDinersWorkPhone") = strOutDinersWorkPhone
			.Parameters("OutDinersMobilePhone") = strOutDinersMobilePhone
			.Parameters("OutANZPhone") = strOutANZPhone
			.Parameters("OutDinersAddress1") = strOutDinersAddress1
			.Parameters("OutDinersAddress2") = strOutDinersAddress2
			.Parameters("RemoveCountdown") = strRemoveCountdown
			.Parameters("FirstUpdated") = NULL 'strFirstUpdated
			.Parameters("LastUpdated") = NULL 'strLastUpdated
			.Parameters("Active") = strActive
			.Parameters("UpdatedBy") = Session("UserID")
			'.Parameters("DateUpdated") = strDateUpdated

			.Parameters("FileID") = strFileID
			.Parameters("Loaded") = "N"'strLoaded
					   
			.ActiveConnection = objCon
			
		End With
			
		objCmd.Execute        
		
		'Return the result of the Save Function.
		intRecord = objCmd.Parameters.Item("CDMCIDOutput")    
                    			     				     		     		
    
End Sub

'Get virtual Pathname and set to where spreadsheets are created and opened from.
Function GetFilePath()
  Dim lsPath, arPath

  ' Obtain the virtual file path. The SCRIPT_NAME
  lsPath = Request.ServerVariables("SCRIPT_NAME")
    
  ' Split the path along the /s.
  arPath = Split(lsPath, "/")

  ' Set the last item in the array to blank string
  ' (The last item actually is the file name)
  arPath(UBound(arPath,1)) = ""

  ' Join the items in the array.
  GetFilePath = Join(arPath, "/")
End Function

If Err.Number <> 0 then
    Response.Write "<font face=arial size=1><b>Error occured while trying to process the request<b><br>Please contact system administrator</font><br>"
    Error.Clear
End If


Public Function GetNextCDMCFileID()

	'Description:	Gets the next FILE ID number for ANZ Cardlist files.
	objRS.Open "SELECT TOP 1 [FileSeqNum] FROM tblCAPSFileLoad WHERE FileType = 'CDMC' AND [Deleted] = 'N' ORDER BY [FileSeqNum] DESC ",objCon

		If Not objRS.EOF Then
			GetNextCDMCFileID = objRS("FileSeqNum") + 1
		Else
			GetNextCDMCFileID = 1
		End If

	objRS.Close
	
End Function

Set objRS = Nothing
Set objCon = Nothing

Public Function ProcessFile(strProcessSeq,strFileSeqNum,strFileName,lngFileID)
'Function to Process an CDMC File which has been loaded into the database 
Dim intRecord
	'Process History
	
	If strProcessSeq = "A" Then
	
		With objCmd

			.CommandType = 4
			.CommandText = "spCAPSCDMCUpdateFromTemp"

			.Parameters.Append objCmd.CreateParameter("UserID", adInteger)
			.Parameters.Append objCmd.CreateParameter("CDMCOutput", adInteger, adParamOutput)
			
			.Parameters("UserID") = Session("UserID")	
			
			.ActiveConnection = objCon
			 
		End With
	   
		objCmd.Execute 

		'Return the result of the Save Function.
		intRecord = objCmd.Parameters.Item("CDMCOutput")
	
	End If
	
	If strProcessSeq = "B" Then
	
		With objCmd

			.CommandType = 4
			.CommandText = "spCAPSCDMCProcessContactDetails"

			.Parameters.Append objCmd.CreateParameter("UserID", adInteger)
			.Parameters.Append objCmd.CreateParameter("CDMCProcessOutput", adInteger, adParamOutput)
			
			.Parameters("UserID") = Session("UserID")	
			
			.ActiveConnection = objCon
			 
		End With
	   
		objCmd.Execute 

		'Return the result of the Save Function.
		intRecord = objCmd.Parameters.Item("CDMCProcessOutput")
	
	End If
	
	Call UpdateCDMCFileLoadSummary(strProcessSeq,"CDMC", strFileSeqNum, strFileName, lngFileID) 

	If intRecord = 0 Then
		Response.Write "<div class=""alert alert-danger"" role=""alert"">No Records Processed in CDMC Upload File " & strFileSeqNum & ". Please notify System Administrators.</div>"
	Else
		Response.Write "<div class=""alert alert-success"" role=""alert""> " & intRecord & " CDMC Upload File records Processed in file " & strFileSeqNum & "</div>"
	End If
	
	ProcessFile = intRecord
	
End Function

 %>


