<%
'Global CAPS functions to be included in most/all pages 
Public Function Validate_Access(UserTypeID,Screen)

Dim objRSFunc

Set objRSFunc = Server.CreateObject("ADODB.Recordset")

    If Session("UserTypeID") = 99 Then
        
        Validate_Access = "Y"
        
    Else
        
        objRSFunc.Open "SELECT ScreenID FROM qryScreenAccess WITH(NOLOCK) WHERE UserTypeID = " & UserTypeID & " AND PageName = '" & Screen & "?TransactionType=" & Session("TransactionType") & "'",objCon

            If objRSFunc.EOF Then
                Validate_Access = "N" 
            Else
                Validate_Access = "Y"
            End If
    
        objRSFunc.Close
    
    End If

	objRSFunc.Close
 
Set objRSFunc = Nothing

End Function

Public Function GetSystemAdmin(strParameterName)

If objCon = "" or isnull(objCon) Then

	Set objCon = Server.CreateObject("ADODB.Connection")

	objCon.Open Session("DBConnection")
End If

Dim objRSFunc

Set objRSFunc = Server.CreateObject("ADODB.Recordset")

	'GetSystemAdmin = "See System Admin"
	'Exit Function
	objRSFunc.Open "SELECT * FROM tblCAPSSystemParameters WITH(NOLOCK) WHERE ParameterName = '" & strParameterName & "'",objCon

		If objRSFunc.EOF Then
			GetSystemAdmin = strParameterName 
		Else
			GetSystemAdmin = objRSFunc(3)
			
		End If

	objRSFunc.Close
 
Set objRSFunc = Nothing
 
End Function

Public Function CheckANZFileStatus(strFileName)

Dim objRSFunc

Set objRSFunc = Server.CreateObject("ADODB.Recordset")

	objRSFunc.Open "SELECT * FROM tblCAPSFileLoad WITH(NOLOCK) WHERE Status = 'N' AND FileType = 'ANZCardList' AND FileName = '" & strFileName & "'",objCon

		If objRSFunc.EOF Then
			CheckANZFileStatus = ""
		Else
			CheckANZFileStatus = objRSFunc("Status")
		End If

	objRSFunc.Close
 
Set objRSFunc = Nothing
 
End Function


Public Function GetEntitlement(strType, strEmployeeType)
'Function which returns a "Y" or "N" (or error message) for the Card Type and Employee Type passed in
Dim objRSFunc

Set objRSFunc = Server.CreateObject("ADODB.Recordset")

	objRSFunc.Open "SELECT * FROM tblCAPSEmployeeType WITH(NOLOCK) WHERE EmployeeType = '" & strEmployeeType & "'",objCon

		If objRSFunc.EOF Then
			GetEntitlement = "EmployeeType does not exist"
		Else
			If strType = "DTC" Then
				GetEntitlement = objRSFunc(2)
			Else
				GetEntitlement = objRSFunc(3)
			End If
		End If

	objRSFunc.Close
 
Set objRSFunc = Nothing
 
End Function

Public Function GetCardNo(strCardID)
'Procedure to get the Card Number from the Card ID Passed in
Dim objRSFunc

Set objRSFunc = Server.CreateObject("ADODB.Recordset")

	'GetSystemAdmin = "See System Admin"
	'Exit Function
	objRSFunc.Open "SELECT [CardNumber] FROM tblCAPSCard WITH(NOLOCK) WHERE CardID = '" & strCardID & "'",objCon

		If objRSFunc.EOF Then
			GetCardNo = "" '& strParameterName 
		Else
			GetCardNo = objRSFunc("CardNumber")
		End If

	objRSFunc.Close
 
Set objRSFunc = Nothing
 
End Function

Public Function GetCardStatus(strCardID)
'Procedure to get the Card Status from the Card ID Passed in
Dim objRSFunc

Set objRSFunc = Server.CreateObject("ADODB.Recordset")

	'GetSystemAdmin = "See System Admin"
	'Exit Function
	objRSFunc.Open "SELECT Status FROM tblCAPSCard WITH(NOLOCK) WHERE CardID = '" & strCardID & "'",objCon

		If objRSFunc.EOF Then
			GetCardStatus = "" '& strParameterName 
		Else
			GetCardStatus = objRSFunc("Status")
		End If

	objRSFunc.Close
 
Set objRSFunc = Nothing
 
End Function


Public Function GetCardNoShort(strCardID)
'Procedure to get the Card Number short from the Card ID Passed in
Dim objRSFunc

Set objRSFunc = Server.CreateObject("ADODB.Recordset")

	'GetSystemAdmin = "See System Admin"
	'Exit Function
	objRSFunc.Open "SELECT [CardNumberShort] FROM tblCAPSCard WITH(NOLOCK) WHERE CardID = '" & strCardID & "'",objCon

		If objRSFunc.EOF Then
			GetCardNoShort = "" '& strParameterName 
		Else
			GetCardNoShort = objRSFunc("CardNumberShort")
		End If

	objRSFunc.Close
 
Set objRSFunc = Nothing
 
End Function

Public Function GetApplicantName(lngApplicationID,strEID)

Dim objRSFunc

Set objRSFunc = Server.CreateObject("ADODB.Recordset")

	'GetSystemAdmin = "See System Admin"
	'Exit Function
	objRSFunc.Open "SELECT [EmployeeID], [FirstName],[Surname] FROM tblCAPSApplication WITH(NOLOCK) WHERE ApplicationID = " & lngApplicationID & "",objCon

		If objRSFunc.EOF Then
			GetApplicantName = "" '& strParameterName 
		Else
			'If there is an EID passed in then return the EID at the start of the Applicant's name
			If strEID = "EID" Then
				GetApplicantName = objRSFunc("EmployeeID") & " - " & objRSFunc("FirstName") & " " & objRSFunc("Surname")
				'Set the ApplicationEmployeeID for the Application Detail screen
				Session("ApplicationEmployeeID") = objRSFunc("EmployeeID")
			Else
				GetApplicantName = objRSFunc("FirstName") & " " & objRSFunc("Surname")
			End If
		End If

	objRSFunc.Close
 
Set objRSFunc = Nothing
 
End Function

Public Function GetApplicationTypeName(lngApplicationID)

Dim objRSFunc

Set objRSFunc = Server.CreateObject("ADODB.Recordset")

	'GetSystemAdmin = "See System Admin"
	'Exit Function
	objRSFunc.Open "SELECT [ApplicationTypeName] FROM tblCAPSApplication WITH(NOLOCK) WHERE ApplicationID = " & lngApplicationID & "",objCon

		If objRSFunc.EOF Then
			GetApplicationTypeName = "" '& strParameterName 
		Else		
			GetApplicationTypeName = objRSFunc("ApplicationTypeName") 
		End If

	objRSFunc.Close
 
Set objRSFunc = Nothing
 
End Function

Public Function GetApplicationStatus(lngApplicationID)

Dim objRSFunc

Set objRSFunc = Server.CreateObject("ADODB.Recordset")

	'GetSystemAdmin = "See System Admin"
	'Exit Function
	objRSFunc.Open "SELECT [Status] FROM tblCAPSApplication WITH(NOLOCK) WHERE ApplicationID = " & lngApplicationID & "",objCon

		If objRSFunc.EOF Then
			GetApplicationStatus = "" '& strParameterName 
		Else		
			GetApplicationStatus = objRSFunc("Status") 
		End If

	objRSFunc.Close
 
Set objRSFunc = Nothing
 
End Function




Public Function GetEmployeeName(strEID,strEIDReturn)
'Function to return the Employee name from the CDMC History table for the EID passed in
Dim objRSFunc

Set objRSFunc = Server.CreateObject("ADODB.Recordset")

	'GetSystemAdmin = "See System Admin"
	'Exit Function
	objRSFunc.Open "SELECT [EmployeeID], [FirstName],[Surname] FROM qryCAPSCDMCHistoryActive WITH(NOLOCK) WHERE EmployeeID = '" & strEID & "'",objCon

		If objRSFunc.EOF Then
			GetEmployeeName = "No CDMC History for " & strEID 
		Else
			'If there is an EID passed in then return the EID at the start of the Applicant's name
			If strEIDReturn = "Y" Then
				GetEmployeeName = objRSFunc("EmployeeID") & " - " & objRSFunc("FirstName") & " " & objRSFunc("Surname")
			Else
				GetEmployeeName = objRSFunc("FirstName") & " " & objRSFunc("Surname")
			End If
		End If

	objRSFunc.Close
 
Set objRSFunc = Nothing
 
End Function


Public Function GetTitleList(strType,strDelimiter,strActive,strSelected)
'Function to return a list of Titles from the database
'strType = whether it is the options for a select object or just delimited
'strDelimier = the delimiter used between records (empty string for select option)
'strActive = The State (Active) of records returned
'strSelected = The record selected (if any)
Dim objRSFunc
Dim strSelectText

Set objRSFunc = Server.CreateObject("ADODB.Recordset")

	If strActive = "All" Then
		strActive = ""
	Else
		strActive = "WHERE [ACTIVE] = '" & strActive & "'"
	End If
	
	'Open a recordset of the titles in the database table
	objRSFunc.Open "SELECT [Title] FROM tblCAPSTitle WITH(NOLOCK) " & strActive & "",objCon

		If objRSFunc.EOF Then
			GetTitleList = ""
		Else
		
			Do Until objRSFunc.EOF
				'If the record is selected then set the text in the option to selected
				If UCASE(Trim(cstr(strSelected))) = UCASE(Trim(cstr(objRSFunc("Title")))) Then
					strSelectText = "SELECTED"
				Else
					strSelectText = " NO "
				End If
				
				'Build the return string for this function based on the type passed in
				If strType = "Select" Then
					GetTitleList = GetTitleList & "<Option value=""" & objRSFunc("Title") & """ " & strSelectText & ">" & objRSFunc("Title") & "</option>" 
				Else
					GetTitleList = GetTitleList & objRSFunc("Title") & strDelimiter
				End If
				
			objRSFunc.Movenext
			Loop
			
		End If

	objRSFunc.Close
 
	'Remove the delimiter
	If strType = "Select" Then
	Else
		GetTitleList = Left(GetTitleList,Len(GetTitleList)-Len(strDelimiter))
	End If
				
Set objRSFunc = Nothing
 
End Function


Public Function SaveFileLoadID(strFileType, strFileName, strFilePath, lngRecordCount,lngCardCount,lngEmployeeCount,lngDTCCount,lngCMCCount,lngCTSCount,lngDPCCount,_ 
							lngFooterCount,strFileDateTime,strFileSeqNum,strStatus,lngLoadedBy,strDeleted)

Dim intRecord
Dim objConF
Dim objCmdF
'response.write "</br>spCAPSFileLoadSave" & strFileType & "," & strFileName & "," & strFilePath & "," & lngRecordCount & "," & lngCardCount & "," & lngEmployeeCount & "," & lngDTCCount & "," & lngCMCCount & "," & strFileDateTime & "," & strFileSeqNum & "," & strStatus
'Open Database Connection
Set objConF = Server.CreateObject("ADODB.Connection")
Set ObjCmdF = Server.CreateObject("ADODB.Command")

objConF.Open Session("DBConnection")

  	With objCmdF
  	
  	    'If the procedure has already run then don't create the parameter objects again (more than once)
  	    'If x = 1 then
			.CommandType = 4
			.CommandText = "spCAPSFileLoadSave"
			
			.Parameters.Append objCmdF.CreateParameter("FileLoadID", adInteger)
			.Parameters.Append objCmdF.CreateParameter("FileType", adVarchar, adParamInput,20)
			.Parameters.Append objCmdF.CreateParameter("FileName", adVarchar, adParamInput,100)
			.Parameters.Append objCmdF.CreateParameter("FilePath", adVarchar, adParamInput,500)
			.Parameters.Append objCmdF.CreateParameter("RecordCount", adInteger)
			.Parameters.Append objCmdF.CreateParameter("CardCount", adInteger)
			.Parameters.Append objCmdF.CreateParameter("EmployeeCount", adInteger)
			.Parameters.Append objCmdF.CreateParameter("DTCCount", adInteger)
			.Parameters.Append objCmdF.CreateParameter("CMCCount", adInteger)
			.Parameters.Append objCmdF.CreateParameter("CTSCount", adInteger)
			.Parameters.Append objCmdF.CreateParameter("DPCCount", adInteger)
			.Parameters.Append objCmdF.CreateParameter("FooterCount", adInteger)
			.Parameters.Append objCmdF.CreateParameter("FileDateTime", adVarchar, adParamInput,20)
			.Parameters.Append objCmdF.CreateParameter("FileSeqNum", adVarchar, adParamInput,20)
			.Parameters.Append objCmdF.CreateParameter("Status", adVarchar, adParamInput,20)
			.Parameters.Append objCmdF.CreateParameter("LoadedBy", adInteger)
			.Parameters.Append objCmdF.CreateParameter("Deleted", adChar, adParamInput,1)
			.Parameters.Append objCmdF.CreateParameter("FileLoadIDOutput", adInteger, adParamOutput)				
            
        'End If
                 
			.Parameters("FileLoadID") = 0'strCSFromDinersID
			.Parameters("FileType") = strFileType
			.Parameters("FileName") = strFileName
			.Parameters("FilePath") = strFilePath
			.Parameters("RecordCount") = lngRecordCount
			.Parameters("CardCount") = lngCardCount
			.Parameters("EmployeeCount") = lngEmployeeCount
			.Parameters("DTCCount") = lngDTCCount
			.Parameters("CMCCount") = lngCMCCount
			.Parameters("CTSCount") = lngCTSCount
			.Parameters("DPCCount") = lngDPCCount
			.Parameters("FooterCount") = lngFooterCount
			.Parameters("FileDateTime") = strFileDateTime
			.Parameters("FileSeqNum") = strFileSeqNum
			.Parameters("Status") = strStatus
			.Parameters("LoadedBy") = lngLoadedBy
			.Parameters("Deleted") = strDeleted
		
		.ActiveConnection = objConF
                
    End With
                
	objCmdF.Execute        
	
	'Return the result of the Save Function.
	SaveFileLoadID = objCmdF.Parameters.Item("FileLoadIDOutput")    
    
objConF.Close
		
Set objConF = Nothing
Set ObjCmdF = Nothing
     
	   
End Function


Public Sub UpdateFileLoadSummary(strFileType, strFileSeqNum, strFileName, lngFileLoadID)
'Procedure to run a stored procedure which updates the summary details for a file just loaded, which is used where summary details are displayed

Dim objConF
Dim objCmdF

'Open Database Connection
Set objConF = Server.CreateObject("ADODB.Connection")
Set ObjCmdF = Server.CreateObject("ADODB.Command")

objConF.Open Session("DBConnection")

  	With objCmdF
  	
		.CommandType = 4
		.CommandText = "spCAPSFileLoadSummaryUpdate"
		
		.Parameters.Append objCmdF.CreateParameter("FileType", adVarchar, adParamInput,20)
		.Parameters.Append objCmdF.CreateParameter("FileSeqNum", adVarchar, adParamInput,20)
		.Parameters.Append objCmdF.CreateParameter("FileName", adVarchar, adParamInput,100)
        .Parameters.Append objCmdF.CreateParameter("FileLoadID", adInteger)
		
		.Parameters("FileType") = strFileType
		.Parameters("FileSeqNum") = strFileSeqNum
		.Parameters("FileName") = strFileName
		.Parameters("FileLoadID") = lngFileLoadID
		
		.ActiveConnection = objConF
                
    End With
                
	objCmdF.Execute        
	
		
Set objConF = Nothing
Set ObjCmdF = Nothing

End Sub

Public Sub UpdateCDMCFileLoadSummary(strProcessSeq,strFileType, strFileSeqNum, strFileName, lngFileLoadID)
'Procedure to run a stored procedure which updates the summary details for a file just loaded, which is used where summary details are displayed
'response.write "EXEC spCAPSCDMCFileLoadSummaryUpdate " & strProcessSeq  & "," & strFileType  & "," & strFileSeqNum  & "," & strFileName  & "," & lngFileLoadID 
Dim objConF
Dim objCmdF

'Open Database Connection
Set objConF = Server.CreateObject("ADODB.Connection")
Set ObjCmdF = Server.CreateObject("ADODB.Command")

objConF.Open Session("DBConnection")

  	With objCmdF
  	
		.CommandType = 4
		.CommandText = "spCAPSCDMCFileLoadSummaryUpdate"
		
		.Parameters.Append objCmdF.CreateParameter("Process", adVarchar, adParamInput,1)
		.Parameters.Append objCmdF.CreateParameter("FileType", adVarchar, adParamInput,20)
		.Parameters.Append objCmdF.CreateParameter("FileSeqNum", adVarchar, adParamInput,20)
		.Parameters.Append objCmdF.CreateParameter("FileName", adVarchar, adParamInput,100)
        .Parameters.Append objCmdF.CreateParameter("FileLoadID", adInteger)
		
		.Parameters("Process") = strProcessSeq
		.Parameters("FileType") = strFileType
		.Parameters("FileSeqNum") = strFileSeqNum
		.Parameters("FileName") = strFileName
		.Parameters("FileLoadID") = lngFileLoadID
		
		.ActiveConnection = objConF
                
    End With
                
	objCmdF.Execute        
	
		
Set objConF = Nothing
Set ObjCmdF = Nothing

End Sub

Public Function GetFileLoadID(strFileType,strFileSeqNum,strFileName)
'Procedure to check if there is a record in the table tblCAPSFileUpload for the File Type and File Sequence Number passed in

Dim objRSFunc
Dim strSQLF

Set objRSFunc = Server.CreateObject("ADODB.Recordset")

	If strFileType = "CSFromDiners" or strFileType = "CSFromDinersDPC" OR strFileType = "NAFile" or strFileType = "ProMasterUser" or strFileType = "ProMasterAccount" Then
		strSQLF = "SELECT * FROM tblCAPSFileLoad WITH(NOLOCK) WHERE [FileType] = '" & strFileType & "' AND [FileSeqNum] = '" & strFileSeqNum & "' AND [Deleted] <> 'Y'"
		
	Else
		If strFileType = "CDMC" Then
			strSQLF = "SELECT * FROM tblCAPSFileLoad WITH(NOLOCK) WHERE [FileType] = '" & strFileType & "' AND [FileName] = '" & strFileName & "' AND [Deleted] <> 'Y' AND Cast(DateLoaded as Date) = '" & MediumDate(Date()) & "'"
		Else
			strSQLF = "SELECT * FROM tblCAPSFileLoad WITH(NOLOCK) WHERE [FileType] = '" & strFileType & "' AND [FileName] = '" & strFileName & "' AND [Deleted] <> 'Y'"
		End If
	End If
	
	objRSFunc.Open strSQLF,objCon

		If objRSFunc.EOF Then
			'Return 1 for ROMAN load as this will always overwrite existing files. Other files return an empty variable to note no existing file has been loaded for that type.
			
			If strFileType = "ROMANCostCentres" Or strFileType = "CDMC" OR strFileType = "NAFile" OR strFileType = "ProMasterUser" OR strFileType = "ProMasterAccount" Then
				GetFileLoadID = 1			
						
			Else
				GetFileLoadID = ""
					
			End If
		Else
			'If strFileType = "CSFromDiners" Then
				GetFileLoadID = objRSFunc("FileSeqNum") + 1
							
			'Else
			'	GetFileLoadID = objRSFunc("FileName")
				
			'End If
		End If

	objRSFunc.Close
 
Set objRSFunc = Nothing


End Function

Public Function GetLastFileLoadID(strFileType,strFileName)
'Procedure to get the last FileSeqNum from tblCAPSFileUpload for the File Type and File Sequence Number passed in

Dim objRSFunc
Dim strSQL

Set objRSFunc = Server.CreateObject("ADODB.Recordset")

	If strFileType = "CSFromDiners" or strFileType = "ROMANCostCentres" or strFileType = "ProMasterAccount" or strFileType = "ProMasterUser" Then
		strSQL = "SELECT Top 1 FileSeqNum FROM tblCAPSFileLoad WITH(NOLOCK) WHERE [FileType] = '" & strFileType & "' Order By FileLoadID Desc"
	Else
		strSQL = "SELECT Top 1 FileSeqNum FROM tblCAPSFileLoad WITH(NOLOCK) WHERE [FileType] = '" & strFileType & "' AND [FileName] = '" & strFileName & "' Order By FileSeqNum Desc"		
	End If
	
	objRSFunc.Open strSQL,objCon
	
		If objRSFunc.EOF Then
			'Return 0 if no files exists
			GetLastFileLoadID = 0
		Else
			'Return the last file Seq Num		
			GetLastFileLoadID = objRSFunc(0)			
		End If

	objRSFunc.Close
 
Set objRSFunc = Nothing

End Function

Public Function GetLastCSFileNumber()
'Procedure to get the the last CS File Number

Dim objRSFunc
Dim strSQL

Set objRSFunc = Server.CreateObject("ADODB.Recordset")
	
	strSQL = "SELECT ParameterValue FROM tblCAPSSystemParameters WITH(NOLOCK) WHERE SystemParameterID = 16"	
		
	objRSFunc.Open strSQL,objCon
	
		If objRSFunc.EOF Then
			'Return 0 if no files exists
			GetLastCSFileNumber = 0
		Else
			'Return the CS File Number	
			GetLastCSFileNumber = objRSFunc(0)			
		End If

	objRSFunc.Close
 
Set objRSFunc = Nothing

End Function

Public Function CheckNumber(strNumber)

	If IsNull(strNumber) or strNumber = "" Then
		CheckNumber = 0
	Else
		CheckNumber = strNumber
	End If

End Function

Public Function CheckString(strString)

	If IsNull(strString) or strString = "" Then
		CheckString = ""
	Else
		CheckString = strString
	End If

End Function


Public Function GetTitleFromNameOnCard(strNameOnCard)
'Function to get any title from a Name On Card (which is one string) and make sure that it is a valid title (from the table tblCAPStitlE)
Dim objRSFunc
Dim strNameOnCard2

	'Exit if there is nothing passed in
	If IsNull(strNameOnCard) or strNameOnCard = "" Then 
		GetTitleFromNameOnCard = ""
		Exit Function
	End If
	
	'Exit if there are no spaces in the string passed in
	If Instr(strNameOnCard," ") = 0 Then
		GetTitleFromNameOnCard = ""
		Exit Function
	End If
	
	'Get the First word of the name on card (characters up to the first space)
	strNameOnCard2 = Trim(Left(strNameOnCard,Instr(strNameOnCard," ")-1))
	
	Set objRSFunc = Server.CreateObject("ADODB.Recordset")
	
	'Open a recordset of the titles in the database table
	objRSFunc.Open "SELECT [Title] FROM tblCAPSTitle WITH(NOLOCK) WHERE Active = 'Y'",objCon

		If objRSFunc.EOF Then
			GetTitleFromNameOnCard = ""
		Else
		
			Do Until objRSFunc.EOF
				
				'If the Title (first word of the string passed in) matches a title in the database, then return the same title
				If Trim(UCASE(strNameOnCard2)) = Trim(UCASE(objRSFunc("Title"))) Then
					GetTitleFromNameOnCard = strNameOnCard2
				End If
				
			objRSFunc.Movenext
			Loop

		End If

	objRSFunc.Close
				
	Set objRSFunc = Nothing

End Function


Function MediumDate(str)
'Function to change all date formats to medium date to avoid American storage challenge!	
Dim aDay
Dim aMonth
Dim aYear

	'Check to see whether the date passed in uses dashes (-) or slashes (/)
	If Instr(1,str,"/") = 0 Then
		If Mid(str,2,1) = "-" Then
			aDay = (Left((str),InStr(1,(str),"-")-1))
			aMonth = Mid(str,(InStr(1,(str),"-")+1),2)
		Else
			aDay = Mid((str),9,2)
			aMonth = Mid(str,(InStr(1,(str),"-")+1),2)
		End If
		
		If Right(aMonth,1) = "-" Then
			aMonth = Left(aMonth,1)
		End If
	Else
		If Mid(str,2,1) = "/" Then
			aDay = (Left((str),InStr(1,(str),"/")-1))
			aMonth = Mid(str,(InStr(1,(str),"/")+1),2)
		Else
			'aDay = Mid((str),9,2)
			aMonth = Mid(str,(InStr(1,(str),"/")+1),2)
			aDay = (Left((str),InStr(1,(str),"/")-1))
			
		End If
		
		If Right(aMonth,1) = "/" Then
			aMonth = Left(aMonth,1)
		End If

	End If
	
	aMonth = MonthName(aMonth)
	aYear = Year(str)
	
	If Len(aDay) = 1 Then aDay = "0" & aDay
	
	MediumDate = aDay & "-" & aMonth & "-" & aYear

End Function

Function MediumDate2 (str)
	
	'Function to change all date formats to medium date to avoid American storage challenge!
	
	Dim aDay
	Dim aMonth
	Dim aYear
	
		aDay = 	(Left((str),InStr(1,(str),"/")-1))
		aMonth = Mid(str,(InStr(1,(str),"/")+1),2)
	
	If Right(aMonth,1) = "/" Then
		aMonth = Left(aMonth,1)
	End If
	
		aMonth = MonthName(aMonth)
		aYear = Year(str)
	
	If Len(aDay) = 1 Then aDay = "0" & aDay
	
		MediumDate2 = aDay & "-" & aMonth & "-" & aYear
		
End Function


Public Function MaskCard(strCardNumber)
'Function to format a Card number so it is masked depending on the card type

	'If there is no number then pass back an empty string
	If IsNull(strCardNumber) Then
		MaskCard = ""
	Else
		If Trim(strCardNumber) = "" Then
			MaskCard = ""
		Else
			'If the 5th digit of the card number is a zero then it is a Mastercard
			If mid(strCardNumber,5,1)=0 Then 
				MaskCard = mid(strCardNumber,6,2) & "****" & right(strCardNumber,4)
			'Otherwise it is a Diners card
			ElseIf Left(strCardNumber,5) = "47152" Then
				MaskCard = left(strCardNumber,2) & "****" & right(trim(strCardNumber),4)
			ElseIf Left(strCardNumber,1)=4 Then
				MaskCard = left(strCardNumber,2) & "****" & right(strCardNumber,4)
			Else
				MaskCard = mid(strCardNumber,4,2) & "****" & right(trim(strCardNumber),4)
			End If
		End If
	End If

End Function


Public Function FormatCardNumber(strCardNumber)
'Function to format a Card number so it is spaced by groups of digits, based on the card type

	'If there is no number then pass back an empty string
	If IsNull(strCardNumber) Then
		FormatCardNumber = ""
	Else
		If Trim(strCardNumber) = "" Then
			FormatCardNumber = ""
		Else
			'If the 5th digit of the card number is a zero then it is a Mastercard
			If left(strCardNumber,5) = "47152" Then
				'NAB Card Number
				FormatCardNumber = mid(strCardNumber,1,4) & " " & mid(strCardNumber,5,4) & " " & mid(strCardNumber,9,4) & " " & mid(strCardNumber,13,4)
			ElseIF  left(strCardNumber,1)=4 Then
				FormatCardNumber = mid(strCardNumber,1,4) & " " & mid(strCardNumber,5,4) & " " & mid(strCardNumber,9,4) & " " & right(strCardNumber,4)
				
			ElseIf Mid(strCardNumber,6,1)=3 Then 
				'FormatCardNumber = mid(strCardNumber,6,4) & " " & mid(strCardNumber,11,4) & " " & mid(strCardNumber,16,4) & " " & right(strCardNumber,4)
				FormatCardNumber = mid(strCardNumber,6,2) & " " & mid(strCardNumber,8,4) & " " & mid(strCardNumber,12,4) & " " & right(strCardNumber,4)
		
			'ElseIf left(strCardNumber,1)=4 Then
				'Visa/ANZ Card Number
			'	FormatCardNumber = mid(strCardNumber,1,4) & " " & mid(strCardNumber,5,4) & " " & mid(strCardNumber,9,4) & " " & right(strCardNumber,4)
			Else
				'FormatCardNumber = mid(strCardNumber,4,2) & "~" & mid(strCardNumber,7,4) & " " & mid(strCardNumber,12,4) & " " & right(strCardNumber,4)
				FormatCardNumber = mid(strCardNumber,4,4) & " " & mid(strCardNumber,8,4) & " " & mid(strCardNumber,12,4) & " " & right(strCardNumber,4)
			End If
		End If
	End If

End Function

Public Function SaveAuditLog(lngAuditLogID,strType,strSubType,strEID,strCardType,strCardNumber,strActionedBy,strValueBefore,strValueAfter,strSourceFile,strChangeDetails,lngCardID,lngApplicationID,lngCSFromDinersID,lngCSToDinersID,strProcess)
'Function to save a record to the CAPS Audit Log and return the ID of the record saved

Dim objConF
Dim objCmdF

If lngAuditLogID = "" or IsNull(lngAuditLogID) Then lngAuditLogID = 0
If lngCardID = "" or IsNull(lngCardID) Then lngCardID = 0
If lngApplicationID = "" or IsNull(lngApplicationID) Then lngApplicationID = 0
If lngCSFromDinersID = "" or IsNull(lngCSFromDinersID) Then lngCSFromDinersID = 0
If lngCSToDinersID = "" or IsNull(lngCSToDinersID) Then lngCSToDinersID = 0

'Open Database Connection
Set objConF = Server.CreateObject("ADODB.Connection")
Set ObjCmdF = Server.CreateObject("ADODB.Command")

objConF.Open Session("DBConnection")

		With objCmdF

			.CommandType = 4
			.CommandText = "spCAPSAuditLogSave"

			.Parameters.Append objCmdF.CreateParameter("AuditLogID", adInteger)
			.Parameters.Append objCmdF.CreateParameter("ChangeDate", adInteger)
			.Parameters.Append objCmdF.CreateParameter("Type", adVarChar, adParamInput,20)
			.Parameters.Append objCmdF.CreateParameter("SubType", adVarChar, adParamInput,30)
			.Parameters.Append objCmdF.CreateParameter("EID", adVarChar, adParamInput, 20)
			.Parameters.Append objCmdF.CreateParameter("CardType", adVarChar, adParamInput, 20)
			.Parameters.Append objCmdF.CreateParameter("CardNumber", adVarChar, adParamInput, 20)
			.Parameters.Append objCmdF.CreateParameter("ActionedBy", adVarChar, adParamInput, 50)
			.Parameters.Append objCmdF.CreateParameter("ValueBefore", adVarChar, adParamInput, 200)
			.Parameters.Append objCmdF.CreateParameter("ValueAfter", adVarChar, adParamInput, 200)
			.Parameters.Append objCmdF.CreateParameter("SourceFile", adVarChar, adParamInput, 100)
			.Parameters.Append objCmdF.CreateParameter("ChangeDetails", adVarChar, adParamInput, 200)
			.Parameters.Append objCmdF.CreateParameter("CardID", adInteger)
			.Parameters.Append objCmdF.CreateParameter("ApplicationID", adInteger)
			.Parameters.Append objCmdF.CreateParameter("CSFromDinersID", adInteger)
			.Parameters.Append objCmdF.CreateParameter("CSToDinersID", adInteger)
			.Parameters.Append objCmdF.CreateParameter("UpdatedBy", adInteger)
			.Parameters.Append objCmdF.CreateParameter("Process", adVarChar, adParamInput, 50)
			.Parameters.Append objCmdF.CreateParameter("AuditLogIDOutput", adInteger, adParamOutput)
			
			.Parameters("AuditLogID") = lngAuditLogID
			.Parameters("ChangeDate") = now()
			.Parameters("Type") = strType
			.Parameters("SubType") = strSubType
			.Parameters("EID") = strEID
			.Parameters("CardType") = strCardType
			.Parameters("CardNumber") = strCardNumber
			.Parameters("ActionedBy") = strActionedBy
			.Parameters("ValueBefore") = strValueBefore
			.Parameters("ValueAfter") = strValueAfter
			.Parameters("SourceFile") = strSourceFile
			.Parameters("ChangeDetails") = strChangeDetails
			.Parameters("CardID") = lngCardID
			.Parameters("ApplicationID") = lngApplicationID
			.Parameters("CSFromDinersID") = lngCSFromDinersID
			.Parameters("CSToDinersID") = lngCSToDinersID
			.Parameters("UpdatedBy") = Session("UserID")
			.Parameters("CardID") = lngCardID
			.Parameters("Process") = strProcess
			
			.ActiveConnection = objCon
			 
		End With
	   
		objCmdF.Execute        
	  
		'Return the result of the Save Function.
		SaveAuditLog = objCmdF.Parameters.Item("AuditLogIDOutput") 
	        
	
Set objConF = Nothing
Set ObjCmdF = Nothing

End Function

Public Function SaveCSToDiners(lngCSToDinersID,strFileDateTime,strFileSeqNum,strEIDNo,strCardNo,strCardUpdateInd,strCardExpiryDate,strCardStatus,_
		strTitle,strSurname,strGivenNames,strNameOnCard,strAddress1,strAddress2,strAddress3,strSuburb,strState,strPostCode,strHomePhone,strWorkPhone,strMobilePhone,strEmail,strReportGroup,strCreditLimit,strStatus,strNotes)
'Function to save a record to the CS To Diners table and return the ID of the record saved

Dim objConF
Dim objCmdF

If lngCSToDinersID = "" or IsNull(lngCSToDinersID) Then lngCSToDinersID = 0

'Open Database Connection
Set objConF = Server.CreateObject("ADODB.Connection")
Set ObjCmdF = Server.CreateObject("ADODB.Command")

objConF.Open Session("DBConnection")

		With objCmdF

			.CommandType = 4
			.CommandText = "spCAPSCSToDinersSave"

			.Parameters.Append objCmdF.CreateParameter("CSToDinersID", adInteger)
			.Parameters.Append objCmdF.CreateParameter("FileDateTime", adVarChar, adParamInput,14)
			.Parameters.Append objCmdF.CreateParameter("FileSeqNum", adVarChar, adParamInput,6)
			.Parameters.Append objCmdF.CreateParameter("EID", adVarChar, adParamInput, 20)
			.Parameters.Append objCmdF.CreateParameter("CardNo", adVarChar, adParamInput, 19)
			.Parameters.Append objCmdF.CreateParameter("CardUpdateInd", adVarChar, adParamInput, 2)
			.Parameters.Append objCmdF.CreateParameter("CardExpiryDate", adDate)
			.Parameters.Append objCmdF.CreateParameter("CardStatus", adVarChar, adParamInput, 2)
			.Parameters.Append objCmdF.CreateParameter("Title", adVarChar, adParamInput, 12)
			.Parameters.Append objCmdF.CreateParameter("Surname", adVarChar, adParamInput, 25)
			.Parameters.Append objCmdF.CreateParameter("GivenNames", adVarChar, adParamInput, 30)
			.Parameters.Append objCmdF.CreateParameter("NameOnCard", adVarChar, adParamInput, 26)
			.Parameters.Append objCmdF.CreateParameter("Address1", adVarChar, adParamInput, 40)
			.Parameters.Append objCmdF.CreateParameter("Address2", adVarChar, adParamInput, 40)
			.Parameters.Append objCmdF.CreateParameter("Address3", adVarChar, adParamInput, 40)
			.Parameters.Append objCmdF.CreateParameter("Suburb", adVarChar, adParamInput, 25)
			.Parameters.Append objCmdF.CreateParameter("State", adVarChar, adParamInput,4)
			.Parameters.Append objCmdF.CreateParameter("PostCode", adVarChar, adParamInput, 12)
			.Parameters.Append objCmdF.CreateParameter("HomePhone", adVarChar, adParamInput, 12)
			.Parameters.Append objCmdF.CreateParameter("WorkPhone", adVarChar, adParamInput, 12)
			.Parameters.Append objCmdF.CreateParameter("MobilePhone", adVarChar, adParamInput, 12)
			.Parameters.Append objCmdF.CreateParameter("Email", adVarChar, adParamInput, 70)
			.Parameters.Append objCmdF.CreateParameter("ReportGroup", adVarChar, adParamInput, 8)
			.Parameters.Append objCmdF.CreateParameter("CreditLimit", adVarChar, adParamInput, 11)
			.Parameters.Append objCmdF.CreateParameter("Status", adVarChar, adParamInput, 20)
			.Parameters.Append objCmdF.CreateParameter("Notes", adVarChar, adParamInput, 100)
			.Parameters.Append objCmdF.CreateParameter("CSToDinersIDDOutput", adInteger, adParamOutput)
			
			.Parameters("CSToDinersID") = lngCSToDinersID
			.Parameters("FileDateTime") = strFileDateTime
			.Parameters("FileSeqNum") = strFileSeqNum
			.Parameters("EID") = strEID
			.Parameters("CardNo") = strCardNo
			.Parameters("CardUpdateInd") = strCardUpdateInd
			.Parameters("CardExpiryDate") = strCardExpiryDate
			.Parameters("CardStatus") = strCardStatus
			.Parameters("Title") = strTitle
			.Parameters("Surname") = strSurname
			.Parameters("GivenNames") = strGivenNames
			.Parameters("NameOnCard") = strNameOnCard
			.Parameters("Address1") = strAddress1
			.Parameters("Address2") = strAddress2
			.Parameters("Address3") = strAddress3
			.Parameters("Suburb") = strSuburb
			.Parameters("State") = strState
			.Parameters("PostCode") = strPostCode
			.Parameters("HomePhone") = strHomePhone
			.Parameters("WorkPhone") = strWorkPhone
			.Parameters("MobilePhone") = strMobilePhone
			.Parameters("Email") = strEmail
			.Parameters("ReportGroup") = strReportGroup
			.Parameters("CreditLimit") = strCreditLimit
			.Parameters("Status") = strStatus
			.Parameters("Notes") = strNotes
			.ActiveConnection = objCon
			 
		End With
	   
		objCmdF.Execute        
	  
		'Return the result of the Save Function.
		SaveCSToDiners = objCmdF.Parameters.Item("CSToDinersIDDOutput") 
	        
	
Set objConF = Nothing
Set ObjCmdF = Nothing

End Function

Public Function CancelCardToCS(lngCSToDinersID,lngCardID,strNotes)
'Function to save a record to the CS To Diners table and return the ID of the record saved when a Card is being cancelled

Dim objConF
Dim objCmdF

If lngCSToDinersID = "" or IsNull(lngCSToDinersID) Then lngCSToDinersID = 0

'Open Database Connection
Set objConF = Server.CreateObject("ADODB.Connection")
Set ObjCmdF = Server.CreateObject("ADODB.Command")

objConF.Open Session("DBConnection")

		With objCmdF

			.CommandType = 4
			.CommandText = "spCAPSCSToDinersCancelCard"

			.Parameters.Append objCmdF.CreateParameter("CSToDinersID", adInteger)
			'.Parameters.Append objCmdF.CreateParameter("EID", adVarChar, adParamInput, 20)
			.Parameters.Append objCmdF.CreateParameter("CardID", adInteger)
			.Parameters.Append objCmdF.CreateParameter("UpdatedBy", adInteger)
			'.Parameters.Append objCmdF.CreateParameter("CardNo", adVarChar, adParamInput, 19)
			.Parameters.Append objCmdF.CreateParameter("Notes", adVarChar, adParamInput, 100)
			.Parameters.Append objCmdF.CreateParameter("CSToDinersIDDOutput", adInteger, adParamOutput)
			
			.Parameters("CSToDinersID") = lngCSToDinersID
			'.Parameters("EID") = strEID
			.Parameters("CardID") = lngCardID
			.Parameters("UpdatedBy") = Session("UserID")
			'.Parameters("CardNo") = strCardNo
			.Parameters("Notes") = strNotes
			.ActiveConnection = objCon
			 
		End With
	   
		objCmdF.Execute        
	  
		'Return the result of the Save Function.
		CancelCardToCS = objCmdF.Parameters.Item("CSToDinersIDDOutput") 
	        
Set objConF = Nothing
Set ObjCmdF = Nothing

End Function

Public Function Send_Email(strFrom,strTo,strSubject,strBody,strAttachment,strEmailType)

Dim objEmail

	Set objEmail = CreateObject("CDO.Message")

		objEmail.From = strFrom
		objEmail.To = strTo
		objEmail.Subject = strSubject
		
		If strAttachment <> "" Then
			objEmail.AddAttachment(strAttachment)
		End If
		
		'If the Email Type is HMTL make the Body HTML Otherwise the email will be plain text
		If strEmailType = "HTML" Then
			objEmail.HTMLBody = strBody
		Else
			objEmail.TextBody = strBody
		End If
		
		objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "relay.dpesit.protectedsit.mil.au" 
		'objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtpmailer" '''old default setting replaced by relay above
		objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
		objEmail.Configuration.Fields.Update

		objEmail.Send
		
	Set objEmail = Nothing
	
	Send_Email = "Email Sent"

End Function

Function PadDigits(val, digits)
  PadDigits = Right(String(digits,"0") & val, digits)
End Function

Function PadDigitsLeft(val, digits)
  PadDigitsLeft = Left(val & String(digits,"0"), digits)
End Function

Function PadSpaceLeft(val, digits)
  PadSpaceLeft = Left(val & String(digits," "), digits)
End Function

Function PadSpaceRight(val, digits)
  PadSpaceRight = Right(String(digits," ") & val, digits)
End Function


Public Function GetDateLeftDefence(strEmployeeID)
'Function to return the Date the Employee passed in has left Defence (if they have left)
Dim strSQL
Dim strActive
Dim objRSFunc
Dim strDaysAgo
Dim dteDateLeft

Set objRSFunc = Server.CreateObject("ADODB.Recordset")

If strEmployeeID = "" Then
	GetDateLeftDefence = "1"
	Exit Function
Else

	strSQL = "SELECT [ActiveEmployee],[LastUpdated] FROM qryCAPSCDMCHistory WITH(NOLOCK) WHERE EmployeeID = '" & strEmployeeID & "' AND Deleted = 'N' ORDER BY [DateUpdated] DESC"
End If

objRSFunc.Open strSQL,objCon
   	
    If Not objRSFunc.EOF Then
	
		
		If IsNull(objRSFunc("ActiveEmployee")) OR objRSFunc("ActiveEmployee") = "" Then
			strActive = "Y"
		Else
			strActive = objRSFunc("ActiveEmployee")
		End If

		IF strActive = "N" then
			'If IsNull(objRSFunc("LastUpdated")) or objRSFunc("LastUpdated")="" Then
			'	GetDateLeftDefence = "No Date"
			'Else
			'	GetDateLeftDefence = objRS("LastUpdated")
			'End If
			
			If IsNull(objRSFunc("LastUpdated")) or objRSFunc("LastUpdated")="" Then
				dteDateLeft = "No Date"
			Else
				dteDateLeft = FormatDateTime(objRSFunc("LastUpdated"),vbShortDate)
				strDaysAgo = "title=""" & DateDiff("D",objRSFunc("LastUpdated"),Now()) & " Days ago"""
			End If

			GetDateLeftDefence = "<span " & strDaysAgo & " class=""badge badge-pill badge-danger"">" & dteDateLeft & "</span>"

		Else
			GetDateLeftDefence = "<span class=""badge badge-pill badge-success"">Active</span>"
		End If

	Else
		GetDateLeftDefence = "2"
	End If
	
				
objRSFunc.Close
 
Set objRSFunc = Nothing

End Function


Public Function GetEIDOnCSFile(strEmplID)
'Function to check whether an EmployeeID is on the CS file for today and returning a Y or N
Dim objRSFunc

Set objRSFunc = Server.CreateObject("ADODB.Recordset")

	objRSFunc.Open "SELECT [CSToDinersID] FROM tblCAPSCSToDiners WITH(NOLOCK) WHERE EIDNo = '" & strEmplID & "' AND [Status]='Awaiting Export'",objCon

		If objRSFunc.EOF Then
			GetEIDOnCSFile = "N" 
		Else
			GetEIDOnCSFile = "Y"
		End If

	objRSFunc.Close
 
Set objRSFunc = Nothing
 
End Function


Public Sub SetSessionEIDFromTwoFA(strTwoFAID)
'Function to check whether an EmployeeID is on the CS file for today and returning a Y or N
Dim objRSFunc

Set objRSFunc = Server.CreateObject("ADODB.Recordset")

	objRSFunc.Open "SELECT [EmployeeID] FROM qryPORTALCards WITH(NOLOCK) WHERE TwoFactorID Like '" & strTwoFAID & "'",objCon

		If objRSFunc.EOF Then
			Session("ApplicationEmployeeID") = "" 
		Else
			Session("ApplicationEmployeeID") = objRSFunc("EmployeeID") 
		End If

	objRSFunc.Close
 
Set objRSFunc = Nothing
 
End Sub
%>