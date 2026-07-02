<%@ Language=VBScript %>
<!-- #include file="../upload.asp" -->
<!-- #Include file=../../ADOVBS.inc -->
<%
'Description:	Genera Expenses Upload Administration screen
'Author:		MG
'Date:			April 2013

    'If the session has expired then send the user back to the Default/login page
	If IsEmpty(Session("DBConnection")) Then
		Response.Redirect "../../Default.asp?State=Expired"
	End If
	
'Instantiate Common Page Variables.
Dim objCon
Dim objCmd
Dim objRS
Dim arrMonthName(12)

'Open Database Connection
Set objCon = Server.CreateObject("ADODB.Connection")
Set ObjCmd = Server.CreateObject("ADODB.Command")
Set objRS = Server.CreateObject("ADODB.Recordset")

objCon.Open Session("DBConnection")

 %>

<html>
<head>
<link rel="stylesheet" type="text/css" href="../../BertStyle.css">
<script language=javascript>
function upload()
{   
    if(document.getElementById('FILE1').value=="")
    {
        alert("Please select the file to upload (click the Browse button above)");
    }
    else
    {   
        if(getFileExt(document.getElementById('FILE1').value)==".xls")
        {
            if(window.confirm('This will overwrite any existing CS From Diners file data! \n \n Continue?')==true)
                {
                document.getElementById('Progress').style.display = "inline";
                frm.submit();
            }
        } 
        else
        {
            alert("Please enter a valid Excel(.xls) file");
        }
    }
}

 function getFileExt(filename)
 {
     var s
     s = filename.charAt(filename.length-4) + filename.charAt(filename.length-3) + filename.charAt(filename.length-2)+ filename.charAt(filename.length-1);
     return s;
}

</script>
</head>
<body>

<form action="UploadCS.asp?Action=Save&chkDelete="  method="POST" enctype="multipart/form-data" id="frm" name="frm">
<br /><br />

<table BORDER="1" CELLSPACING="1" CELLPADDING="1" width="100%">
<tr>
    <th width=100% align="left" style="font-size:18px;">CS Data Upload Screen
   
</th>
</tr>
</table>

<br />
<div class="buttons">
<table BORDER="1" CELLSPACING="1" CELLPADDING="1" width="100%">
<tr>
    <th align=left width="180"><b>Select a file to upload</b></th>
    <td><INPUT TYPE=FILE SIZE=50 NAME="FILE1" id="FILE1"></td>
</tr>
<tr>
    <td><input type="checkbox" id="chkDelete" name="chkDelete" title="Check to OVERWRITE all Uploaded Employee values"/> <b>Overwrite Zero Values</b></td>
    <td><button type="button" onclick="upload()";><img src="../../images/disk.png" alt="" /> Upload </button></td>
</tr>
<tr><td></td>
<td><span id="Progress" style="display:none"><img src="../../Images/progress.gif" />  &nbsp;&nbsp;&nbsp; <b>Loading...</b></span></td></tr>
</table>
<br>
<font color=red size=2><b>NOTE: </Font><font color=black size=2>When loading from Excel the Worksheet MUST be named 'CSData' (no spaces)
            <BR>* The file must be '.xls' only
            <BR>* Do not change the first row headers from the template files (below)</B></Font>
<br><br>
<table Width="800px" BORDER="0" CELLSPACING="0" CELLPADDING="0">
<tr>
	<td Width="100px"><button type="button" onclick="parent.location='AdminFrameset.asp';"><img src="../../images/door.png" alt="" /> Close </button></td>
</tr>
</table><br>

<TABLE Width="800px" BORDER="0" CELLSPACING="0" CELLPADDING="0">
<TR>
    <td Width="200px"><button type="button" onclick="window.open('CSFromDinersTemplateExcel.asp')"><img src="../../images/page_excel.png" title="Click here to get an Excel Template for loading the CS FRom Diners file" /> CS From Diners Template Excel </button></td>
</TR>
</TABLE>
</div>

<BR>

<% 
   

        objRS.Open "SELECT TOP 50 * FROM qryCAPSCSFromDiners ",objCon
		
		    If objRS.eof Then
		        Response.Write"<table WIDTH=""100%"" BORDER=""1"" CELLSPACING=""1"" CELLPADDING=""1"">" & _
                "<tr><th colspan=7 style=""text-align:left"">There is no CS From Diners data loaded</B></th></tr>" & _
		        "<tr><td colspan=7 style=""text-align:left"">Okay to Uplod Data</td></tr>"
		    Else
		    
		         Response.Write"<table WIDTH=""100%"" BORDER=""1"" CELLSPACING=""1"" CELLPADDING=""1"">" & _
                "<tr><th colspan=15 style=""text-align:left"">Sample of Existing CS From Diners Data already in CAPS2</th></tr>" & _
                "<tr><td colspan=15 style=""text-align:left; color:red;font-size:20px""><B>WARNING! The data below will be deleted if you Upload a new CS From Diners file!</B></td></tr><tr>" & _
		        "<th>ApplicationID</th>" & _
				"<th>EmployeeID</th>" & _
				"<th>Card No</th><th>Card Type</th>" & _	
		        "<th>Title</th>" & _
	 	        "<th>FirstName</th>" & _	
	 	        "<th>Surname</th>" & _
		        "<th>Address1</th><th>Address2</th>" & _
                "<th>Address3</th><th>Suburb</th>" & _
                "<th>State</th><th>PostCode</th><th>EmailAddress</th>" & _
                "<th>Status</th></th></tr>"
		        
		    End If
		    
		    Do until objRS.eof
		        'If objRS("GLCode") <> 0 Then
			    Response.Write "<TR><TD>&nbsp;" & objRS(0) & "</A></B></TD>" & _
			                    "<TD style=""text-align:center"">" & objRS(1) & "</TD><TD style=""text-align:center"">" & objRS(5) & "</TD><TD style=""text-align:center"">" & objRS(3) & " " & objRS(4) & "</TD>" & _
			                    "<TD style=""text-align:center"">" & objRS(6) & "</TD><TD style=""text-align:center"">" & objRS(7) & "</TD><TD style=""text-align:center"">" & objRS(8) & "</TD>" & _
			                    "<TD style=""text-align:center"">" & objRS(9) & "</TD><TD style=""text-align:center"">" & objRS(10) & "</TD><TD style=""text-align:center"">" & objRS(11) & "</TD>" & _
			                    "<TD style=""text-align:center"">" & objRS(12) & "</TD><TD style=""text-align:center"">" & objRS(13) & "</TD><TD style=""text-align:center"">" & objRS(14) & "</TD>" & _
								"<TD style=""text-align:center"">" & objRS(15) & "</TD><TD style=""text-align:center"">" & objRS(16) & "</TD>" & _
			                    "</TR>"
    			'End If
    			
			    objRS.movenext
		    Loop
    			
	    objRS.Close

        Response.Write "</table>"

%>
</form>

<%
'On Error Resume Next 
Dim objExcelCon
Dim strUploadPath

dim errors
dim lineNo
dim GeneralExpenseID
dim BudgetID
dim VersionID
dim CostCentreID
dim TransactionType
dim GLCode
dim BM1
dim BM2
dim BM3
dim BM4
dim BM5
dim BM6
dim BM7
dim BM8
dim BM9
dim BM10
dim BM11
dim BM12
dim OY1
dim OY2
dim OY3
dim Comments
dim updatedBy 

Dim strEmployeeID
Dim strTitle
Dim strFirstName
Dim strSurname
Dim strAddress1
Dim strAddress2
Dim strAddress3
Dim strSuburb
Dim strState
Dim strPostCode
Dim strStatus
Dim strUpdatedBy
Dim strEmailAddress

Dim dblCreditLimit
Dim dteDateReceived
Dim strReviewedBy
Dim dteDateReviewed
Dim lngBatchNo
Dim strCardNo
Dim strCardType
Dim strCardTypeSub
				
errors = "" 
lineNo = 1    
strUpdatedBy = Session("UserID")

Dim Uploader, File, filePath

Set Uploader = New FileUploader

' This starts the upload process
Uploader.Upload()

if request.QueryString("Action")="Save" Then

	'If Session("StatusID") <> 1 Then
   
    '    Response.write "<Font Color=red><B>Budget is closed, no changes can be made!</B></FONT>"
	'Else
	' Check if any file is sucessfully uploaded
	    If Uploader.Files.Count = 0 Then
		    Response.Write "File(s) not uploaded."
	    Else
		    ' Loop through the uploaded file
		    For Each File In Uploader.Files.Items					  		    		    
			    'set the upload path
	            strUploadPath = Server.MapPath(GetFilePath()) & "\Attachments"
				File.SaveToDisk strUploadPath
			    filePath = Server.MapPath(GetFilePath()) & "\Attachments\" & File.FileName
		    Next
		    
		    'After uploading, Read excel file
		    
		    Set objExcelCon = Server.CreateObject("ADODB.connection")     
	        objExcelCon.Open "DBQ=" & filePath & "; DRIVER={Microsoft Excel Driver (*.xls)};" 
	   
	        'Write the SQL Query 
	        objRS.open "SELECT * FROM [CSFromDinersFile$]", objExcelCon  
		           
		    ReadExcel 
		    
		    'Check for errors
	        if(errors<>"") then
	            'Print the errors and return
	            Response.Write "<font face=arial size=1><b>File not uploaded. Please correct the followings errors and try again</b> <br> "& errors &"</font><br>"         
	        else  
	            'If the no errors found in the ReadExcel method, then start uploading the records in the database       	    	    
		        UploadExcel Uploader.Form("chkDelete")
		        Response.Write "<b>CS From Diners File Sucessfully Uploaded!!<b><br>"
		    End if
		    	    	                
	    'Close the recordset/connection 
	    objRS.Close 
	    objExcelCon.Close 
		    
	    'End If
	End IF
	
End if

'Function for validating the excel file values
Sub ReadExcel 

    'First loop to check all the values
    Do until objRS.EOF 
        'Read each recors present in the Excel file and check for the validation
 
                 
        strEmployeeID = objRS("EmployeeID") 'Should be integer and Not Null values       
        If IsNull(strEmployeeID) Then
            errors = errors & "Error in line no. " & lineNo & ": Employee ID should NOT be Null value <br>" 
        End if
        
		strEmailAddress = objRS("EmailAddress") 'Should be integer and Not Null values       
        If IsNull(strEmailAddress) Then
            errors = errors & "Error in line no. " & lineNo & ": Email Address should NOT be Null value <br>" 
        End if
       
	   strSurname = objRS("Surname") 'Should be integer and Not Null values       
        If IsNull(strSurname) Then
            errors = errors & "Error in line no. " & lineNo & ": Surname/LastName should NOT be Null value <br>" 
        End if
                          
            lineNo = lineNo + 1            
            
        objRS.movenext 
    
        If IsNull(objRS("EmployeeID")) AND IsNull(objRS("EmployeeID")) Then       
        
            exit Sub

        End If
    
    Loop        
    
End Sub	

'Function for validating the excel file values
Sub UploadExcel(chkDelete)

Dim x
Dim intTotal
  
    objRS.MoveFirst()
    
    Do until objRS.EOF 
        
		strApplicationID = objRS("ApplicationID")
        strEmployeeID = objRS("EmployeeID") 
        
		strCardNo = objRS("CardNo") 
		strCardType = objRS("CardType")
		strCardTypeSub = objRS("CardTypeSub")
        strTitle = objRS("Title") 
        strFirstName = objRS("FirstName") 
		strSurname = objRS("Surname") 
		strAddress1 = objRS("Address1") 
		strAddress2 = objRS("Address2") 
		strAddress3 = objRS("Address3") 
		strSuburb = objRS("Suburb") 
		strState = objRS("State") 
		strPostCode = objRS("PostCode") 
		strEmailAddress = objRS("EmailAddress") 
		strStatus = objRS("Status") 
		strUpdatedBy = Session("UserID") 
        
		dblCreditLimit = objRS("CreditLimit") 
		dteDateReceived = objRS("DateReceived") 
		strReviewedBy = objRS("ReviewedBy") 
		dteDateReviewed = objRS("DateReviewed") 
		lngBatchNo = objRS("BAtchNo") 


        'Only save the record if the DeleteCheckbox is checked otherwise it will overwrite zero values and clear existing data
        If Cstr(Trim(chkDelete)) = " " AND intTotal = 0 Then
        
        Else
            x = x + 1
            'Save the record
            SaveRecord strApplicationID,strEmployeeID,strCardNo,strCardType,strCardTypeSub,strTitle,strFirstName,strSurname,strAddress1,strAddress2,strAddress3,strSuburb,strState,strPostCode,strEmailAddress,strStatus,dblCreditLimit,dteDateReceived,strReviewedBy,dteDateReviewed,lngBatchNo,strUpdatedBy, x
            'response.write  "exec spGeneralExpensesSave ="& CostCentreID & "," & GLCode & "," & BM1 & "," & BM2 & "," & BM3 & "," & BM4 & "," & BM5 & "," & _
            '                BM6 & "," & BM7 & "," & BM8 & "," & BM9 & "," & BM10 & "," & BM11 & "," & BM12 & "," & OY1 & "," & OY2 & "," & OY3 & "," & Comments & "," & x
        
        End IF
        
     objRS.movenext 
    
    If IsNull(objRS("EmployeeID")) Then               
       
       exit Sub
       
    End If
        
    Loop 
    
End Sub	


Sub SaveRecord(strApplicationID,strEmployeeID,strCardNo,strCardType,strCardTypeSub,strTitle,strFirstName,strSurname,strAddress1,strAddress2,strAddress3,strSuburb,strState,strPostCode,strEmailAddress,strStatus,dblCreditLimit,dteDateReceived,strReviewedBy,dteDateReviewed,lngBatchNo,strUpdatedBy, x)

Dim intRecord

  	With objCmd
  	
  	    'If the procedure has akready run then don't create the parameter objects again (more than once)
  	    If x = 1 then
                .CommandType = 4
                .CommandText = "spCSFromDinersSave"
                
				.Parameters.Append objCmd.CreateParameter("CSFromDinersID", adInteger) 
				.Parameters.Append objCmd.CreateParameter("ApplicationID", adInteger) 
				.Parameters.Append objCmd.CreateParameter("CardNo", adVarChar, adParamInput,50) 
				.Parameters.Append objCmd.CreateParameter("CardType", adVarChar, adParamInput,20) 
				.Parameters.Append objCmd.CreateParameter("CardTypeSub", adVarChar, adParamInput,20)
                .Parameters.Append objCmd.CreateParameter("EmployeeID", adVarChar, adParamInput,20)
                .Parameters.Append objCmd.CreateParameter("Title", adVarChar, adParamInput,10)
				.Parameters.Append objCmd.CreateParameter("FirstName", adVarChar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("Surname", adVarChar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("Address1", adVarChar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("Address2", adVarChar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("Address3", adVarChar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("Suburb", adVarChar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("State", adVarChar, adParamInput,10)
				.Parameters.Append objCmd.CreateParameter("PostCode", adVarChar, adParamInput,10)
				.Parameters.Append objCmd.CreateParameter("EmailAddress", adVarChar, adParamInput,100)
				.Parameters.Append objCmd.CreateParameter("Status", adVarChar, adParamInput,20)
				.Parameters.Append objCmd.CreateParameter("CreditLimit", adInteger)
				.Parameters.Append objCmd.CreateParameter("DateReceived", adDate)
				.Parameters.Append objCmd.CreateParameter("ReviewedBy", adVarChar, adParamInput,20)
				.Parameters.Append objCmd.CreateParameter("DateReviewed", adDate)
				.Parameters.Append objCmd.CreateParameter("BatchNo", adInteger)
				.Parameters.Append objCmd.CreateParameter("UpdatedBy", adVarChar, adParamInput,10)
                .Parameters.Append objCmd.CreateParameter("CSFromDinersIDOutput", adInteger, adParamOutput)                         
            
        End If
                 
				.Parameters("CSFromDinersID") = 0		
				.Parameters("ApplicationID") = strApplicationID
				.Parameters("CardNo") = strCardNo
				.Parameters("CardType") = strCardType
				.Parameters("CardTypeSub") = strCardTypeSub
				.Parameters("EmployeeID") = strEmployeeID	
				.Parameters("Title") = strTitle					
                .Parameters("FirstName") = strFirstName
                .Parameters("Surname") = strSurname          
                .Parameters("Address1") = strAddress1
                .Parameters("Address2") = strAddress2
                .Parameters("Address3") = strAddress3
                .Parameters("Suburb") = strSuburb
                .Parameters("State") = strState
                .Parameters("PostCode") = strPostCode
				.Parameters("EmailAddress") = strEmailAddress
                .Parameters("Status") = strStatus
				.Parameters("CreditLimit") = dblCreditLimit
				.Parameters("DateReceived") = dteDateReceived
				.Parameters("ReviewedBy") = strReviewedBy
				.Parameters("DateReviewed") = dteDateReviewed
				.Parameters("BatchNo") = lngBatchNo
                .Parameters("UpdatedBy") = strUpdatedBy
                           
                .ActiveConnection = objCon
                
            End With
                
            objCmd.Execute        
            
            'Return the result of the Save Function.
     		intRecord = objCmd.Parameters.Item("CSFromDinersIDOutput")    
     		                  			     				     		     		
       'response.write  "exec spGeneralExpensesSave =0," & Session("BudgetID") & "," & Session("VersionID") & "," & CostCentreID & ",'GEXP'" & GLCode & "," & BM1 & "," & BM2 & "," & BM3 & "," & BM4 & "," & BM5 & "," & _
       '                     BM6 & "," & BM7 & "," & BM8 & "," & BM9 & "," & BM10 & "," & BM11 & "," & BM12 & "," & OY1 & "," & OY2 & "," & OY3 & ",'" & Comments & "','" & UpdatedBy & "'," & Session("ColumnLock")
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

Public Sub GetMonthNames()
'This is a procedure to get the order of Month names to be used as titles for Month Columns
Dim intFirstMonth

    'set the First Month name to an integer
    'intFirstMonth = Month("21-" & Session("FirstMonth") & "-2012")
	intFirstMonth = Month("21-1-2012")
    'intFirstMonth = intFirstMonth -1
    arrMonthName(0) = intFirstMonth
    For x = 1 to 12
    
        arrMonthName(x) = Left(MonthName(intFirstMonth + x - 1),3)'intFirstMonth + x
  '      arrMonthName(x) = intFirstMonth'MonthName(intFirstMonth)
        
        'Once the count goes over 12 then go back to 1 to fill the remaining months
        If intFirstMonth + x - 1 > 11 Then 
            If intFirstMonth > 6 Then
                intFirstMonth = 2 - intFirstMonth
            Else
                intFirstMonth = (x - 1) * - 1
            End If
        End If
        
  '      intFirstMonth = x
        
    Next
    
End Sub

Set objRS = Nothing
Set objCon = Nothing

 %>
 </body>
</html>

