<!DOCTYPE html>
<html lang="en">
<%

Dim objCon
Dim objRS
Dim objRS1
Dim x
Dim strBatchNo
Dim strEmails
Dim strSQL

'Open Database Connection
Set objCon = Server.CreateObject("ADODB.Connection")
    
    objCon.Open Session("DBConnection")

Set objRS = Server.CreateObject("ADODB.Recordset")
Set objRS1 = Server.CreateObject("ADODB.Recordset")

If Not IsEmpty(Request.QueryString("FileSeqNum")) Then
	'strBatchNo = " WHERE FileSeqNum Like '%" & Request.QueryString("FileSeqNum") & "%' AND [FileType] = 'CSFromDiners' AND [Deleted] = 'N'"
End If

strBatchNo = " WHERE AlreadyUnactivated Is Null"

Public Sub ShowDetails()

Dim strWarning
Dim strSelected

	objRS.Open "SELECT * FROM tblCAPSEmailDetail WITH(NOLOCK) ",objCon
		
	If Not objRS.EOF Then
		   
		Do Until objRS.EOF
		
			If objRS("EmailDetailID") = 24 Then
				strSelected = " SELECTED "
			Else
				strSelected = ""
			End If
			
			strEmails = strEmails & "<option id=""" & objRS("EmailDetailID") & """ value=""" & objRS("EmailDetailID") & """ " & strSelected & " >" & objRS("EmailTemplateName") & "</option>"
			
		objRS.Movenext
		Loop
	Else
		strEmails = "No email Templates available"
	End If
	
	objRS.Close
	
	'Set the Warning Select
	strWarning = "<option id=""N"" value=""N"">N</option>"
	
	'strSQL = "qryCAPSUnactivatedCards WITH(NOLOCK) WHERE [ActivationFlag] = 'N' AND (DATEDIFF(day, DateIssued, GETDATE()) > 45) ORDER By [DateIssued]"
	
	'Old Batch Statement
	'strSQL = "SELECT * FROM qryCAPSCardDaysIssued " & strBatchNo
	
	'Description:	Loads Position details into page if applicable.
	objRS.Open "SELECT * FROM qryCAPSCardDaysIssued " & strBatchNo,objCon

		Response.write "<div class=""card mb-3""><div class=""card-body""><div class=""table-responsive"">" & _
            "<table class=""table table-bordered table-hover header-fixed"" id=""dataTable"" width=""100%"" cellspacing=""0"">" & _
			"<thead><tr><th>Batch No</th><th>Email Template</th><th>Number of Cards in Batch</th></tr>"
			 
		'''''This was removed to avoid complication and confusion so now (above) does not include Number of days to select and Warning Yes or No'''''
		'Response.write "<div class=""card mb-3""><div class=""card-body""><div class=""table-responsive"">" & _
         '   "<table class=""table table-bordered table-hover header-fixed"" id=""dataTable"" width=""100%"" cellspacing=""0"">" & _
		'	"<thead><tr><th>Batch No</th><th>Email Template</th><th>Days Inactive</th><th title=""Select Y to send an email those who already have a Warning/Email sent to them"">Warning</th><th>Number of Cards in Batch</th></tr>"

			
		If Not objRS.EOF Then
		   
				
				Response.write "<tr><td style=""text-align:center;"">New</td>" & _
					"<td style=""text-align:center;""><select id=""SelEmail"" name=""SelEmail""><option id=""0"" value=""0"">Select an Email Template</option>" & strEmails & "</select></td>" & _
					"<td style=""text-align:center;"">" & objRS("DaysSinceIssued") & "</td></tr>"
			
			'''''This was removed to avoid complication and confusion so now (above) does not include Number of days to select and Warning Yes or No'''''
				'Response.write "<tr><td style=""text-align:center;"">New</td>" & _
				'	"<td style=""text-align:center;""><select id=""SelEmail"" name=""SelEmail""><option id=""0"" value=""0"">Select an Email Template</option>" & strEmails & "</select></td>" & _
				'	"<td><input id=""DaysEmail"" name=""DaysEmail"" value=""45"" type=""number""></td>" & _
				'	"<td style=""text-align:center;""><select id=""SelWarning"" name=""SelWarning""><option id=""Y"" value=""Y"">Y</option>" & strWarning & "</select></td>" & _
				'	"<td style=""text-align:center;"">" & objRS("DaysSinceIssued") & "</td></tr>"
					
			'"<td style=""text-align:center;""><select id=""SelEmail"" name=""SelEmail""><option id=""0"" value=""0"">Select an Email Template</option>" & strEmails & "</select></td>"&_
			
		Else
			Response.write "There are no Unactivated Cards greater than 45 days unactivated as at today (" & now() & "). Try creating a New Batch in another day or two."
	   End If

	objRS.Close
	
	Response.write "</table></div></div></div>"
End Sub


Public Sub GetEmailTemplates()

objRS1.Open "SELECT * FROM tblCAPSEmailDetail WITH(NOLOCK) ",objCon
		
	If Not objRS1.EOF Then
		   
		Do Until objRS1.EOF
		
			strEmails = strEmails & "<option id=""" & objRS1("EmailDetailID") & """ value=""" & objRS1("EmailDetailID") & """>" & objRS1("EmailTemplateName") & "</option>"
			
		objRS.Movenext
		Loop
	Else
		strEmails = "No email Templates available"
	End If
	
	objRS.Close

End Sub

%>
<head>
  
</head>

<body>
  <%
  Call ShowDetails()
  
  %>
</body>

</html>
<%

Set objRS = Nothing
Set objCon = Nothing
  
  %>