<%@ Language=VBScript %>
<%
Option Explicit

	Response.Expires = -1500
	
	'******Change 18/12/2013
    'The browser now determines the header details as IE 11 behaves differently
    If Session("UBrowser") = "FF" Then
        Response.ContentType = "application/vnd.ms-excel"
    Else
        Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    End If
    
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment; filename=CSFromDinersFile.xls" 
 
	
Dim objCon
Dim objRS
Dim x
Dim strSQL
Dim strData

Set objCon = Server.CreateObject("ADODB.Connection")
Set objRS = Server.CreateObject("ADODB.Recordset")


objCon.Open Session("DBConnection")	

'Get any parameters passed when calling this page
If Not IsEmpty(Request.QueryString("Data")) Then
	strData = Request.QueryString("Data")
End If


strData = ""

%>

<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</head>
<body>

<table WIDTH="1560px" BORDER="1" CELLSPACING="1" CELLPADDING="1">
	
	<%
	
Dim strWhere
Dim fldField

'If Session("Filter") = "GCFO" Then
'	strWhere = "WHERE [GCFOSigned] IS NULL"

'End If

	If Session("EmployeeID") = "" Then
		'strSQL = "SELECT * FROM qryApplications WHERE EmployeeID = " & Session("EmployeeID") & ""
		strSQL = "SELECT * FROM qryCAPSCSFromDiners " & strWhere
	Else
		strSQL = "SELECT * FROM qryCAPSCSFromDiners " & strWhere
		'strSQL = "SELECT * FROM qryApplications WHERE EmployeeID = " & Session("EmployeeID") & ""
	End If

	objRS.Open strSQL,objCon
		
		x = 1
		
		'If NOT objRS.Eof Then
		
			
			Response.write "<tr Style=""font-weight:bold;"">"
		
			For each fldField in objRS.fields
			
				Response.write "<th>" & fldField.name & "</th>"
			
			next
		
		Response.write "</tr>"
			
		
		'End If
		
		If strData = "Yes" Then
			Do until objRS.EOF 
			
				
				Response.write "<tr>"
				
				For each fldField in objRS.fields
				
					Response.write "<td>" & fldField.value & "</td>"
				
				next
				
				Response.write "</tr>"
				
				objRS.movenext
			Loop
		
		end if
		
						
objRS.Close

	 %>
	 
	
</table>
</body>
</html>
