<%@ Language=VBScript %>
<!-- #Include file=../ADOVBS.inc -->
<%
'Description:	Build an Excel worksheet from the dataset parameters passed in
'Author:		Michael Giacomin
'Date:			May 2020

Dim strWorksheetName

	'If the session has expired then send the user back to the Default/login page
	If IsEmpty(Session("UserID")) Then Response.Redirect("../Timeout.asp")
	
	Response.Expires = -1500
	
	'Get the table/query name to use as the spreadsheet name, if one has been sent through the querystring
	If Not IsEmpty(Request.QueryString("tbl")) Then
		strWorksheetName = cstr(Request.QueryString("tbl"))
		
		'Remove the qry or tbl part of the table name to make it more appropriate for the Workcheet name
		If Len(strWorksheetName)>3 Then
			If Left(strWorksheetName,3)="qry" OR Left(strWorksheetName,3)="tbl" Then
				strWorksheetName = Right(strWorksheetName,Len(strWorksheetName)-3)
			End If
		End If
	Else
		strWorksheetName = "ANZCancelCards"
	End If

	'******Change 18/12/2013
    'The browser now determines the header details as IE 11 behaves differently
    If Session("UBrowser") = "FF" Then
        Response.ContentType = "application/vnd.ms-excel"
    Else
        Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    End If
    
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment; filename=" & strWorksheetName & ".xls" 
	
'Instantiate Common Page Variables.
Dim objCon
Dim objCmd
Dim objRS

Dim strEID
Dim strTop
Dim strTable
Dim strWhere

'Open Database Connection
Set objCon = Server.CreateObject("ADODB.Connection")
Set ObjCmd = Server.CreateObject("ADODB.Command")
Set objRS = Server.CreateObject("ADODB.Recordset")

objCon.Open Session("DBConnection")

If Not IsEmpty(Request.QueryString("EIDNo")) Then

	strEID = Request.QueryString("EIDNo")
End If

If Not IsEmpty(Request.QueryString("tbl")) Then

	strTable = cstr(Request.QueryString("tbl"))
End If

If Not IsEmpty(Request.QueryString("W")) Then

	strWhere = cstr(Request.QueryString("W"))
End If

strTop = ""

If Not IsEmpty(Request.QueryString("Top")) Then

	strTop = cstr(Request.QueryString("Top"))
	If IsNumeric(strTop) Then
		strTop = " TOP " & strTop
	End If
End If

'Response.ContentType = "application/vnd.ms-excel"
'Response.AddHeader "Content-Disposition", "attachment; filename=" & strTable & ".xls" 
	
'response.write "strTable=" & strTable & " w=" & strWhere

 %>

<html>
<head>

<style>

    table.newd th, table.newd td{

        padding: 4px; 

    }

</style>
</head>
<body>

		
				<%
        
      DisplayTableDetails()
        
%>	

        
    
</body>
</html>

<%
Sub DisplayTableDetails()

Dim fldField
Dim strHeader
Dim strDetail
Dim x
Dim intFields
Dim strColour
Dim strColourFore
Dim intFieldCount

	'If Session("EmployeeID") = 0 Then
		'strSQL = "SELECT * FROM qryApplications WHERE EmployeeID = " & Session("EmployeeID") & ""
	'	strSQL = "SELECT * FROM tblCSToDiners " & strWhere
	'Else
	If Not IsEmpty(Request.QueryString("SQL")) Then
		strSQL = Session("SQL")
	Else
		strSQL = "SELECT " & strTop & " * FROM " & strTable & " WITH(NOLOCK) " & strWhere
		'strSQL = "SELECT * FROM qryApplications WHERE EmployeeID = " & Session("EmployeeID") & ""
	End If

'response.write "SQ=" & strSQL

'Response.Write strSQL

	intFields = 0
	
	Response.Write"<table BORDER=""1"" cellspacing=""1"" cellpadding=""1"">"

	objRS.Open strSQL,objCon
		
		If NOT objRS.Eof Then
		
			For each fldField in objRS.fields
			
				strHeader = strHeader & "<th>" & fldField.name & "</th>"
			
			Next
		
			strHeader = strHeader & "</tr>"
			
		End If
		
		Do until objRS.EOF
			
			strDetail = strDetail & "<tr>"
			intFieldCount = 0
			
			For each fldField in objRS.fields
			
				intFieldCount = intFieldCount + 1
				'Temporary for the Training Report to show formatting (change cell colours)
				If strTable = "qryCAPSTrainingReport" AND intFieldCount = 33 Then
					strDetail = strDetail & "<td style=""color:white; background-color:blue;"">" & fldField.value & "</td>"
					
				ElseIf strTable = "qryCAPSTrainingReport" AND intFieldCount = 32 Then
					strDetail = strDetail & "<td style=""color:red; font-weight:bold;"">" & fldField.value & "</td>"
				Else
					strDetail = strDetail & "<td>" & fldField.value & "</td>"
				End If
				'Only count the fields if this is the first record
				If intFields = 0 Then x = x + 1
			next
			
			strDetail = strDetail & "</tr>"
			
			objRS.movenext
			
			intFields = intFields + 1
			
		Loop
			
	objRS.Close

        strDetail = strDetail & "</table>"
		
		'Go back and add the first row in after counting the number of fields (columns)
		strHeader = "<tr><th colspan=""" & x & """>" & strTable & " " & strWhere & "</th></tr>" & strHeader
				
		
		'Write the header and Detail to the screen
		Response.Write strHeader & strDetail
		
End Sub




Set objRS = Nothing
Set objCon = Nothing

 %>


