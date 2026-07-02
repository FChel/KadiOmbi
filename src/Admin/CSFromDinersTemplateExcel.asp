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
Dim objRS1
Dim arrHeadings(5)
Dim intFinYearPart1
Dim intFinYearPart2
Dim strBackColour
Dim dblDisplay
Dim dblActual
Dim dblBudget
Dim dblVariance
Dim dblVariancePercentage
Dim strForeColour
Dim intMode
Dim dblTotal
Dim dblTotal1
Dim dblOriginal
Dim dblVarianceTotal
Dim dblVariancePercentageTotal
Dim dblOriginalTotal
Dim dblVarianceTotal1
Dim dblVariancePercentageTotal1
Dim x
Dim strMessage
Dim strLevelName
Dim strSQL
Dim dblVar
Dim strCostCentreName
Dim strVersionName
Dim arrMonthName(12)
Dim strData

Set objCon = Server.CreateObject("ADODB.Connection")
Set objRS = Server.CreateObject("ADODB.Recordset")
Set objRS1 = Server.CreateObject("ADODB.Recordset")

objCon.Open Session("DBConnection")	

'Get Level1ID

If Not IsEmpty(Request.QueryString("Level1ID")) Then
	Session("Level1ID") = Request.QueryString("Level1ID")
End If

If Not IsEmpty(Request.QueryString("TransactionType")) Then
	Session("TransactionType") = Request.QueryString("TransactionType")
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

	If Session("EmployeeID") = 0 Then
	'strSQL = "SELECT * FROM qryApplications WHERE EmployeeID = " & Session("EmployeeID") & ""
	strSQL = "SELECT * FROM qryCSFromDIners " & strWhere
	Else
		strSQL = "SELECT * FROM qryCSFromDIners " & strWhere
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
