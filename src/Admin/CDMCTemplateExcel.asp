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
	Response.AddHeader "Content-Disposition", "attachment; filename=CDMCData.xls" 
 
	
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

%>

<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</head>
<body>

<table WIDTH="1560px" BORDER="1" CELLSPACING="1" CELLPADDING="1">
	
	<%Response.Write "<tr><th>EmployeeID</th>" & _	
		        "<th>Title</th>" & _
	 	        "<th>FirstName</th>" & _	
	 	        "<th>Surname</th>" & _
		        "<th>Address1</th><th>Address2</th>" & _
                "<th>Address3</th><th>Suburb</th>" & _
                "<th>State</th><th>PostCode</th>" & _
				"<th>EmailAddress</th>" & _
                "<th>Status</th><th>UpdatedBy</th>" & _
                "<th>DateUpdated</th></tr>"
	 %>
	 
	
</table>
</body>
</html>
