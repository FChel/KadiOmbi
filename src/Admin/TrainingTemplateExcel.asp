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
	Response.AddHeader "Content-Disposition", "attachment; filename=TrainingData.xls" 
 
	
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
	
	<%Response.Write "<tr><th>Course ID</th>" & _	
		        "<th>Offering ID</th>" & _
	 	        "<th>Course Title</th>" & _	
	 	        "<th>PMKeyS/ODS ID</th>" & _
		        "<th>First Name</th><th>Last Name</th>" & _
                "<th>Email</th><th>Completion Date</th></tr>"
	 %>
	 
	
</table>
</body>
</html>
