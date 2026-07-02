<%@ Language=VBScript %>
<% Option Explicit %>
<!-- #Include file=../ADOVBS.inc -->

<%
 
'Description:	Student Screen
'Author:		Andrew Bull - Isidore Limited (www.isidore.com - 07969 589413)
'Date:			August 2004

'Declare default variables

Dim objCon
Dim objCmd
Dim objRS
Dim objRS1
Dim strScreen
Dim strSelected
Dim x 
Dim strMessage
Dim strColour

'Set Database objects

Set objCon = Server.CreateObject("ADODB.Connection")
Set objCmd = Server.CreateObject("ADODB.Command")
Set objRS = Server.CreateObject("ADODB.Recordset")
Set objRS1 = Server.CreateObject("ADODB.Recordset")

'Open database connection

objCon.Open Session("DBConnection")



'1. Declare screen specific variables naming convention = variable type prefix + database field name

Dim lngBudgetID
Dim lngCostCentreID
Dim lngStatusID
Dim strBudgetName
Dim intFinancialYearID
Dim intDefaultVersionID
Dim strBalanceSheet
Dim lngCashFlowGLCode
Dim lngBadDebtGLCode
Dim lngPrePaymentGLCode
Dim lngInvestmentGLCode
Dim lngLoanGLCode
Dim lngAPGLCode
Dim lngARGLCode
Dim intVariancePercentage
Dim strComments
Dim strActive

Dim dblRevenue
Dim dblExpenditure
Dim dblNetTotal

Dim dblApprovedRev
Dim dblApprovedExp
Dim dblApprovedNet

Dim dblTotalRev
Dim dblTotalExp
Dim dblTotalNet

Dim dblOpexRev
Dim dblOpexExp
Dim dblOpexNet

Dim dblCapxRev
Dim dblCapxExp
Dim dblCapxNet
Dim dblEstimate
Dim dblEstimateTotal

'Declare and set default arrays

Dim arrYesNo(2)
	
	arrYesNo(1) = "Y"
	arrYesNo(2) = "N"
	
	'3. Capture Querystring variables
	
	If Not IsEmpty(Request.QueryString("CostCentreID")) Then
		
		lngCostCentreID = Request.QueryString("CostCentreID")
						
	End If
	
	If Not IsEmpty(Request.QueryString("StatusID")) Then
		
		lngStatusID = Request.QueryString("StatusID")
					
	End If	
	
    
Dim arrStatus(7)

arrStatus(0) = "<IMG SRC='../images/delete.png'"
arrStatus(1) = "<IMG SRC='../images/open.png'"
arrStatus(2) = "<IMG SRC='../images/ready.gif'"
arrStatus(3) = "<IMG SRC='../images/cross.png'" 
arrStatus(4) = "<IMG SRC='../images/tick.png'"	
arrStatus(5) = "<IMG SRC='../images/Closed.png'"
arrStatus(6) = "<IMG SRC='../images/bug.png'"   
arrStatus(7) = "<IMG SRC='../images/wrench.png'"
 			
%>

<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="../BERTStyle.css">

	
</head>

<body>
<form action="CostCentreStatus.asp?Action=Save" method="POST" id="frm" name="frm">


<table WIDTH="100%" BORDER="1" CELLSPACING="1" CELLPADDING="1">
 
	<tr><th colspan="4" Style="Height:25px;font-size:12px;font-weight:bold;">Applications</td></tr>
	<tr>
		<th Style="Height:25px; font-size:12px;font-weight:bold;" Width="20%">Card Type</th>
		<th Style="Height:25px; font-size:12px;font-weight:bold;" Width="40%">Applicant</th>
        <th Style="Height:25px;font-size:12px;font-weight:bold;" Width="25%">Status</th>
		<th Style="Height:25px;font-size:12px;font-weight:bold;" Width="15%">Received</th>
	</tr>
	
<%

Dim y

Dim strCardType, strApplicant, strStatus, dteReceived
Dim dblTotal
Dim strStatusHold
Dim strReceivedFormat
Dim strLink
Dim strReceivedFull

'Open the master recordset
objRS.Open "SELECT * FROM tblApplication  Order by Status,DateReceived DESC",objCon

	strStatusHold = ""
	
	Do Until objRS.EOF
		'Insert Row Heading
		'Response.Write "<TR><TD >&nbsp;" & objRS("CardType") & "</TD>"	

		If strStatusHold = objRS("Status") Then
		
		Else
			Response.Write "<TR><TH colspan=""4"" style=""text-align:center; font-size:12px;font-weight:bold;"">" & objRS("Status") & "</TH></TD>"
		End If
			
		If objRS.EOF Then
			'Response.Write "<TD style=""text-align:right;"">0</TD><TD style=""text-align:right;"">0</TD><TD style=""text-align:right;"">0</TD>"
		Else
			
			If IsNull(objRS("CardTypeSub")) Then
				strCardType = ""
			Else
				strCardType = objRS("CardTypeSub")
			End If
			
			If strCardType = "Diners" Then
				strCardType = "<img src=""../Frest/app-assets/images/cards/Diners3.png"" height=""17px"" width=""30px""> " & strCardType
				'strCardType = "<img src=""img/Diners2.png"" height=""15px"" width=""30px""> " & strCardType
			ElseIf strCardType = "ANZ" Then
				strCardType = "<img src=""../Frest/app-assets/images/cards/visa.png"" height=""15px"" width=""30px""> " & strCardType
				'strCardType = "<img src=""img/ANZ2.png"" height=""15px"" width=""30px""> " & strCardType
			ElseIf strCardType = "Mastercard" Then
				strCardType = "<img src=""../Frest/app-assets/images/cards/mastercard.png"" height=""15px"" width=""30px""> " & strCardType
				'strCardType = "<img src=""img/ANZ2.png"" height=""15px"" width=""30px""> " & strCardType
			Else
				strCardType = "<img src=""../Frest/app-assets/images/cards/Diners5.png"" height=""15px"" width=""30px""> " & strCardType
				'strCardType = "<img src=""img/Mastercard2.png"" height=""15px"" width=""30px""> " & strCardType
				'"C:\Apps\CAPS\CAPS2\ASP2\Frest\app-assets\images\cards\mastercard.png"
			End If
			
			If IsNull(objRS("DateReceived")) Then
				dteReceived = ""
				strReceivedFull = ""
			Else
				dteReceived = objRS("DateReceived")
				strReceivedFull = "Title=""" & objRS("DateReceived") & """"
				If dteReceived < now() -10 then
					strReceivedFormat = "color:red; font-weight:bold;"
				Else
					strReceivedFormat = "color:black;"
				End If
				dteReceived = FormatDateTime(objRS("DateReceived"),2)
			End If
			
			If IsNull(objRS("FirstName")) Then
				strApplicant = ""
			Else
				strApplicant = objRS("FirstName") & " " & objRS("Surname")
			End If
			
			If isnull(objRS("Status")) Then
				strStatus = ""
			Else
				strStatus = objRS("Status")
			End If
			
			strLink = "<A target=""_parent"" href=""Applications3.asp?ApplicationID=" & objRS("ApplicationID") & "&EmployeeID=" & objRS("EmployeeID") & " "" Style=""" & strReceivedFormat & "font-size:12px;"">"
			
				
			Response.Write "<TR><TD style=""text-align:left; padding-left:20px; "">" & strLink & "" & strCardType & "</A></TD><TD style=""text-align:center; "">" & strLink & "" & strApplicant & "</A></TD>" & _
				"<TD style=""text-align:center; "">" & strLink & "" & strStatus & "</A></TD><TD style=""text-align:center; " & strReceivedFormat & """ " & strReceivedFull & ">" & strLink & "" & dteReceived & "</A></TD></a></TR>"
			
		
		End If
		
		strStatusHold = objRS("Status")
		
		objRS.Movenext
	Loop
	objRS.Close
	
	
%>
</table>
</body>

</html>
<%

Set objRS = Nothing
Set objCon = Nothing


%>
