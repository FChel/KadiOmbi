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
 
	<tr><th colspan="4"><%=Session("Logon")%> Details</td></tr>
	<tr>
		<th Style="Height:25px" Width="20%">Benefit</th>
		<th Width="40%">Gross Taxable Value</th>
        <th Width="25%">Contributions</th>
		<th Width="15%">FBT Payable</th>
	</tr>
	
<%

Dim y

Dim intCars, intDays, dblEmpCont
Dim dblTotal

'Open the master recordset
objRS.Open "SELECT * FROM qryCarParkingStaffing WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND BusinessAreaID = " & Session("BusinessAreaID") & " AND StaffingClassificationID = " & Session("StaffingClassificationID") & " Order by CostCentreID",objCon

	Do Until objRS.EOF
		'Insert Row Heading
		Response.Write "<TR><TD >&nbsp;" & objRS("BenefitType") & " GTV</TD>"	

			
		If objRS.EOF Then
			Response.Write "<TD style=""text-align:right;"">0</TD><TD style=""text-align:right;"">0</TD><TD style=""text-align:right;"">0</TD>"
		Else
			
			If IsNull(objRS("Cars")) Then
				intCars = 0
			Else
				intCars = objRS("Cars")
			End If
			
			If IsNull(objRS("Days")) Then
				intDays = 0
			Else
				intDays = objRS("Days")
			End If
			
			If IsNull(objRS("EmployeeContribution")) Then
				dblEmpCont = 0
			Else
				dblEmpCont = objRS("EmployeeContribution")
			End If
			
			If objRS("BenefitType") = "Car Parking" Then
				dblTotal = intCars*intDays*dblEmpCont
			Else
				dblTotal = intCars
			End If
				Response.Write "<TD style=""text-align:right;"">" & FormatCurrency(dblTotal,2) & "</TD><TD style=""text-align:right;"">0</TD><TD style=""text-align:right;"">0</TD>"
			
		
		End If
		
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
