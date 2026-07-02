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
 
	<tr>
		<th Style="Height:25px" Width="20%">Business Area</th>
		<th Width="40%">Name</th>
        <th Width="25%">Estimate</th>
		<th Width="15%">Status</th>
	</tr>
	
<%
Dim lngParentCCID
Dim dblGrandTotal

 objRS.Open "SELECT * FROM qryCostCentreApprovals WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND BusinessAreaID = " & Session("BusinessAreaID") & " Order By ParentCostCentreID,CostCentreID ASC",objCon
		
		Do until objRS.eof
		    dblEstimate = objRS("BMTotal")
            If IsNull(dblEstimate) Then dblEstimate = 0 End If
            If objRS("CostObjectTypeID") = 1 Then
        		Response.Write "<TR><TH style=""text-align:center""><B>" & objRS("ProgramCode") & "</A></B></TH><TH style=""text-align:left"">&nbsp;" & objRS("CostCentreName") & "</TH><TH style=""text-align:right"">&nbsp;" & formatnumber(dblEstimate,0,0) & "</TH><TH style=""text-align:center""><B>" & objRS("StatusName") & "</B>&nbsp;&nbsp;&nbsp;" & arrStatus(objRS("StatusID")) & "</TH></TR>"
			Else
                Response.Write "<TR><TD style=""text-align:center""><B>" & objRS("ProgramCode") & "</A></B></TD><TD style=""text-align:left"">&nbsp;" & objRS("CostCentreName") & "</TD><TD style=""text-align:right"">&nbsp;" & formatnumber(dblEstimate,0,0) & "</TD><TD style=""text-align:center""><B>" & objRS("StatusName") & "</B>&nbsp;&nbsp;&nbsp;" & arrStatus(objRS("StatusID")) & "</TD></TR>"
            End If
                dblEstimateTotal = dblEstimateTotal + dblEstimate
                dblGrandTotal = dblGrandTotal + dblEstimate
            lngParentCCID = objRS("ParentCostCentreID")
            objRS.movenext

            If Not objRS.EOF THen
                If objRS("ParentCostCentreID") <> lngParentCCID Then
                    Response.Write "<TR><TH Style=""Height:20px;"" Colspan=""2"">Total</TH><TD Style=""Text-Align:Right;"">" & formatnumber(dblEstimateTotal,0) & "</TD><TH></TH></TR>"
                    dblEstimateTotal = 0
                End If
            End If
		Loop
			
	objRS.Close

            If IsNull(dblEstimateTotal) Then dblEstimateTotal = 0 End If
	
            Response.Write "<TR><TH Colspan=""2"" Style=""text-align:Right; Height:20px"">&nbsp;Total&nbsp;</TH><TD Style=""text-align:Right""><B>" & formatnumber(dblGrandTotal,0,0) & "</B></TD><TH Colspan=""2""></TH>"
%>
</table>
</body>

</html>
<%

Set objRS = Nothing
Set objCon = Nothing


%>
