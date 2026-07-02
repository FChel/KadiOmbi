<%@ Language=VBScript %>
<% Option Explicit %>
<!-- #Include file=../ADOVBS.inc -->

<%

    Response.Expires = -1500
	
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment; filename=CostCentreStatusExcel.xls" 
 
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
	
	'Execute save 	
	If Request.QueryString("Action") = "Save" Then
		SaveDetails()
		
	End If
    
Dim arrStatus(4)

arrStatus(0) = "<IMG SRC='../images/delete.png'"
arrStatus(1) = "<IMG SRC='../images/open.png'"
arrStatus(2) = "<IMG SRC='../images/ready.gif'"
arrStatus(3) = "<IMG SRC='../images/reject.gif'" 
arrStatus(4) = "<IMG SRC='../images/Closed.png'"	
		
%>

<html>
<head>

<body>

<table WIDTH="100%" BORDER="1" CELLSPACING="1" CELLPADDING="1">
    <tr>
        <th Style="width:800px; background-color:silver;" colspan="4">Workflow Status&nbsp;:&nbsp;<%=Session("Vote")%></th>
    </tr>
    <tr>
        <td colspan="4">&nbsp;</td>
    </tr>
	<tr>
		<th Style="width:200px; background-color:silver;">Responsiblity</th>
		<th Style="width:600px; background-color:silver;">Responsiblity Name</th>
		<th Style="width:200px; background-color:silver;">Type</th>
		<th Style="width:200px; background-color:silver;">Status</th>
	</tr>
	
<%

 objRS.Open "SELECT * FROM qryCostCentreApprovals WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND BusinessAreaID = " & Session("BusinessAreaID") & " Order By ParentCostCentreID,CostCentreID ASC",objCon
		
		Do until objRS.eof	   
		     
			Response.Write "<TR><TD style=""text-align:center""><B>" & objRS("ProgramCode") & "</A></B></TD><TD style=""text-align:left"">&nbsp;" & objRS("CostCentreName") & "</TD><TD style=""text-align:left"">&nbsp;" & objRS("CostObjectTypeName") & "</TD><TD style=""text-align:center""><B>" & objRS("StatusName") & "</B></TD></TR>"
			
            objRS.movenext
		Loop
			
	objRS.Close
	

%>
</table>
</body>

</html>

<%



Function MediumDate (str)
	
	'Function to change all date formats to medium date to avoid American storage challenge!
	
	Dim aDay
	Dim aMonth
	Dim aYear
	
		aDay = 	(Left((str),InStr(1,(str),"/")-1))
		aMonth = Mid(str,(InStr(1,(str),"/")+1),2)
	
	If Right(aMonth,1) = "/" Then
		aMonth = Left(aMonth,1)
	End If
	
		aMonth = MonthName(aMonth)
		aYear = Year(str)
	
	If Len(aDay) = 1 Then aDay = "0" & aDay
	
		MediumDate = aDay & "-" & aMonth & "-" & aYear
		
End Function

Set objRS = Nothing
Set objCon = Nothing


%>
