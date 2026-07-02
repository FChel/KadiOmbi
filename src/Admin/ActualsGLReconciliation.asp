<%@ Language=VBScript %>
<% Option Explicit %>
<!-- #Include file=../ADOVBS.inc -->

<%

Response.Expires = -1500

If IsEmpty(Session("BudgetID")) Then Response.Redirect("../Timeout.asp")
If IsEmpty(Session("FiscalYear")) Then Session("FiscalYear") = Session("FinancialYear")

Session("CurrentPage") = "Admin/ActualsGLReconciliation.asp"
 
'Description:	Business Area Screen
'Author:		Andrew Bull - Isidore Limited (www.isidore.com - 07969 589413)
'Date:			November 2007

'Declare default variables

Dim objCon
Dim objCmd
Dim objRS
Dim strScreen
Dim strSelected
Dim x 
Dim strMessage
Dim strMessageIcon
Dim dblTotal
Dim dblTotalRev
Dim dblTotalExp
Dim strMonth
Dim arrMonth(2)
Dim arrMonthName(2)
Dim strPeriod

arrMonth(1) = "1"
arrMonth(2) = "2"
arrMonthName(1) = "MAIN"
arrMonthName(2) = "BUDGET"


'Set Database objects

Set objCon = Server.CreateObject("ADODB.Connection")
Set objCmd = Server.CreateObject("ADODB.Command")
Set objRS = Server.CreateObject("ADODB.Recordset")

'Open database connection

objCon.Open Session("DBConnection")

If Not IsEmpty(Request.QueryString("Month")) Then		
	Session("Month") = Request.QueryString("Month")
    strMonth = Session("Month")
Else
    strMonth = Session("Month")  
End If	

'If Request.QueryString("Month") = "" Then
   ' strMonth = "MAIN"
   ' Session("Month") = strMonth
'End IF

If Not IsEmpty(Request.QueryString("BusinessAreaID")) Then	   
    Session("BusinessAreaID") = Request.QueryString("BusinessAreaID")
End IF

If Not IsEmpty(Request.QueryString("FiscalYear")) Then	   
    Session("FiscalYear") = Request.QueryString("FiscalYear")
End IF

If Request.QueryString("Action") = "ACCY" Then
    Load_Actuals "MAIN","ACCY"
End If

If Request.QueryString("Action") = "ACAM" Then
    Load_Actuals "MAIN","ACAM"
End If

If Request.QueryString("Action") = "BBUD" Then
    Load_Actuals "BUDGET","BBUD"
End If

Dim strCurrentYear
Dim strPreviousYear

objRS.Open "SELECT ERPFiscalYear FROM tblFinancialYears WHERE FinancialYearID = " & Session("FinancialYearID") & "",objCon
    If Not objRS.EOF Then
        strCurrentYear = objRS("ERPFiscalYear")
    End If

objRS.Close

objRS.Open "SELECT ERPFiscalYear FROM tblFinancialYears WHERE FinancialYearID = " & Session("FinancialYearID") - 1 & "",objCon
    If Not objRS.EOF Then
        strPreviousYear = objRS("ERPFiscalYear")
    End If

objRS.Close

		
%>

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title></title>
<meta name="GENERATOR" content="Microsoft Visual Studio 6.0"/>
<link rel="stylesheet" type="text/css" href="../BERTStyle.css"/>
<script type="text/javascript" src="../formChek.js"></script>
<script type="text/javascript" src="../ButtonRollOver.js"></script>
<script type="text/javascript" language="javascript">

    function MonthChange() {
        alert('sdsd');
        self.location = "ActualsGLReconciliation.asp?Month=" + document.frm.Months.value 
    }

    function LoadActuals(Month) {
        
        document.getElementById('Progress').style.display = "inline";
        self.location = 'ActualsGLReconciliation.asp?Action=' + Month;

    }
//-->
</script>
</head>
<body>
<h3>Import Actuals Administration Screen
    <%
	Response.write "for the currently selected Budget : <FONT color=Red>" &  Session("BudgetName") & "</FONT></H3>"
%>
<form action="ActualsGLReconcilitation.asp" method="POST" id="frm" name="frm">
<br />
<table width="100%" align="Left" border="1" cellspacing="1" cellpadding="1">

	<tr>
	<th style="height:25px; width:15%; text-align:left">&nbsp;Ledger Book</th>
		<td width="15%">
		    <select Style="Width:50%" tabindex="1" id="Months" name="Months" onchange="self.location='ActualsGLReconciliation.asp?Month=' + frm.Months.value"><OPTION Value=0>Please Select..</OPTION>
		    <%
		        For x = 1 to 2
			        If arrMonthName(x) = Session("Month") Then
				        strSelected = " SELECTED "
			        Else
				        strSelected = ""
			        End If
				        Response.Write "<option Value=""" & arrMonthName(x) & """" & strSelected & ">" & arrMonthName(x) & "</OPTION>"
		        Next
	        %>
	      </select>
	    </td>
        <th style="height:25px; width:15%; text-align:left">&nbsp;Fiscal Year</th>
		<td width="15%">
		    <select Style="Width:50%" tabindex="1" id="FiscalYear" name="FiscalYear" onchange="self.location='ActualsGLReconciliation.asp?FiscalYear=' + frm.FiscalYear.value"><OPTION Value=0>Please Select..</OPTION>
		    <%
		        
			objRS.Open "SELECT ERPFiscalYear FROM tblFinancialYears",objCon
    		
		        Do until objRS.EOF
			        If objRS("ERPFiscalYear") = cint(Session("FiscalYear")) Then
				        strSelected = " SELECTED "
			        Else
				        strSelected = ""
			        End if
				        Response.Write "<option Value=""" & objRS("ERPFiscalYear") & """" & strSelected & ">" & objRS("ERPFiscalYear") & "</OPTION>"
			        objRS.Movenext
		        Loop
    		
		    objRS.Close
	        %>
	      </select>
	    </td>
	    <th style="height:25px; width:20%; text-align:left">&nbsp;Business Area</th>
		<td width="20%">
		    <select Style="Width:90%" tabindex="1" id="Vote" name="Vote" onchange="self.location='ActualsGLReconciliation.asp?BusinessAreaID=' + frm.Vote.value">
		    <%
		        
			objRS.Open "SELECT * FROM tblBusinessArea WHERE BudgetID = " & Session("BudgetID") & " AND Active = 'Y'",objCon
    		
		        Do until objRS.EOF
			        If objRS("BusinessAreaID") = cint(Session("BusinessAreaID")) Then
				        strSelected = " SELECTED "
			        Else
				        strSelected = ""
			        End if
				        Response.Write "<option Value=""" & objRS("BusinessAreaCode") & """" & strSelected & ">" & objRS("BusinessAreaCode") & " - " & objRS("BusinessAreaName") & "</OPTION>"
			        objRS.Movenext
		        Loop
    		
		    objRS.Close
	        %>
	      </select>
	    </td>
	
	
	</tr>

</table>
<br />
<br />
<br />

<table WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
<tr>
<td class='locked' Width="100px" style="border-right:0px"><button type="button" tabindex="8" onClick="self.location='AdminMenu.asp';"><img src="../images/door.png" alt="" /> Close </button></td>    
<td Width="250px" style=text-align:center><button type="button" name=tabindex="15" onclick="javascript:LoadActuals('ACCY');"><img src="../images/database_save.png" alt="" /> Import Prior Year Actuals</button></td>
<td Width="250px" style=text-align:center><button type="button" name=tabindex="15" onclick="javascript:LoadActuals('ACAM');"><img src="../images/database_save.png" alt="" /> Import Current Year Actuals</button></td>
<td Width="250px" style=text-align:center><button type="button" name=tabindex="15" onclick="javascript:LoadActuals('BBUD');"><img src="../images/database_save.png" alt="" /> Import Base Budget</button></td>
<TD Width="100px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=strMessageIcon %></TD>
<TD Width="300px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=strMessage %></TD>
    </tr>
  </table>
  <hr />
  <bR />
<span id="Progress" style="display:none"><img src=../Images/progress.gif />  &nbsp;&nbsp;&nbsp; <b><FONT Color="Red">LOADING ACTUALS...</FONT></b></span>


<table width="100%" align="Center" border="1" cellspacing="1" cellpadding="1">
	<tr><th  style="height:25px" style="text-align:left" colspan="5">&nbsp;General Ledger Actual Reconciliation</th></tr>
	<tr><td  style="height:25px" colspan="5">&nbsp;</td></tr>
	<tr>
		<th style="height:25px;width:50%">GL Code</th>
        <th style="height:25px;width:10%">GL Code Type</th>
		<th style="height:25px;width:10%">Book ID</th>
        <th style="height:25px;width:10%">Fiscal Year</th>
		<th style="height:25px;width:20%">Amount</th>
	</tr>
<%

    If Session("BusinessAreaID") = 1000 Then
        objRS.Open "SELECT GLCode,BookID,Sum(TransAmt) AS Amount,GLCodeType,GLCodeName,FiscalYear FROM qryERPRawData WHERE BudgetID = " & clng (Session("BudgetID")) & " AND VersionID = " & Session("VersionID") & " AND BookID = '" & strMonth & "' AND FiscalYear = " & Session("FiscalYear") & " Group By GLCode,BookID,GLCodeType,GLCodeName, FiscalYear Order By GLCode",objCon
        'Response.Write "SELECT GLCode,BookID,Sum(TransAmt) AS Amount,GLCodeType,GLCodeName,FiscalYear FROM qryERPRawData WHERE BudgetID = " & clng (Session("BudgetID")) & " AND VersionID = " & Session("VersionID") & " AND BookID = '" & strMonth & "' AND FiscalYear = " & Session("FiscalYear") & " Group By GLCode,BookID,GLCodeType,GLCodeName, FiscalYear Order By GLCode"
    Else
        objRS.Open "SELECT GLCode,BookID,Sum(TransAmt) AS Amount,GLCodeType,GLCodeName,FiscalYear FROM qryERPRawData WHERE BudgetID = " & clng (Session("BudgetID")) & " AND VersionID = " & Session("VersionID") & " AND BookID = '" & strMonth & "' AND FiscalYear = " & Session("FiscalYear") & " AND Left(CostCentreID,4) = '" & Session("BusinessAreaID") & "' Group By GLCode,BookID,GLCodeType,GLCodeName, FiscalYear Order By GLCode",objCon
       'Response.Write "SELECT GLCode,BookID,Sum(TransAmt) AS Amount,GLCodeType,GLCodeName,FiscalYear FROM qryERPRawData WHERE BudgetID = " & clng (Session("BudgetID")) & " AND VersionID = " & Session("VersionID") & " AND BookID = '" & strMonth & "' AND FiscalYear = " & Session("FiscalYear") & " AND Left(CostCentreID,4) = '" & Session("BusinessAreaID") & "' Group By GLCode,BookID,GLCodeType,GLCodeName, FiscalYear Order By GLCode"
    End If

	Do Until objRS.EOF
	   If objRS("GLCodeType") = "R" Then
	        dblTotal = dblTotal + objRS("Amount")
	        dblTotalRev = dblTotalRev + objRS("Amount")
	   Else
	        dblTotal = dblTotal + objRS("Amount")
	        dblTotalExp = dblTotalExp + objRS("Amount")
	   End If
	   
            'If objRS("Month") = "CY" Then
               ' strPeriod = "Previous Year"
           ' Else
               ' strPeriod = "Current Year"
           ' End If


   	   Response.Write "<TR><TD style=""text-align:left"">&nbsp;" & objRS("GLCode") & " : " & objRS("GLCodeName") & "</TD><TD style=""text-align:center"">" & objRS("GLCodeType") & "</TD><TD style=""text-align:center"">" & objRS("BookID") & "</TD><TD style=""text-align:center"">" & objRS("FiscalYear") & "</TD><TD style=""text-align:right"">" & formatnumber(objRS("Amount"),0,0) & "</TD></TR>"
       objRS.movenext
	Loop
		
	objRS.Close
	dblTotal = dblTotal * -1
	Response.Write "<TR><TH Colspan=""3""></TH><TD style=""text-align:center""><B>Revenue Total</B></TD><TD style=""text-align:right""><B>" & formatnumber(dblTotalRev,0,0) & "</B></TD></TR>"
	Response.Write "<TR><TH Colspan=""3""></TH><TD style=""text-align:center""><B>Expense Total</B></TD><TD style=""text-align:right""><B>" & formatnumber(dblTotalExp,0,0) & "</B></TD></TR>"
	Response.Write "<TR><TH Colspan=""3""></TH><TD style=""text-align:center""><B>Result Total</B></TD><TD style=""text-align:right""><B>" & formatnumber(dblTotal,0,0) & "</B></TD></TR>"
%>

</table>
</form> 
</body>
<% 
Public Sub Load_Actuals(Book,TransactionType)

     objCon.Execute "spLoadEpicorLedger " & Session("BudgetID") & "," & Session("VersionID") & ",'" & Book & "','" & TransactionType & "'," & Session("UserID") & ""
    'Response.Write "<font face=arial size=2 color=red><b>" & Month & " Actuals have been loaded.</font><br>"
    'Response.Write "spLoadBudgetDataGLActuals " & Session("BudgetID") & "," & Session("VersionID") & ",'" & Month & "'," & Session("UserID") & ""
    strMessage = "<B>ACTUALS HAVE BEEN SUCCESSFULLY LOADED.</B>"
    strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"	

End Sub
%>
</html>

