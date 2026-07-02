<%@ Language=VBScript %>
<% Option Explicit %>
<!-- #Include file=../ADOVBS.inc -->

<%

'Response.Write Session("UserID")

Response.Expires = -1500

If IsEmpty(Session("BudgetID")) Then Response.Redirect("../Timeout.asp")
 
'Description:	Student Screen
'Author:		Andrew Bull - Isidore Limited (www.isidore.com - 07969 589413)
'Date:			August 2004

'Declare default variables
Dim objCon
Dim objCmd
Dim objRS
Dim strScreen
Dim strSelected
Dim x 
Dim strMessage
Dim strMessageIcon

'Set Database objects

Set objCon = Server.CreateObject("ADODB.Connection")
Set objCmd = Server.CreateObject("ADODB.Command")
Set objRS = Server.CreateObject("ADODB.Recordset")

'Open database connection

objCon.Open Session("DBConnection")

'1. Declare screen specific variables naming convention = variable type prefix + database field name

Dim lngBudgetID
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
Dim strDefault

'Declare and set default arrays

Dim arrYesNo(2)
	
	arrYesNo(1) = "Y"
	arrYesNo(2) = "N"
	
	'3. Capture Querystring variables
	
	If Not IsEmpty(Request.QueryString("BudgetID")) Then
		
		Session("BudgetID") = Request.QueryString("BudgetID")
		lngBudgetID = clng(Request.QueryString("BudgetID"))	
				
	End If	
	'Execute save 	
	If Request.QueryString("Action") = "Save" Then
		SaveDetails()
	End If

	'Load page details
	LoadDetails()
		
%>

<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="../BERTStyle.css">
	<script src="../formChek.js">
	</script>
	<script src="../ButtonRollOver.js">
	</script>

<script LANGUAGE="javascript">
<!--
function SaveData()
{	    
    var varSubmit = true						
    var varAlert="";	

	if(isWhitespace(frm.BudgetID.value) || frm.BudgetID.value=="0")
	{
       varAlert += "Budget ID Cannot Be Blank. \n \n";
       document.getElementById('BudgetID').style.backgroundColor="ff8080";
       varSubmit = false;
    }   
    else document.getElementById('BudgetID').style.backgroundColor="ffffff";
    
    if((isNonnegativeInteger(frm.DefaultVersionID.value)==false) || (frm.DefaultVersionID.value == 0))
	   {            
		   varAlert += "Please enter Default Version. Default Version must be a numeric value.  \n \n";
		   document.getElementById('DefaultVersionID').style.backgroundColor="ff8080";
		   varSubmit = false;
	   }			
	else document.getElementById('DefaultVersionID').style.backgroundColor="ffffff";
	
   
    if(isWhitespace(frm.BudgetName.value))
	{
       varAlert += "Budget Name Cannot Be Blank. \n \n";
       document.getElementById('BudgetName').style.backgroundColor="ff8080";
       varSubmit = false;
    }   
    else document.getElementById('BudgetName').style.backgroundColor="ffffff";

    if(frm.FinancialYearID.value == 0 )
	    {
		    varAlert += "Please select a financial year. \n \n";
		    document.getElementById('FinancialYearID').style.backgroundColor="ff8080";
		    varSubmit = false;
	    }	
	    else document.getElementById('FinancialYearID').style.backgroundColor="ffffff";		
					
	if(varSubmit == true)
	{
	    frm.submit();		
	}
	else
	{
	    alert(varAlert);
	}
}

//-->
</script>
</head>
<body>
<h3>Budget Administration Screen</h3>
<form action="Budget.asp?Action=Save" method="POST" id="frm" name="frm">

<TABLE WIDTH=100% ALIGN=Center BORDER=1 CELLSPACING=1 CELLPADDING=1>
	<tr>
		<th height="20px" style="text-align:left">&nbsp;Budget ID</th>
		<td>&nbsp;<input style="text-align:left" style="width:90%" id="BudgetID" name="BudgetID" maxlength="50" TABINDEX="1" value="<%=lngBudgetID%>"></td>
		<th style="text-align:left">&nbsp;Budget Name</th>
		<td>&nbsp;<input style="text-align:left" style="width:90%" id="BudgetName" name="BudgetName" maxlength="50" TABINDEX="2" value="<%=strBudgetName%>"></td>
	</tr>
	<tr>
		<th height="20px" style="text-align:left">&nbsp;Default Version</th>
		<td>&nbsp;<input style="text-align:left" style="width:90%" id="DefaultVersionID" name="DefaultVersionID" maxlength="50" TABINDEX="3" value="<%=intDefaultVersionID%>"></td>
	 	<th style="text-align:left">&nbsp;Financial Year</th>
		<td><select Style="Width:40%" tabindex="4" id="FinancialYearID" name="FinancialYearID"><OPTION Value=0>Please Select..</OPTION>
	<%	
		objRS.Open "SELECT * FROM tblFinancialYears",objCon
		
		Do until objRS.EOF
			If objRS("FinancialYearID") = intFinancialYearID Then
				strSelected = " SELECTED "
			Else
				strSelected = ""
			End if
				Response.Write "<option Value=""" & objRS("FinancialYearID") & """" & strSelected & ">" & objRS("FinancialYearName") & "</OPTION>"
			objRS.Movenext
		Loop
		
		objRS.Close
		
	%></select></td>
	</TR>
	
	
	<th height="20px" style="text-align:left">&nbsp;System Default</th><td><select Style="Width:40%" tabindex="5" id="Active" name="Active">
	<%
		For x = 1 to 2
			If arrYesNo(x) = cstr(strDefault) Then
				strSelected = " SELECTED "
			Else
				strSelected = ""
			End If
				Response.Write "<option Value=""" & arrYesNo(x) & """" & strSelected & ">" & arrYesNo(x) & "</OPTION>"
		Next
	%>
		</select></td><th colspan="2"></th>
	</tr>
	<tr>
		<td colspan="4" Align="left">&nbsp;</td>
	</tr>
	<tr>
		<th height="20px" colspan="4" Align="left">&nbsp;Comments</th>
	</tr>
	<tr>
	    <td colspan="4"><TEXTAREA rows=4 cols=190 id=Comments name=Comments tabindex="6"><%=strComments%>
	
</TEXTAREA></td>
	</tr>
</table>
<br>

<table Class="CallCentre" WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
	<tr>
	    <td class='locked' Width="100px" style="border-right:0px"><button type="button" tabindex="8" onClick="self.location='AdminMenu.asp';"><img src="../images/door.png" alt="" /> Close </button></td>    
        <td class='locked' Width="100px" style="border-right:0px"><button type="button" tabindex="9" onclick="SaveData()";><img src="../images/tick.png" alt="" /> Save </button></td>  
	<td class='locked' Width="100px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><button type="button" onclick="self.location='Budget.asp?BudgetID=0'" )""><img src="../images/page_white_stack.png" alt="" /> Add New </button></td>        
	<TD class='locked' Width="100px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=strMessageIcon %></TD>
        <TD class='locked' Width="600px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=strMessage %></TD>
	</tr>
</table>
<hr>
</form>

<table WIDTH="100%" BORDER="1" CELLSPACING="1" CELLPADDING="1">
  
	<tr>
        <th height="20px">Budget Name</th>
		<th>Budget ID</th>
		<th>Financial Year</th>
		<th>Default Version</th>
		<th>System Default</th>
		<th>Updated By</th>
		<th>Date Updated</th>
	</tr>
      <tr>
        <td colspan="8">&nbsp;</td>
    </tr>
	
<%

 objRS.Open "SELECT * FROM qryBudgets Order By BudgetID ASC",objCon
		Do until objRS.eof
			Response.Write "<TR><TD>&nbsp;<A Target=""_self"" HREF=""Budget.asp?BudgetID=" & objRS("BudgetID") & """>" & objRS("BudgetName") & "</A></TD><TD>&nbsp;" & objRS("BudgetID") & "</B></TD><TD style=""text-align:center"">&nbsp;" & objRS("FinancialYearName") & "</TD><TD style=""text-align:center"">&nbsp;" & objRS("DefaultVersionID") & "</TD><TD style=""text-align:center"">&nbsp;" & objRS("SystemDefault") & "</TD><TD style=""text-align:center"">&nbsp;" & objRS("UpdatedBy") & "</TD><TD style=""text-align:center"">&nbsp;" & objRS("DateUpdated") & "</TD></TR>"
			objRS.movenext
		Loop
			
	objRS.Close

%>
</table>
</body>

</html>

<% 

Sub LoadDetails()

'Description:	Loads Caller's details into page if applicable.
		
		objRS.Open "SELECT * FROM tblBudget WHERE BudgetID = " & clng(Session("BudgetID")) & "",objCon
						
			If Not objRS.EOF Then
				
				lngBudgetID = objRS("BudgetID")
				strBudgetName = objRS("BudgetName")
				intFinancialYearID = objRS("FinancialYearID")
				intDefaultVersionID = objRS("DefaultVersionID")
				lngCashFlowGLCode = objRS("CashFlowGLCode")
				lngBadDebtGLCode = objRS("BadDebtGLCode")
				lngLoanGLCode = objRS("LoanGLCode")
				lngPrepaymentGLCode = objRS("PrepaymentGLCode")
				lngInvestmentGLCode = objRS("InvestmentGLCode")
				lngAPGLCode = objRS("APGLCode")
				lngARGLCode = objRS("ARGLCode")
				intVariancePercentage = objRS("VariancePercentage")
			    strComments = objRS("Comments")
				strActive = objRS("Active")	
				strDefault = objRS("SystemDefault")	
				
			End If

		objRS.Close
	

End Sub

Sub SaveDetails()

		If Not IsEmpty(Request.Form("BudgetName")) Then
		
		 With objCmd
                .CommandType = 4
                .CommandText = "spBudgetSave"
                
                .Parameters.Append objCmd.CreateParameter("BudgetID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("BudgetName", adVarChar, adParamInput, 50)
                .Parameters.Append objCmd.CreateParameter("FinancialYearID", adInteger)
                .Parameters.Append objCmd.CreateParameter("DefaultVersionID", adInteger)
                .Parameters.Append objCmd.CreateParameter("BalanceSheet", adChar, adParamInput, 1)
                .Parameters.Append objCmd.CreateParameter("CashFlowGLCode", adInteger)
                .Parameters.Append objCmd.CreateParameter("BadDebtGLCode", adInteger)
                .Parameters.Append objCmd.CreateParameter("PrepaymentGLCode", adInteger)
                .Parameters.Append objCmd.CreateParameter("LoanGLCode", adInteger)
                .Parameters.Append objCmd.CreateParameter("InvestmentGLCode", adInteger)
                .Parameters.Append objCmd.CreateParameter("APGLCode", adInteger)
                .Parameters.Append objCmd.CreateParameter("ARGLCode", adInteger)
                .Parameters.Append objCmd.CreateParameter("VariancePercentage", adInteger)                
                .Parameters.Append objCmd.CreateParameter("Comments", adLongVarChar, adParamInput, -1)
                .Parameters.Append objCmd.CreateParameter("SystemDefault", adChar, adParamInput, 1)
                .Parameters.Append objCmd.CreateParameter("ClientID", adInteger)    
                .Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger)                
             
              
				.Parameters("BudgetID") = Request.Form("BudgetID")	
				.Parameters("BudgetName") = Request.Form("BudgetName")			
                .Parameters("FinancialYearID") = Request.Form("FinancialYearID")
                .Parameters("DefaultVersionID") = Request.Form("DefaultVersionID")    
                .Parameters("BalanceSheet") = "N"
                .Parameters("CashFlowGLCode") = 0
                .Parameters("BadDebtGLCode") = 0  
                .Parameters("PrepaymentGLCode") = 0
                .Parameters("LoanGLCode") = 0   
                .Parameters("InvestmentGLCode") = 0   
                .Parameters("APGLCode") = 0   
                .Parameters("ARGLCode") = 0 
                .Parameters("VariancePercentage") = 130
                .Parameters("Comments") = Request.Form("Comments")
                .Parameters("SystemDefault") = Request.Form("Active")
                .Parameters("ClientID") = 1
                .Parameters("UpdatedBy") = Session("UserID")
                           
                .ActiveConnection = objCon
                
            End With
                
            objCmd.Execute          
               
			'Return the result of the Save Function.
     		Session("BudgetID") = Request.Form("BudgetID")
     		strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"
            strMessage = "<B>RECORD SAVED.</B>"
			
						
	End If
				
	

End Sub	

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
