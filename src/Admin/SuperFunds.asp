<%@ Language=VBScript %>
<% Option Explicit %>
<!-- #Include file=../ADOVBS.inc -->

<%
 
'Description:	Super Funds Admin Screen
'Author:		Michael Giacomin
'Date:			January 2010

'Declare default variables

Dim objCon
Dim objCmd
Dim objRS
Dim strScreen
Dim strSelected
Dim x 
Dim strMessage
Dim strMessageIcon
Dim strMessageColour

'Set Database objects

Set objCon = Server.CreateObject("ADODB.Connection")
Set objCmd = Server.CreateObject("ADODB.Command")
Set objRS = Server.CreateObject("ADODB.Recordset")

'Open database connection

objCon.Open Session("DBConnection")

'1. Declare screen specific variables naming convention = variable type prefix + database field name

Dim intSuperFundID
Dim strSuperFundName
Dim strSuperFundDesc
Dim dblRate
Dim strActive
Dim lngGLCode

Dim arrYesNo(2)

	arrYesNo(1) = "Y"
	arrYesNo(2) = "N"
	
'3. Capture Querystring variables
If Session("SuperFundID") = ""  Then Session("SuperFundID") = 0

    If Not IsEmpty(Request.QueryString("SuperFundID")) Then
	   Session("SuperFundID") = Request.QueryString("SuperFundID")
    End If
		
	'Execute save 	
	If Request.QueryString("Action") = "Save" Then
		SaveDetails()
	End If

	'Execute Recalculate procedure 	
	If Request.QueryString("Action") = "recalculate" Then
		RecalcStaff()
	End If
	
	'Execute Delete 	
	If Request.QueryString("Action") = "Delete" Then
		DeleteRecord(Request.QueryString("SuperFundName"))
	End If
	
	'Load page details
	LoadDetails()
		
%>

<html>
<head>
<title></title>
<meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="../BERTStyle.css">
<script type="text/javascript" src="../formChek.js"></script>
<script type="text/javascript" language="javascript">
<!--
   function SaveData()
    {
        var varSubmit = true
        var varAlert =""  
        
        
        if((isFloat(frm.Rate.value)==false))
	    {            
		    varAlert += "Please enter a Rate. The Rate must be a decimal value.  \n \n";
		    document.getElementById('Rate').style.backgroundColor="ff8080";
		    varSubmit = false;
	    }			
	    else document.getElementById('Rate').style.backgroundColor="ffffff";	
	    
        
          	    
		    
	    if(isWhitespace(frm.SuperFundName.value))
        {            
		    varAlert += "Please enter the Super Fund Name. \n \n";
		    document.getElementById('SuperFundName').style.backgroundColor="ff8080";
		    varSubmit = false;
	    }			
	    else document.getElementById('SuperFundName').style.backgroundColor="ffffff";    
	    
	        
	    if(frm.Active.value == 0 )
	    {
		    varAlert += "Please select a Active Status. \n \n";
		    document.getElementById('Active').style.backgroundColor="ff8080";
		    varSubmit = false;
	    }	
	    else document.getElementById('Active').style.backgroundColor="ffffff";
	    	   		  
	   	
	  if(varSubmit == true)
	  {
	        frm.submit();
	  }
	  else
	  {
	    window.alert ("" + varAlert);	    
	  }
  }
  
 
function StaffRecalc(varBud, varVer)
 {
 	if(window.confirm('Would you like to Recalculate ALL existing Staff values? \n \n For the currently selected Budget ' + varBud + ' and Version ' + varVer + '.')==true){
 	
 		document.getElementById('Progress').style.display = "inline";
        self.location='SuperFunds.asp?Action=recalculate';
	}
}
function DeleteData(){
	if( confirm("Delete the selected record?") )
	{
		self.location="SuperFunds.asp?Action=Delete&SuperFundName="+frm.SuperFundName.value;
		frm.elements['msgbox'].value = 'Deleting...';}
	}
//-->
</script>
</head>
<body>
<form action="SuperFunds.asp?Action=Save" method="POST" id="frm" name="frm">

<table width="100%" align="Center" border="1" cellspacing="1" cellpadding="1">
	<tr><th style="width:20%; height:20px">Superannuation Funds</th><th style="width:30%">

	</th><th style="width:20%"></th><th style="width:30%"></th></tr>
	<tr><td colspan="4">&nbsp;</td></tr>
	<tr>
		<th align="Left">&nbsp;Super Fund Name</th>
	    <td>&nbsp;<input style="text-align:left;width:90%" id="SuperFundName" name="SuperFundName" maxlength="50" tabindex="1" value="<%=strSuperFundName%>"></td>
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<th align="Left">&nbsp;Super Fund Description</th>
	    <td>&nbsp;<input style="text-align:left;width:90%" id="SuperFundDesc" name="SuperFundDesc" tabindex="2" value="<%=strSuperFundDesc%>"></td>
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<th align="left">&nbsp;Rate</th>
	    <td>&nbsp;<input style="text-align:left;width:90%" id="Rate" name="Rate" maxlength="10" tabindex="3" value="<%=dblRate%>"></td>
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<th align="left">&nbsp;GL Code</th>
		<td><select style="Width:95%;" tabindex="4" id="GLCode" name="GLCode"><option Value="0">Please Select....</option>
		<%
		objRS.Open "SELECT * FROM tblGLCodes WITH(NOLOCK) WHERE BudgetID = " & Session("BudgetID") & " AND GLCodeType = 'E' AND Active = 'Y'",objCon,0,1
		
		Do until objRS.EOF
			If objRS("GLCode") = lngGLCode Then
				strSelected = " SELECTED "
			Else
				strSelected = ""
			End if
				Response.Write "<option Value=""" & objRS("GLCode") & """" & strSelected & ">" & objRS("GLCode") & " : " & objRS("GLCodeName") & "</OPTION>"
			objRS.Movenext
		Loop
		
		objRS.Close
		%>
		</select> </td>
		<td colspan="2">&nbsp;</td>
	</TR>
	<tr>
		<th align="left">&nbsp;Active</th>		
		<td><select Style="Width:40%" tabindex="5" id="Active" name="Active"><option Value="0">Please Select....</option>
		<%
		For x = 1 to 2
			If strActive = arrYesNo(x)Then
				strSelected = " SELECTED " 
			Else
				strSelected = ""	
			End If
			
			Response.Write "<option Value=""" & arrYesNo(x) & """" & strSelected & ">" & arrYesNo(x) & "</OPTION>"
		Next
		%>
		</select> </td>
		<td colspan="2">&nbsp;</td>
	</tr>
	
	<tr>
		<td colspan="4" align="left">&nbsp;</td>
	</tr>
	
</table>
<br/>
<div class="buttons">
<TABLE Width="1500px" BORDER="0" CELLSPACING="0" CELLPADDING="0">
<TR>

    <td Width="100px"><button type="button" onclick="self.location='AdminMenu.asp';"><img src="../images/door.png" alt="" /> Close </button></td>    
    <td Width="100px"><button type="button" onclick="SaveData()";><img src="../images/tick.png" alt="" /> Save </button></td> 
    <td Width="150px"><button type="button" onclick="self.location='SuperFunds.asp?SuperFundID=0'" )""><img src="../images/add.png" alt="" /> Clear/Add New </button></td>
    <td Width="100px"><button type="button" onclick="DeleteData()";><img src="../images/delete.png" alt="" /> Delete </button></td>
    <TD class='locked' Width="100px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=strMessageIcon %></TD>
    <TD class='locked' Width="300px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=strMessage %></TD>
	
</TR>
</TABLE>
</div>

<hr>
</form>

<table WIDTH="100%" BORDER="1" CELLSPACING="1" CELLPADDING="1">
<tr><td colspan="8">&nbsp;<span style="color:gray; font-weight:bold;">Note:</span><span style="color:gray;"> Changes to Super Fund Rates does not update existing Staff records. The Recalculate button (above) must be clicked to recalculate existing records.</span></td></tr>
	<tr>
		<th>Super Fund Name</th>
		<th>Super Fund Description</th>
	    <th>Rate</th>
		<th>Active</th>
		<th>Updated By</th>
		<th>Date Updated</th>
	</tr>
<%
    objRS.Open "SELECT * FROM  qrySuperFund WHERE BudgetID = " & clng (Session("BudgetID")) & " AND VersionID = " & Session("VersionID") & "",objCon

	Do until objRS.eof			
   	   Response.Write "<TR><TD><A Target=""_self"" HREF=""SuperFunds.asp?SuperFundID=" & objRS("SuperFundID") & """>&nbsp;" & objRS("SuperFundName") & "</TD><TD style=""text-align:left"">&nbsp;" & objRS("SuperFundDesc") & "</TD><TD style=""text-align:right"">&nbsp;" & formatpercent(objRS("Rate")/100,2) & "</TD><TD style=""text-align:center"">&nbsp;" & objRS("Active") & "</TD><TD style=""text-align:center"">&nbsp;" & objRS("UpdatedByName") & "</TD><TD style=""text-align:center"">&nbsp;" & objRS("DateUpdated") & "</TD></TR>"
       objRS.movenext
	Loop
		
	objRS.Close
	
%>

</table>
</body>

</html>

<% 

Sub LoadDetails()

       'Description:	Loads Super Fund details into page if applicable.
		objRS.Open "SELECT * FROM tblSuperFund WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND SuperFundID = " & Session("SuperFundID") & "",objCon

			If Not objRS.EOF Then
				strSuperFundName = objRS("SuperFundName")
				strSuperFundDesc = objRS("SuperFundDesc")
				dblRate = objRS("Rate")
				strActive = objRS("Active")
				lngGLCode = objRS("GLCode")
    		Else
			  
				strSuperFundName = ""
				strSuperFundDesc = ""
				dblRate = 0
				strActive = "Y"
				lngGLCode = 0
				
           End If

		objRS.Close
	
End Sub

Sub SaveDetails()

		 With objCmd
                .CommandType = 4
                .CommandText = "spSuperFundSave"
                
				.Parameters.Append objCmd.CreateParameter("BudgetID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("VersionID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("SuperFundID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("SuperFundName", adVarChar, adParamInput, 50)
                .Parameters.Append objCmd.CreateParameter("SuperFundDesc", adLongVarChar, adParamInput,-1)
                .Parameters.Append objCmd.CreateParameter("Rate", adDouble, adParamInput)
				.Parameters.Append objCmd.CreateParameter("GLCode", adInteger)
                .Parameters.Append objCmd.CreateParameter("Active", adChar, adParamInput, 1)
                .Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger)
          	
				.Parameters("BudgetID") = clng(Session("BudgetID"))
   				.Parameters("VersionID") = clng(Session("VersionID"))
   				.Parameters("SuperFundID") = clng(Session("SuperFundID"))
				.Parameters("SuperFundName") = Request.Form("SuperFundName")	
				.Parameters("SuperFundDesc") = Request.Form("SuperFundDesc")	
				.Parameters("Rate") = Request.Form("Rate")
				.Parameters("GLCode") = Request.Form("GLCode")  				
                .Parameters("Active") = Request.Form("Active")
                .Parameters("UpdatedBy") = Session("UserID")
                           
                .ActiveConnection = objCon
                
            End With
                
            objCmd.Execute          
               
	   'Return the result of the Save Function.
				strMessage = "RECORD SAVED."
                strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"
                strMessageColour = "Black"
	
End Sub	


Sub RecalcStaff()


		 With objCmd
                .CommandType = 4
                .CommandText = "spStaffDataSalaryUpdate"
                
				.Parameters.Append objCmd.CreateParameter("BudgetID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("VersionID", adInteger, adParamInput)
          	
				.Parameters("BudgetID") = clng(Session("BudgetID"))
   				.Parameters("VersionID") = clng(Session("VersionID"))
				
                .ActiveConnection = objCon
                
            End With
                
            objCmd.Execute          
               
			'Return the result of the Save Function.
     	    strMessage = "Staff Values have been recalculated!"
	
End Sub	

Public Sub DeleteRecord(strSuperFundName)
'Procedure to delete Salary Classification records if no values exist against it.

Dim intAllow
Dim intCostCentreID

	'First Check to see if any values exist for the GL Code
    objRS.Open "SELECT [SuperFundID],[Deleted], [CostCentreID] FROM tblStaffingClassifications With(NoLock) WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND SuperFundID = " & Session("SuperFundID") & " GROUP BY [SuperFundID],[CostCentreID],[Deleted]" ,objCon,adOpenStatic,adLockReadOnly
                	
        If objRS.EOF Then
        
            intAllow = 1
        Else
            intCostCentreID = objRS("CostCentreID")
            If objRS("Deleted") = "Y" Then
                intAllow = 1
            End If
        End If
                	
    objRS.Close
    
    
    'If there are no values for the account code then it can be deleted
	If intAllow = 1 Then
	
        objCon.Execute("DELETE FROM tblSuperFund WHERE [BudgetID] = " & Session("BudgetID") & " AND SuperFundID = " & Session("SuperFundID") & " AND VersionID = " & Session("VersionID") & "")
        'objCon.Execute("DELETE FROM tblGLCodes WHERE [BudgetID] = " & Session("BudgetID") & " AND GLCode = " & intGLCode & "")
        strMessage = "DELETED Super Fund = " & strSuperFundName 
        'strMessage = "DELETE FROM tblSalaryClassifications WHERE [BudgetID] = " & Session("BudgetID") & " AND SalaryClassification = '" & strSalaryClassification & "' AND VersionID = " & Session("VersionID") & ""
        'Response.Write "DELETE FROM tblSalaryClassifications WHERE [BudgetID] = " & Session("BudgetID") & " AND SalaryClassification = '" & strSalaryClassification & "' AND VersionID = " & Session("VersionID") & ""
	Else
        strMessage = "NOT DELETED!. Super Fund: " & strSuperFundName & " is assigned to an employee in this Budget, so CANNOT BE DELETED unless removed from Employee(s) in the Salary Rates screen (access from Salary screen)...(CCID=" & intCostCentreID & ")"
    End If

End Sub

Set objRS = Nothing
Set objCon = Nothing


%>
