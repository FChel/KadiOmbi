<%@ Language=VBScript %>
<% Option Explicit %>
<!-- #Include file=../ADOVBS.inc -->

<%
 
'Description:	Positions Admin Screen
'Author:		Michael Giacomin
'Date:			April 2016

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

Dim intPositionID
Dim strPositionName
Dim strPositionDesc
Dim dblRate
Dim strActive

Dim arrYesNo(2)

	arrYesNo(1) = "Y"
	arrYesNo(2) = "N"
	
'3. Capture Querystring variables
If Session("PositionID") = ""  Then Session("PositionID") = 0

    If Not IsEmpty(Request.QueryString("PositionID")) Then
	   Session("PositionID") = Request.QueryString("PositionID")
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
		DeleteRecord(Request.QueryString("PositionName"))
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
        	    
	    if(isWhitespace(frm.PositionName.value))
        {            
		    varAlert += "Please enter the Position Name. \n \n";
		    document.getElementById('PositionName').style.backgroundColor="ff8080";
		    varSubmit = false;
	    }			
	    else document.getElementById('PositionName').style.backgroundColor="ffffff";    
	    
	        
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
        self.location='Positions.asp?Action=recalculate';
	}
}
function DeleteData(){
	if( confirm("Delete the selected record?") )
	{
		self.location="Positions.asp?Action=Delete&PositionName="+frm.PositionName.value;
		frm.elements['msgbox'].value = 'Deleting...';}
	}
//-->
</script>
</head>
<body>
<form action="Positions.asp?Action=Save" method="POST" id="frm" name="frm">

<table width="100%" align="Center" border="1" cellspacing="1" cellpadding="1">
	<tr><th style="width:20%">Superannuation Funds</th><th style="width:30%">

	</th><th style="width:20%"></th><th style="width:30%"></th></tr>
	<tr><td colspan="4">&nbsp;</td></tr>
	<tr>
		<th align="Left">&nbsp;Position Name</th>
	    <td>&nbsp;<input style="text-align:left;width:90%" id="PositionName" name="PositionName" maxlength="50" tabindex="1" value="<%=strPositionName%>"></td>
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<th align="Left">&nbsp;Position Description</th>
	    <td>&nbsp;<input style="text-align:left;width:90%" id="PositionDesc" name="PositionDesc" tabindex="2" value="<%=strPositionDesc%>"></td>
		<td colspan="2">&nbsp;</td>
	</tr>
		
	<tr>
		<th align="left">&nbsp;Active</th>		
		<td><select Style="Width:40%" tabindex="4" id="Active" name="Active"><option Value="0">Please Select....</option>
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
    <td Width="150px"><button type="button" onclick="self.location='Positions.asp?PositionID=0'" )""><img src="../images/add.png" alt="" /> Clear/Add New </button></td>
    <td Width="100px"><button type="button" onclick="DeleteData()";><img src="../images/delete.png" alt="" /> Delete </button></td>
    <!--<td Width="150px"><button type="button" onClick="StaffRecalc('<%=Session("BudgetName")%>','<%=Session("VersionName")%>');" )"" Title="Click to recalculate ALL staff data with Super Rates below (if changed)"><img src="../images/calculator.png" alt="" /> Recalculate </button></td>-->
    <TD Width="200px"><span id="Progress" style="display:none"><img src="../images/progress.gif">  &nbsp;&nbsp;&nbsp; <b>Recalculating Staff...</b></span></TD>
    <TD class='locked' Width="100px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=strMessageIcon %></TD>
    <TD class='locked' align="left" Width="800x" style="BORDER-RIGHT:0px"><INPUT style="Align:Left; font-weight:Bold; width:100%; text-align:left; color:<%=strMessageColour%>;" type="text" id="msgbox" name="msgbox" value="<%=strMessage%>"></TD>
</TR>
</TABLE>
</div>

<hr>
</form>

<table WIDTH="100%" BORDER="1" CELLSPACING="1" CELLPADDING="1">
<tr><td colspan="8">&nbsp;<span style="color:gray; font-weight:bold;">Note:</span><span style="color:gray;"> Active Positions listed here will appear in the Salary Rates screen for each Cost Centre.</span></td></tr>
	<tr>
		<th>Position Name</th>
		<th>Position Description</th>
		<th>Active</th>
		<th>Updated By</th>
		<th>Date Updated</th>
	</tr>
<%
    objRS.Open "SELECT * FROM  qryPosition WHERE BudgetID = " & clng (Session("BudgetID")) & " AND VersionID = " & Session("VersionID") & "",objCon

	Do until objRS.eof			
   	   Response.Write "<TR><TD><A Target=""_self"" HREF=""Positions.asp?PositionID=" & objRS("PositionID") & """>&nbsp;" & objRS("PositionName") & "</TD><TD style=""text-align:left"">&nbsp;" & objRS("PositionDescription") & "</TD><TD style=""text-align:center"">&nbsp;" & objRS("Active") & "</TD><TD style=""text-align:center"">&nbsp;" & objRS("UpdatedByName") & "</TD><TD style=""text-align:center"">&nbsp;" & objRS("DateUpdated") & "</TD></TR>"
       objRS.movenext
	Loop
		
	objRS.Close
	
%>

</table>
</body>

</html>

<% 

Sub LoadDetails()

       'Description:	Loads Position details into page if applicable.
		objRS.Open "SELECT * FROM tblPosition WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND PositionID = " & Session("PositionID") & "",objCon

			If Not objRS.EOF Then
				strPositionName = objRS("PositionName")
				strPositionDesc = objRS("PositionDescription")
				
				strActive = objRS("Active")
    		Else
			  
				strPositionName = ""
				strPositionDesc = ""
				
				strActive = "Y"
           End If

		objRS.Close
	
End Sub

Sub SaveDetails()

		 With objCmd
                .CommandType = 4
                .CommandText = "spPositionSave"
                
				.Parameters.Append objCmd.CreateParameter("BudgetID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("VersionID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("PositionID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("PositionName", adVarChar, adParamInput, 50)
                .Parameters.Append objCmd.CreateParameter("PositionDescription", adLongVarChar, adParamInput,-1)
                .Parameters.Append objCmd.CreateParameter("Active", adChar, adParamInput, 1)
                .Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger)
          	
				.Parameters("BudgetID") = clng(Session("BudgetID"))
   				.Parameters("VersionID") = clng(Session("VersionID"))
   				.Parameters("PositionID") = clng(Session("PositionID"))
				.Parameters("PositionName") = Request.Form("PositionName")	
				.Parameters("PositionDescription") = Request.Form("PositionDesc")	         
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

Public Sub DeleteRecord(strPositionName)
'Procedure to delete Salary Classification records if no values exist against it.

Dim intAllow
Dim intCostCentreID

	'First Check to see if any values exist for the GL Code
    objRS.Open "SELECT [PositionID],[Deleted], [CostCentreID] FROM qryStaffingClassifications With(NoLock) WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND PositionID = " & Session("PositionID") & " GROUP BY [PositionID],[CostCentreID],[Deleted]" ,objCon,adOpenStatic,adLockReadOnly
    'objRS.Open "SELECT [PositionID],[Deleted], [CostCentreID] FROM tblStaffingClassifications With(NoLock) WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND PositionID = " & Session("PositionID") & " GROUP BY [PositionID],[CostCentreID],[Deleted]" ,objCon,adOpenStatic,adLockReadOnly
            	
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
	
        objCon.Execute("DELETE FROM tblPosition WHERE [BudgetID] = " & Session("BudgetID") & " AND PositionID = " & Session("PositionID") & " AND VersionID = " & Session("VersionID") & "")
        'objCon.Execute("DELETE FROM tblGLCodes WHERE [BudgetID] = " & Session("BudgetID") & " AND GLCode = " & intGLCode & "")
        strMessage = "DELETED Position = " & strPositionName 
        'strMessage = "DELETE FROM tblSalaryClassifications WHERE [BudgetID] = " & Session("BudgetID") & " AND SalaryClassification = '" & strSalaryClassification & "' AND VersionID = " & Session("VersionID") & ""
        'Response.Write "DELETE FROM tblSalaryClassifications WHERE [BudgetID] = " & Session("BudgetID") & " AND SalaryClassification = '" & strSalaryClassification & "' AND VersionID = " & Session("VersionID") & ""
	Else
        strMessage = "NOT DELETED!. Position: " & strPositionName & " is assigned to an employee in this Budget, so CANNOT BE DELETED unless removed from Employee(s) in the Salary Rates screen (access from Salary screen)...(CCID=" & intCostCentreID & ")"
    End If

End Sub

Set objRS = Nothing
Set objCon = Nothing


%>
