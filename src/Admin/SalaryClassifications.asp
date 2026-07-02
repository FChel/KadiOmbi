<%@ Language=VBScript %>
<% Option Explicit %>
<!-- #Include file=../ADOVBS.inc -->

<%
 
'Description:	Salary Classification Editing Admin
'Author:		MG
'Date:			January 2014

'Declare default variables
Dim objCon
Dim objCmd
Dim objRS
Dim strScreen
Dim strSelected
Dim x 
Dim strMessage

'Set Database objects
Set objCon = Server.CreateObject("ADODB.Connection")
Set objCmd = Server.CreateObject("ADODB.Command")
Set objRS = Server.CreateObject("ADODB.Recordset")

'Open database connection
objCon.Open Session("DBConnection")

'1. Declare screen specific variables naming convention = variable type prefix + database field name
Dim strSalaryClassification
Dim dblSalary
Dim strActive

Dim arrYesNo(2)

	arrYesNo(1) = "Y"
	arrYesNo(2) = "N"
	
'3. Capture Querystring variables
    If Not IsEmpty(Request.QueryString("SalaryClassification")) Then
	   strSalaryClassification = Request.QueryString("SalaryClassification")
    Else
		strSalaryClassification = "None"	
    End If
		
	'Execute save 	
	If Request.QueryString("Action") = "Save" Then
		SaveDetails()
	End If

	'Execute save 	
	If Request.QueryString("Action") = "recalculate" Then
		RecalcStaff()
	End If
	
	'Execute Delete 	
	If Request.QueryString("Action") = "Delete" Then
		DeleteRecord(Request.QueryString("SalaryClassification"))
	End If
	
	'Load page details
	LoadDetails()
		
%>

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title></title>
<meta name="GENERATOR" content="Microsoft Visual Studio 6.0"/>
<link rel="stylesheet" type="text/css" href="../BERTStyle.css"/>
<script type="text/javascript" src="../formChek.js"></script>
<script type="text/javascript" language="javascript">
<!--
   function SaveData()
    {
        var varSubmit = true
        var varAlert =""  
        
        
        if((isNonnegativeInteger(frm.Salary.value)==false))
	    {            
		    varAlert += "Please enter Salary. Salary must be a numeric value.  \n \n";
		    document.getElementById('Salary').style.backgroundColor="ff8080";
		    varSubmit = false;
	    }			
	    else document.getElementById('Salary').style.backgroundColor="ffffff";	
	    
        
          	    
		    
	    if(isWhitespace(frm.SalaryClassification.value))
        {            
		    varAlert += "Please enter SalaryClassification. \n \n";
		    document.getElementById('SalaryClassification').style.backgroundColor="ff8080";
		    varSubmit = false;
	    }			
	    else document.getElementById('SalaryClassification').style.backgroundColor="ffffff";    
	    
	        
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
  
  function SalaryClassificationSearch()
{	
	self.location="SalaryClassifications.asp?SalaryClassification=" + frm.SalaryClassification.value
}

function StaffRecalc(varBud, varVer)
 {
 	if(window.confirm('Would you like to Recalculate ALL existing Staff values? \n \n For the currently selected Budget ' + varBud + ' and Version ' + varVer + '.')==true){
 	
 		document.getElementById('Progress').style.display = "inline";
        self.location='SalaryClassifications.asp?Action=recalculate';
	}
}
function DeleteData(){
	if( confirm("Delete the selected record?") )
	{
		self.location="SalaryClassifications.asp?Action=Delete&SalaryClassification="+frm.SalaryClassification.value;
		frm.elements['msgbox'].value = 'Deleting...';}
	}
//-->
</script>
</head>
<body>
<form action="SalaryClassifications.asp?Action=Save" method="POST" id="frm" name="frm">

<table width="100%" align="Center" border="1" cellspacing="1" cellpadding="1">
	<tr><th style="width:20%">Salary Classifications</th><th style="width:30%">
	<%
	Response.write "for the currently selected Budget - <FONT color=Red>" &  Session("BudgetName") & "</FONT> and Version - <FONT color=Red>" & Session("VersionName") & "</FONT>"
%>
	</th><th style="width:20%"></th><th style="width:30%"></th></tr>
	<tr><td colspan="4">&nbsp;</td></tr>
	<tr>
		<th align="Left">&nbsp;Salary Classification</th>
	    <td>&nbsp;<input style="text-align:left;width:90%" id="SalaryClassification" name="SalaryClassification" maxlength="20" tabindex="1" value="<%=strSalaryClassification%>" onblur="SalaryClassificationSearch()"/></td>
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<th align="left">&nbsp;Salary</th>
	    <td>&nbsp;<input style="text-align:left;width:90%" id="Salary" name="Salary" maxlength="50" tabindex="2" value="<%=dblSalary%>" /></td>
		<td colspan="2">&nbsp;</td>
	</tr>
	
	<tr>
		<th align="left">&nbsp;Active</th>		
		<td><select Style="Width:40%" tabindex="3" id="Active" name="Active"><option Value="0">Please Select....</option>
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
    <td Width="150px"><button type="button" onclick="self.location='SalaryClassifications.asp?SalaryClassification=0'" )""><img src="../images/add.png" alt="" /> Clear/Add New </button></td>
    <td Width="100px"><button type="button" onclick="DeleteData()";><img src="../images/delete.png" alt="" /> Delete </button></td>
    <TD Width="200px"><span id="Progress" style="display:none"><img src="../images/progress.gif">  &nbsp;&nbsp;&nbsp; <b>Recalculating Staff...</b></span></TD>
    <td Width="700px"><font Color="Red"><b><%=strMessage%></b></font></td>
</TR>
</TABLE>
</div>

<hr />
</form>

<table WIDTH="100%" BORDER="1" CELLSPACING="1" CELLPADDING="1">
<tr><td colspan="8">&nbsp;<span style="color:gray; font-weight:bold;">Note:</span><span style="color:gray;"> Changes to Salary Classifications does not update existing Staff records. The Recalculate button (above) must be clicked to recalculate existing records.</span></td></tr>
	<tr>
		<th>Salary Classification (Click to Edit)</th>
	    <th>Salary</th>
		<th>Active</th>
		<th>Updated By</th>
		<th>Date Updated</th>
	</tr>
<%
    objRS.Open "SELECT * FROM  tblSalaryClassifications WHERE BudgetID = " & clng (Session("BudgetID")) & " AND VersionID = " & Session("VersionID") & "",objCon

	Do until objRS.eof			
   	   Response.Write "<TR><TD style=""text-align:left""><A Target=""_self"" HREF=""SalaryClassifications.asp?SalaryClassification=" & objRS("SalaryClassification") & """>&nbsp;" & objRS("SalaryClassification") & "</TD><TD style=""text-align:left"">&nbsp;" & formatnumber(objRS("Salary"),0,0) & "</TD><TD style=""text-align:center"">&nbsp;" & objRS("Active") & "</TD><TD style=""text-align:center"">&nbsp;" & objRS("UpdatedBy") & "</TD><TD style=""text-align:center"">&nbsp;" & objRS("DateUpdated") & "</TD></TR>"
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
		objRS.Open "SELECT * FROM tblSalaryClassifications WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND SalaryClassification = '" & cstr(strSalaryClassification) & "'",objCon
		'Response.Write "SELECT * FROM tblSalaryClassifications WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND SalaryClassification = '" & strSalaryClassification & "'"
			If Not objRS.EOF Then
				strSalaryClassification = objRS("SalaryClassification")
				dblSalary = Round(objRS("Salary"))
				strActive = objRS("Active")
    		Else
			  
				'strSalaryClassification = ""
				dblSalary = 0
				strActive = "Y"
           End If

		objRS.Close
	
End Sub

Sub SaveDetails()

       'response.Write(Request.Form("VersionID"))

		 With objCmd
                .CommandType = 4
                .CommandText = "spSalaryClassificationsSave"
                
		.Parameters.Append objCmd.CreateParameter("BudgetID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("VersionID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("SalaryClassification", adVarChar, adParamInput, 50)
                .Parameters.Append objCmd.CreateParameter("Salary", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("Active", adChar, adParamInput, 1)
                .Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger)
          	
		.Parameters("BudgetID") = clng(Session("BudgetID"))
   		.Parameters("VersionID") = clng(Session("VersionID"))
		.Parameters("SalaryClassification") = Request.Form("SalaryClassification")		
		.Parameters("Salary") = Request.Form("Salary")             
                .Parameters("Active") = Request.Form("Active")
                .Parameters("UpdatedBy") = Session("UserID")
                           
                .ActiveConnection = objCon
                
            End With
                
            objCmd.Execute          
               
	   'Return the result of the Save Function.
     	    strMessage = "Salary Classification record saved !"
	
End Sub	


Sub RecalcStaff()

       'response.Write(Request.Form("VersionID"))

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

Public Sub DeleteRecord(strSalaryClassification)
'Procedure to delete Salary Classification records if no values exist against it.

Dim intAllow
Dim intCostCentreID

	'First Check to see if any values exist for the GL Code
    objRS.Open "SELECT [StaffClassification],Sum([BM1]+[BM2]+[BM3]+[BM4]+[BM5]+[BM6]+[BM7]+[BM8]+[BM9]+[BM10]+[BM11]+[BM12]) AS AllTotal, [CostCentreID] FROM tblStaffData With(NoLock) WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND StaffClassification = '" & strSalaryClassification & "' GROUP BY [StaffClassification],[CostCentreID]" ,objCon,adOpenStatic,adLockReadOnly
                	
        If objRS.EOF Then
        
            intAllow = 1
        Else
            intCostCentreID = objRS("CostCentreID")
            If objRS("AllTotal") = 0 Then
                intAllow = 1
            End If
        End If
                	
    objRS.Close
    
    
    'If there are no values for the account code then it can be deleted
	If intAllow = 1 Then
	
        objCon.Execute("DELETE FROM tblSalaryClassifications WHERE [BudgetID] = " & Session("BudgetID") & " AND SalaryClassification = '" & strSalaryClassification & "' AND VersionID = " & Session("VersionID") & "")
        'objCon.Execute("DELETE FROM tblGLCodes WHERE [BudgetID] = " & Session("BudgetID") & " AND GLCode = " & intGLCode & "")
        strMessage = "DELETED Salary Classification = " & strSalaryClassification 
        'strMessage = "DELETE FROM tblSalaryClassifications WHERE [BudgetID] = " & Session("BudgetID") & " AND SalaryClassification = '" & strSalaryClassification & "' AND VersionID = " & Session("VersionID") & ""
        'Response.Write "DELETE FROM tblSalaryClassifications WHERE [BudgetID] = " & Session("BudgetID") & " AND SalaryClassification = '" & strSalaryClassification & "' AND VersionID = " & Session("VersionID") & ""
	Else
        strMessage = "NOT DELETED!. Salary Classification: " & strSalaryClassification & " has values entered against it in this Budget, so CANNOT BE DELETED unless all values are removed...(CCID=" & intCostCentreID & ")"
    End If

End Sub

Set objRS = Nothing
Set objCon = Nothing


%>
