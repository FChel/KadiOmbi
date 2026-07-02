<%@ Language=VBScript %>
<% Option Explicit %>
<!-- #Include file="../ADOVBS.inc" -->

<%

Response.Expires = -1500

If IsEmpty(Session("BudgetID")) Then Response.Redirect("../Timeout.asp")
 
'Description:	Business Area Status Screen
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
Dim arrStatus(5)


arrStatus(0) = "<IMG SRC='../images/delete.png'"
arrStatus(1) = "<IMG SRC='../images/open.png'"
arrStatus(2) = "<IMG SRC='../images/ready.gif'"
arrStatus(3) = "<IMG SRC='../images/cross.png'" 
arrStatus(4) = "<IMG SRC='../images/tick.png'"	
arrStatus(5) = "<IMG SRC='../images/Closed.png'"

'Set Database objects

Set objCon = Server.CreateObject("ADODB.Connection")
Set objCmd = Server.CreateObject("ADODB.Command")
Set objRS = Server.CreateObject("ADODB.Recordset")

'Open database connection

objCon.Open Session("DBConnection")

'1. Declare screen specific variables naming convention = variable type prefix + database field name
Dim lngRecordID
Dim lngBudgetID
Dim lngCeilingLevelName
Dim lngAllocationMethod
Dim lngCeilingLevelID
Dim strSort
Dim strOrder
Dim strCeilingLevelID
Dim lngAllocationValue	
Dim lngMandatory
Dim strCalculatedField

'3. Capture Querystring variables

    If Not IsEmpty(Request.QueryString("CeilingLevelID")) Then

	   lngCeilingLevelID = Request.QueryString("CeilingLevelID")
	 
	  Session("CeilingLevelID") = lngCeilingLevelID
	   'If lngCeilingLevelID <> "0" Then strCeilingLevelID = " AND CeilingLevelID='" & lngCeilingLevelID & "'"

	 
	Else 
		lngCeilingLevelID=1
    End If
    
    If Not IsEmpty(Request.QueryString("Sort")) Then
	   strSort = Request.QueryString("Sort")
    Else
	   strSort = "CeilingLevelID"
    End If

    If Not IsEmpty(Request.QueryString("Ordered")) Then
	If Request.QueryString("Ordered") = "asc" Then
		strOrder = "desc"
	Else
	   strOrder = "asc"
	End If
    Else
	   strOrder = "asc"
    End If

	'Execute save 	
	If Request.QueryString("Action") = "Save" Then
		SaveDetails()
	End If

	'Call the Save procedure to Update ALL Business Areas.
	If Request.QueryString("Action") = "SaveAll" Then
		SaveAllRecord(Request.QueryString("StatusID"))
	End If

    'Call the Save procedure to Update ALL Business Areas.
	If Request.QueryString("Action") = "Reset" Then
		Reset_Ceilings("N")
	End If

     'Call the Save procedure to Update ALL Business Areas.
	If Request.QueryString("Action") = "ResetDelete" Then
		Reset_Ceilings("Y")
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
<script type="text/javascript" src="../ButtonRollOver.js"></script>
<script type="text/javascript" language="javascript">
<!--
   function SaveData()
    {
        var varSubmit = true
        var varAlert =""       

	    if(frm.StatusID.value == 0 )
	    {
		    varAlert += "Please select a status. \n \n";
		    document.getElementById('StatusID').style.backgroundColor="ff8080";
		    varSubmit = false;
	    }	
	    else document.getElementById('StatusID').style.backgroundColor="ffffff";		   		  
	   	
	  if(varSubmit == true)
	  {
	        frm.submit();
	  }
	  else
	  {
	    window.alert ("" + varAlert);	    
	  }
	  
    }

   function DeleteData()
    {
 	if(window.confirm('Confirm delete')==true){
	self.location="Version.asp?Action=Delete"
	}
}

function BAIDSearch()
{	 
    self.location="CeilingsLevel2.asp?CeilingLevelID=" + frm.CeilingLevelID.value	
}

function SaveData2(){
	var varSubmit = true
	if(document.frm.StatusID.value==0){
		alert("A Status must be selected!");
		varSubmit = false;
	}
	if(varSubmit == true){
	if ( confirm("Would you like to UPDATE all to Allocation Method " + document.frm.StatusID.options[document.frm.StatusID.selectedIndex].text + " ?"))
		self.location="CeilingsLevel2.asp?Action=SaveAll&StatusID=" + document.frm.StatusID.value;
	}else{
		//alert("Status NOT Updated!");
	}

}

function Reset() {
    var varSubmit = true

    if (varSubmit == true) {
        if (confirm("Would you like to RESET the Account Class Ceilings ?"))
            var answer = confirm("Delete existing ceilings?")
        if (answer) {
            self.location = "CeilingsLevel2.asp?Action=ResetDelete";
        }
        else {
            self.location = "CeilingsLevel2.asp?Action=Reset";
        }
            
    } else {

    }
}

//-->
</script>
</head>
<body onload=padding();>
<h3>Account Class Ceilings Administration Screen
<%
	Response.write "for the currently selected Budget : <FONT color=Red>" &  Session("BudgetName") & "</FONT> and Version : <FONT color=Red>" & Session("VersionName") & "</FONT>"
%>
</h3>
<form action="CeilingsLevel2.asp?Action=Save" method="POST" id="frm" name="frm">

<table width="50%" align="left" border="1" cellspacing="1" cellpadding="1">
<tr><th style="text-align:left; height:20px">&nbsp;Ceiling Level ID</th>				
		<td align="left" >&nbsp;<input style="text-align:left; height:20px" type="text"  name="lngCeilingLevelID" value="<%=lngCeilingLevelID%>" /></td>
		</tr>
<tr>
	<th style="text-align:left; height:20px; width:40%;">&nbsp;Ceiling Level Name</th>
		<td style="text-align:left; height:20px; width:60%;">
		    <select Style="Width:100%" tabindex="20" id="lngCeilingLevelName" name="lngCeilingLevelName"><OPTION Value=0>Please Select..</OPTION>
	        <%	
		   objRS.Open "SELECT * FROM tblCalculatedFields WHERE BudgetID = " & Session("BudgetID") & " AND CalculatedFieldType In ('Custom','MTFF') AND Active = 'Y'",objCon

		    Do until objRS.EOF
			    If objRS("CalculatedField") = strCalculatedField  Then
				    strSelected = " SELECTED "
			    Else
				    strSelected = ""
			    End if
				    Response.Write "<option Value=""" & objRS("CalculatedField") & """" & strSelected & ">" & objRS("CalculatedField")  & " - " & objRS("CalculatedFieldName")  & "</OPTION>"
			    objRS.Movenext
		    Loop
    		
		    objRS.Close
    		
	        %></select>
	    </td>
	  	
	<tr>   
	
		<th style="text-align:left; height:20px">&nbsp;Allocation Method</th>		
		<td><select Style="Width:100%" tabindex="6" id="StatusID" name="StatusID"><option Value="">Please Select....</option>
		<%
		objRS.Open "SELECT * FROM tblAllocationMethods WHERE BudgetID = " & Session("BudgetID") & " AND Active = 'Y'",objCon
		
		Do until objRS.EOF
			If objRS("AllocationMethodName") = lngAllocationMethod Then
				strSelected = " SELECTED "
			Else
				strSelected = ""
			End if
				Response.Write "<option Value=""" & objRS("AllocationMethodName") & """" & strSelected & ">" & objRS("AllocationMethodName") & "</OPTION>"
			objRS.Movenext
		Loop
		
		objRS.Close
		%>
		</select> </td>
		</tr>
		<tr><th style="text-align:left; height:20px">&nbsp;Allocation Value</th>				
		<td align="left" >&nbsp;<input style="text-align:left; height:20px" type="number" step='0.01' name="lngAllocationValue" value="<%=lngAllocationValue%>" placeholder='0.00' /></td>
		</tr>
		<tr><th style="text-align:left; height:20px">&nbsp;Mandatory</th>				
		<td align="left" >&nbsp;
		<select Style="Width:100%" id="lngMandatory" name="lngMandatory">
		<%If lngMandatory="Y" Or lngMandatory="" Then%>
			<option Value="Y" Selected>Yes</option>
			<option Value="N">No</option>
		<%Else%>
			<option Value="Y">Yes</option>
			<option Value="N" Selected>No</option>
		<%End If%>
		</select> </td>
		
		</tr>
	<tr>
		<td style="height:20px" colspan="4" align="left">&nbsp;</td>
	</tr>
	
</table>
<br/>
<br/>
<br/>
<br/>
<br/>
<br/>
<br/>
<br/>



<table Class="CallCentre" WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
	<tr>
	    <td class='locked' Width="100px" style="border-right:0px"><button type="button" tabindex="8" onClick="self.location='AdminMenu.asp';"><img src="../images/door.png" alt="" /> Close </button></td> 
        <td class='locked' Width="100px" style="border-right:0px"><button type="button" tabindex="9" onClick="SaveData()";><img src="../images/tick.png" alt="" /> Save </button></td>  
		<td class='locked' Width="150px"><button type="button" tabindex="13" onClick="javascript:SaveData2();"><img src="../images/table_save.png" alt="" /> Apply Method to All</button></td>
		<td class='locked' Width="150px"><button type="button" tabindex="14" onClick="javascript:Reset();"><img src="../images/table_save.png" alt="" /> Reset Ceilings</button></td>
        <TD class='locked' Width="100px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=strMessageIcon %></TD>
        <TD class='locked' Width="400px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=strMessage %></TD>
	</tr>
</table>
<hr />
</form>

<table WIDTH="100%" BORDER="1" CELLSPACING="1" CELLPADDING="1">
	<tr>
<%
	'Dynamically build the menu items depending on the sort selection 

    response.write"</A></B></th>" & _
		"<th style=""height:20px""><B><A Target=""_self"" HREF=""CeilingsLevel2.asp?Sort=CeilingLevelID&Ordered=" & strOrder & """>Ceiling Level ID "
		If strSort = "ProgramCode" Then Response.write "<img height=20 =width=20 src=../images/" & strOrder & ".gif Align=absmiddle>"

	response.write"</A></B></th>" & _
		"<th style=""height:20px""><B><A Target=""_self"" HREF=""CeilingsLevel2.asp?Sort=CeilingLevelName&Ordered=" & strOrder & """>Ceiling Level Name"
		If strSort = "CostCentreName" Then Response.write "<img height=20 =width=20 src=../images/" & strOrder & ".gif Align=absmiddle>"
	
	response.write"</A></B></th>"
		response.write"</A></B></th>" & _
		"<th><B><A Target=""_self"" HREF=""CeilingsLevel2.asp?Sort=AllocationMethod&Ordered=" & strOrder & """>Allocation Method"
		If strSort = "AllocationMethod" Then Response.write "<img height=20 =width=20 src=../images/" & strOrder & ".gif Align=absmiddle>"
	
	
	response.write"</A></B></th>"
		response.write"</A></B></th>" & _
		"<th><B><A Target=""_self"" HREF=""CeilingsLevel2.asp?Sort=AllocationValue&Ordered=" & strOrder & """>Allocation Value"
		If strSort = "AllocationValue" Then Response.write "<img height=20 =width=20 src=../images/" & strOrder & ".gif Align=absmiddle>"
	
	
	
	response.write"</A></B></th>" & _
		"<th>Updated By</th>" & _
		"<th>Date Updated</th></tr>"

    objRS.Open "SELECT * FROM tblCeilingsLevel2 WHERE BudgetID = " & clng (Session("BudgetID"))  & " Order By " & strSort & " " & strOrder

	Do until objRS.eof	
		
   	    Response.Write "<TR><TD><A Target=""_self"" HREF=""CeilingsLevel2.asp?CeilingLevelID=" & objRS("CeilingLevelID") & "&Sort=" & strSort & """>&nbsp;" & objRS("CeilingLevelID") & "</TD><TD>&nbsp;" & objRS("CeilingLevelName") & "</TD><TD style=""text-align:center"">" & objRS("AllocationMethod") & "</TD>	<TD style=""text-align:center"">" & objRS("AllocationValue")  & "</TD><TD style=""text-align:center"">&nbsp;" & objRS("UpdatedBy") & "</TD><TD style=""text-align:center"">" & objRS("DateUpdated") & "</TD></TR>"
       
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
	
		objRS.Open "SELECT * FROM tblCeilingsLevel2 WHERE BudgetID = " & Session("BudgetID")  & " And CeilingLevelID='" & lngCeilingLevelID & "'",objCon
		
			If Not objRS.EOF Then
			
			  lngCeilingLevelID = objRS("CeilingLevelID")
			  lngBudgetID = objRS("BudgetID")
              lngCeilingLevelName = objRS("CeilingLevelName")
              lngAllocationMethod = objRS("AllocationMethod")
 			  lngAllocationValue = objRS("AllocationValue")
			  lngMandatory= objRS("Mandatory")          

			Else
			 lngAllocationMethod = ""
				
			End If

		objRS.Close

End Sub

Sub SaveDetails()

		
      If Request.Form("StatusID")<>"" Then
      
	  lngAllocationValue= Request.Form("lngAllocationValue")
	  If Request.Form("lngAllocationValue")="" Then lngAllocationValue=0
	  
	  objCon.Execute("UPDATE tblCeilingsLevel2 SET CeilingLevelName='" & Request.Form("lngCeilingLevelName") & "', AllocationMethod = '" & Request.Form("StatusID") & "', AllocationValue=" & lngAllocationValue & ", UpdatedBy = " & clng(Session("UserID")) & ", Mandatory='" & Request.Form("lngMandatory")  & "', DateUpdated = GetDate() WHERE  BudgetID = " & clng(Session("BudgetID")) & " And CeilingLevelID=" &   Request.Form("lngCeilingLevelID") & "") 
	  
			'Return the result of the Save Function.
     		strMessage = "<B>Record Saved.</B>"
            strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"
     		
     
	 
	 End If

	 
End Sub

Sub SaveAllRecord(StatusID)

  objCon.Execute("UPDATE tblCeilingsLevel2 SET AllocationMethod = '" & StatusID & "', UpdatedBy = " & clng(Session("UserID")) & ", DateUpdated = GetDate() WHERE  BudgetID = " & clng(Session("BudgetID"))  & "") 
 ' objCon.Execute("UPDATE tblCostCentreStatus SET StatusID = " & clng(StatusID) & ", UpdatedBy = " & clng(Session("UserID")) & ", DateUpdated = GetDate() WHERE  BudgetID = " & clng(Session("BudgetID")) & " and VersionID = " & clng(Session("VersionID")) & "") 
  strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"
  strMessage = "<B>All ALLOCATION METHOD HAVE BEEN UPDATED.</B>"

End Sub	

Sub Reset_Ceilings(Delete)

  objCon.Execute "spApplyBudgetClassCeilings " & Session("BudgetID") & "," & Session("VersionID") & ",'" & Delete & "'," & Session("UserID") & ""
 ' objCon.Execute("UPDATE tblCostCentreStatus SET StatusID = " & clng(StatusID) & ", UpdatedBy = " & clng(Session("UserID")) & ", DateUpdated = GetDate() WHERE  BudgetID = " & clng(Session("BudgetID")) & " and VersionID = " & clng(Session("VersionID")) & "") 
  strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"
  strMessage = "<B>ACCOUNT CLASS CEILINGS HAVE BEEN RESET.</B>"

End Sub	


Set objRS = Nothing
Set objCon = Nothing


%>
