<%@ Language=VBScript %>
<% Option Explicit %>
<!-- #Include file="ADOVBS.inc" -->

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
Dim lngVersionID
Dim lngAllocationMethod
Dim lngCostCentreID
Dim strSort
Dim strOrder
Dim strCostCentreID
Dim lngAllocationValue	
'3. Capture Querystring variables

    If Not IsEmpty(Request.QueryString("CostCentreID")) Then
	   lngCostCentreID = clng(Request.QueryString("CostCentreID"))
	   Session("CostCentreID") = lngCostCentreID
	   If lngCostCentreID <> 0 Then strCostCentreID = "AND CostCentreID=" & lngCostCentreID
    End If
    
    If Not IsEmpty(Request.QueryString("Sort")) Then
	   strSort = Request.QueryString("Sort")
    Else
	   strSort = "CostCentreID"
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
    self.location="CostCentreCeilings.asp?CostCentreID=" + frm.CostCentreID.value	
}

function SaveData2(){
	var varSubmit = true
	if(document.frm.StatusID.value==0){
		alert("A Status must be selected!");
		varSubmit = false;
	}
	if(varSubmit == true){
	if ( confirm("Would you like to UPDATE all Allocation Method to Status " + document.frm.StatusID.options[document.frm.StatusID.selectedIndex].text + " ?"))
		self.location="CostCentreCeilings.asp?Action=SaveAll&StatusID=" + document.frm.StatusID.value;
	}else{
		//alert("Status NOT Updated!");
	}

}

//-->
</script>
</head>
<body onload=padding();>
<h3>Cost Centre Ceilings Administration Screen
<%
	Response.write "for the currently selected Budget : <FONT color=Red>" &  Session("BudgetName") & "</FONT> and Version : <FONT color=Red>" & Session("VersionName") & "</FONT>"
%>
</h3>
<form action="CostCentreCeilings.asp?Action=Save" method="POST" id="frm" name="frm">

<table width="50%" align="left" border="1" cellspacing="1" cellpadding="1">
	<th style="text-align:left; height:20px; width:40%;">&nbsp;Cost Centre</th>
		<td style="text-align:left; height:20px; width:60%;">
		    <select Style="Width:100%" tabindex="20" id="CostCentreID" name="CostCentreID" onChange="BAIDSearch()"><OPTION Value=0>Please Select..</OPTION>
	        <%	
		    objRS.Open "SELECT * FROM tblCostCentres WHERE BudgetID = " & Session("BudgetID") & " AND Active = 'Y'",objCon

		    Do until objRS.EOF
			    If objRS("CostCentreID") = clng(Session("CostCentreID")) Then
				    strSelected = " SELECTED "
			    Else
				    strSelected = ""
			    End if
				    Response.Write "<option Value=""" & objRS("CostCentreID") & """" & strSelected & ">" & objRS("ProgramCode") & " - " & objRS("CostCentreName") & "</OPTION>"
			    objRS.Movenext
		    Loop
    		
		    objRS.Close
    		
	        %></select>
	    </td>
	  	
	<tr>
	
		<th style="text-align:left; height:20px">&nbsp;Status</th>		
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
		<tr><th style="text-align:left; height:20px">&nbsp;Alloocation Value</th>				
		<td align="left" >&nbsp;<input style="text-align:left; height:20px" type="number" step='0.01' name="lngAllocationValue" value="<%=lngAllocationValue%>" placeholder='0.00' /></td>
		</tr>
	</tr>

	
	
	<tr>
		<td style="height:20px" colspan="4" align="left">&nbsp;</td>
	</tr>
	
</table>
<br/>
<br/>
<br/>
<br/>

<table Class="CallCentre" WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
	<tr>
	    <td class='locked' Width="100px" style="border-right:0px"><button type="button" tabindex="8" onClick="self.location='../AdminMenu.asp';"><img src="../images/door.png" alt="" /> Close </button></td>    
        <td class='locked' Width="100px" style="border-right:0px"><button type="button" tabindex="9" onClick="SaveData()";><img src="../images/tick.png" alt="" /> Save </button></td>  
		<td class='locked' Width="100px"><button type="button" tabindex="13" onClick="javascript:SaveData2();"><img src="../images/table_save.png" alt="" /> Apply Status to All</button></td>
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
		"<th style=""height:20px""><B><A Target=""_self"" HREF=""CostCentreCeilings.asp?Sort=ProgramCode&Ordered=" & strOrder & """>Program Code"
		If strSort = "ProgramCode" Then Response.write "<img height=20 =width=20 src=../images/" & strOrder & ".gif Align=absmiddle>"

	response.write"</A></B></th>" & _
		"<th style=""height:20px""><B><A Target=""_self"" HREF=""CostCentreCeilings.asp?Sort=CostCentreName&Ordered=" & strOrder & """>Cost Centre Name"
		If strSort = "CostCentreName" Then Response.write "<img height=20 =width=20 src=../images/" & strOrder & ".gif Align=absmiddle>"
	
	response.write"</A></B></th>"
		response.write"</A></B></th>" & _
		"<th><B><A Target=""_self"" HREF=""CostCentreCeilings.asp?Sort=AllocationMethod&Ordered=" & strOrder & """>Allocation Method"
		If strSort = "AllocationMethod" Then Response.write "<img height=20 =width=20 src=../images/" & strOrder & ".gif Align=absmiddle>"
	
	
	response.write"</A></B></th>"
		response.write"</A></B></th>" & _
		"<th><B><A Target=""_self"" HREF=""CostCentreCeilings.asp?Sort=AllocationValue&Ordered=" & strOrder & """>Allocation Value"
		If strSort = "AllocationValue" Then Response.write "<img height=20 =width=20 src=../images/" & strOrder & ".gif Align=absmiddle>"
	
	
	
	response.write"</A></B></th>" & _
		"<th>Updated By</th>" & _
		"<th>Date Updated</th></tr>"

    objRS.Open "SELECT * FROM qryCostCentreCeilings WHERE BudgetID = " & clng (Session("BudgetID")) & " AND VersionID = " & clng (Session("VersionID")) & " Order By " & strSort & " " & strOrder

	Do until objRS.eof	
		
   	    Response.Write "<TR><TD><A Target=""_self"" HREF=""CostCentreCeilings.asp?CostCentreID=" & objRS("CostCentreID") & "&Sort=" & strSort & """>&nbsp;" & objRS("ProgramCode") & "</TD><TD>&nbsp;" & objRS("CostCentreName") & "</TD><TD style=""text-align:center"">" & objRS("AllocationMethod") & "</TD>	<TD style=""text-align:center"">" & objRS("AllocationValue")  & "</TD><TD style=""text-align:center"">&nbsp;" & objRS("UpdatedBy") & "</TD><TD style=""text-align:center"">" & objRS("DateUpdated") & "</TD></TR>"
       
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

		objRS.Open "SELECT * FROM tblCostCentreCeilings WHERE BudgetID = " & Session("BudgetID") & " AND CostCentreID = " & Session("CostCentreID") & "",objCon
		
			If Not objRS.EOF Then
			
			  lngCostCentreID = objRS("CostCentreID")
			  lngBudgetID = objRS("BudgetID")
              lngVersionID = objRS("VersionID")
              lngAllocationMethod = objRS("AllocationMethod")
 			  lngAllocationValue = objRS("AllocationValue")
			Else
			 lngAllocationMethod = ""
				
			End If

		objRS.Close


End Sub

Sub SaveDetails()

		
      If Request.Form("StatusID")<>"" Then
      
	  lngAllocationValue= Request.Form("lngAllocationValue")
	  If Request.Form("lngAllocationValue")="" Then lngAllocationValue=0
	  
	  objCon.Execute("UPDATE tblCostCentreCeilings SET AllocationMethod = '" & Request.Form("StatusID") & "', AllocationValue=" & lngAllocationValue & ", UpdatedBy = " & clng(Session("UserID")) & ", DateUpdated = GetDate() WHERE  BudgetID = " & clng(Session("BudgetID")) & " and VersionID = " & clng(Session("VersionID")) & " And CostCentreID=" &   Request.Form("CostCentreID") & "") 
	  
			'Return the result of the Save Function.
     		strMessage = "<B>Record Saved.</B>"
            strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"
     		
     
	 
	 End If

	 
End Sub

Sub SaveAllRecord(StatusID)

  objCon.Execute("UPDATE tblCostCentreCeilings SET AllocationMethod = '" & StatusID & "', UpdatedBy = " & clng(Session("UserID")) & ", DateUpdated = GetDate() WHERE  BudgetID = " & clng(Session("BudgetID")) & " and VersionID = " & clng(Session("VersionID")) & "") 
 ' objCon.Execute("UPDATE tblCostCentreStatus SET StatusID = " & clng(StatusID) & ", UpdatedBy = " & clng(Session("UserID")) & ", DateUpdated = GetDate() WHERE  BudgetID = " & clng(Session("BudgetID")) & " and VersionID = " & clng(Session("VersionID")) & "") 
  strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"
  strMessage = "<B>All ALLOCATION METHOD HAVE BEEN UPDATED.</B>"

End Sub	


Set objRS = Nothing
Set objCon = Nothing


%>
