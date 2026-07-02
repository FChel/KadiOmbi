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
Dim lngCalculatedFieldType
Dim lngCalculatedFieldDesc 
Dim lngCalculatedField
Dim strSort
Dim strOrder
Dim strCalculatedField
Dim lngCalculatedFieldName 	
Dim lngActive
Dim strGLRangeStart
Dim strGLRangeEnd
Dim strClassRangeStart
Dim strClassRangeEnd
'3. Capture Querystring variables

    If Not IsEmpty(Request.QueryString("CalculatedField")) Then

	   lngCalculatedField = Request.QueryString("CalculatedField")
	 
	  ' Session("CalculatedField") = lngCalculatedField
	   If lngCalculatedField <> "" Then strCalculatedField = " AND CalculatedField='" & lngCalculatedField & "'"

	 
	Else 
		lngCalculatedField="-1"
    End If
    
    If Not IsEmpty(Request.QueryString("Sort")) Then
	   strSort = Request.QueryString("Sort")
    Else
	   strSort = "CalculatedField"
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

	    if(frm.lngCalculatedFieldType.value == "" )
	    {
		    varAlert += "Please select a Field Type. \n \n";
		    document.getElementById('lngCalculatedFieldType').style.backgroundColor="ff8080";
		    varSubmit = false;
	    }	
	    else document.getElementById('lngCalculatedFieldType').style.backgroundColor="ffffff";		   		  
	   	
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
    self.location="CalculatedFields.asp?CalculatedField=" + frm.CalculatedField.value	
}

function SaveData2(){
	var varSubmit = true
	if(document.frm.StatusID.value==0){
		alert("A Status must be selected!");
		varSubmit = false;
	}
	if(varSubmit == true){
	if ( confirm("Would you like to UPDATE all Allocation Method to Status " + document.frm.StatusID.options[document.frm.StatusID.selectedIndex].text + " ?"))
		self.location="CalculatedFields.asp?Action=SaveAll&StatusID=" + document.frm.StatusID.value;
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
<form action="CalculatedFields.asp?Action=Save&CalculatedField=<%=lngCalculatedField%>" method="POST" id="frm" name="frm">

<table width="50%" align="left" border="1" cellspacing="1" cellpadding="1">
<tr><th style="text-align:left; height:20px">&nbsp;Calculated Field</th>				
		<td align="left" >&nbsp;<input style="text-align:left; height:20px" type="text"  id="lngCalculatedField" name="lngCalculatedField" value="<%=lngCalculatedField%>"  /></td>
		</tr>

	<th style="text-align:left; height:20px; width:40%;">&nbsp;Calculated Field Type</th>
		<td style="text-align:left; height:20px; width:60%;">
		    <select Style="Width:100%" tabindex="20" id="lngCalculatedFieldType" name="lngCalculatedFieldType"><OPTION Value=0>Please Select..</OPTION>
	        <%		
		    objRS.Open "SELECT CalculatedFieldType FROM tblCalculatedFieldType",objCon

		    Do until objRS.EOF
			    If objRS("CalculatedFieldType") = lngCalculatedFieldType Then
				    strSelected = " SELECTED "
			    Else
				    strSelected = ""
			    End if
				    Response.Write "<option Value=""" & objRS("CalculatedFieldType") & """" & strSelected & ">" & objRS("CalculatedFieldType")  & "</OPTION>"
			    objRS.Movenext
		    Loop
    		
		    objRS.Close

	        %></select>
	    </td>
	  	

		<tr><th style="text-align:left; height:20px">&nbsp;Calculated Field Name </th>				
		<td align="left" >&nbsp;<input style="text-align:left; height:20px" type="text" name="lngCalculatedFieldName" value="<%=lngCalculatedFieldName%>"  /></td>
		</tr>
		<tr><th style="text-align:left; height:20px">&nbsp;Calculated Field Desc </th>				
		<td align="left" >&nbsp;<input style="text-align:left; height:20px" type="text" name="lngCalculatedFieldDesc" value="<%=lngCalculatedFieldDesc%>"  /></td>
		</tr>
        <tr><th style="text-align:left; height:20px">&nbsp;GL Range Start </th>				
		<td align="left" >&nbsp;<input style="text-align:left; height:20px" type="text" id="GLRangeStart" name="GLRangeStart" value="<%=strGLRangeStart%>"  /></td>
         <tr><th style="text-align:left; height:20px">&nbsp;GL Range End </th>				
		<td align="left" >&nbsp;<input style="text-align:left; height:20px" type="text" id="GLRangeEnd" name="GLRangeEnd" value="<%=strGLRangeEnd%>"  /></td>
		</tr>
		<tr><th style="text-align:left; height:20px">&nbsp;Class Range Start </th>				
		<td align="left" >&nbsp;<input style="text-align:left; height:20px" type="text" id="ClassRangeStart" name="ClassRangeStart" value="<%=strClassRangeStart%>"  /></td>
        <tr><th style="text-align:left; height:20px">&nbsp;Class Range End </th>				
		<td align="left" >&nbsp;<input style="text-align:left; height:20px" type="text" id="ClassRangeEnd" name="ClassRangeEnd" value="<%=strClassRangeEnd%>"  /></td>
		</tr>
		<tr><th style="text-align:left; height:20px">&nbsp;Active</th>				
		<td align="left" >&nbsp;
		<select Style="Width:100%" id="lngActive" name="lngActive">
		<%If lngActive="Y" Or lngActive="" Then%>
			<option Value="Y" Selected>Yes</option>
			<option Value="N">No</option>
		<%Else%>
			<option Value="Y">Yes</option>
			<option Value="N"  Selected>No</option>
		<%End If%>
		</select> </td>
		
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
	    <td class='locked' Width="100px" style="border-right:0px"><button type="button" tabindex="8" onClick="self.location='AdminMenu.asp';"><img src="../images/door.png" alt="" /> Close </button></td>   
		<td class='locked' Width="100px"><button type="button" tabindex="12" onClick="self.location='CalculatedFields.asp?CalculatedField=0';"><img src="../images/page_white_stack.png" alt="" /> New&nbsp;&nbsp;</button></td> 
        <td class='locked' Width="100px" style="border-right:0px"><button type="button" tabindex="9" onClick="SaveData()";><img src="../images/tick.png" alt="" /> Save </button></td>  
		
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
		"<th style=""height:20px""><B><A Target=""_self"" HREF=""CalculatedFields.asp?Sort=CalculatedField&Ordered=" & strOrder & """>Calculated Field "
		If strSort = "CalculatedField" Then Response.write "<img height=20 =width=20 src=../images/" & strOrder & ".gif Align=absmiddle>"

	response.write"</A></B></th>" & _
		"<th style=""height:20px""><B><A Target=""_self"" HREF=""CalculatedFields.asp?Sort=CalculatedFieldType&Ordered=" & strOrder & """>Calculated Field Type"
		If strSort = "CalculatedFieldType" Then Response.write "<img height=20 =width=20 src=../images/" & strOrder & ".gif Align=absmiddle>"
	
	response.write"</A></B></th>"
		response.write"</A></B></th>" & _
		"<th><B><A Target=""_self"" HREF=""CalculatedFields.asp?Sort=CalculatedFieldName&Ordered=" & strOrder & """>Calculated Field Name"
		If strSort = "CalculatedFieldName" Then Response.write "<img height=20 =width=20 src=../images/" & strOrder & ".gif Align=absmiddle>"
	
	
	
	response.write"</A></B></th>"
		response.write"</A></B></th>" & _
		"<th><B><A Target=""_self"" HREF=""CalculatedFields.asp?Sort=CalculatedFieldDesc&Ordered=" & strOrder & """>Calculated Field Desc"
		If strSort = "CalculatedFieldDesc" Then Response.write "<img height=20 =width=20 src=../images/" & strOrder & ".gif Align=absmiddle>"
	
	
	response.write"</A></B></th>"
		response.write"</A></B></th>" & _
		"<th><B><A Target=""_self"" HREF=""CalculatedFields.asp?Sort=Active&Ordered=" & strOrder & """>Active"
		If strSort = "Active" Then Response.write "<img height=20 =width=20 src=../images/" & strOrder & ".gif Align=absmiddle>"
	
	
	response.write"</A></B></th>" & _
		"<th>Updated By</th>" & _
		"<th>Date Updated</th></tr>"

    objRS.Open "SELECT * FROM tblCalculatedFields WHERE BudgetID = " & clng (Session("BudgetID"))  & " Order By " & strSort & " " & strOrder

	Do until objRS.eof	
		
   	    Response.Write "<TR><TD><A Target=""_self"" HREF=""CalculatedFields.asp?CalculatedField=" & objRS("CalculatedField") & "&Sort=" & strSort & """>&nbsp;" & objRS("CalculatedField") & "</TD><TD>&nbsp;" & objRS("CalculatedFieldType") & "</TD><TD style=""text-align:center"">" & objRS("CalculatedFieldName")  & "</TD><TD style=""text-align:center"">" & objRS("CalculatedFieldDesc") & "</TD>	<TD style=""text-align:center"">" & objRS("Active") & "</TD> <TD style=""text-align:center"">&nbsp;" & objRS("UpdatedBy") & "</TD><TD style=""text-align:center"">" & objRS("DateUpdated") & "</TD></TR>"
       
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
	
		objRS.Open "SELECT * FROM tblCalculatedFields WHERE BudgetID = " & Session("BudgetID")  & " And CalculatedField='" & lngCalculatedField & "'",objCon
		
			If Not objRS.EOF Then
			
			  lngCalculatedField = objRS("CalculatedField")
              lngBudgetID = objRS("BudgetID")
              lngCalculatedFieldType = objRS("CalculatedFieldType")
              lngCalculatedFieldDesc  = objRS("CalculatedFieldDesc")
 			  lngCalculatedFieldName  = objRS("CalculatedFieldName")
			  lngActive= objRS("Active")
              strGLRangeStart = objRS("GLRangeStart")
              strGLRangeEnd = objRS("GLRangeEnd")
              strClassRangeStart = objRS("ClassRangeStart")
              strClassRangeEnd = objRS("ClassRangeEnd")

			Else
			 lngCalculatedFieldDesc  = ""
				
			End If

		objRS.Close
			

End Sub

Sub SaveDetails()

	 With objCmd
                .CommandType = 4
                .CommandText = "spCalculatedFieldSave"   
                    
                
                .Parameters.Append objCmd.CreateParameter("BudgetID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("CalculatedField", adVarChar, adParamInput,50)                                
                .Parameters.Append objCmd.CreateParameter("CalculatedFieldType", adVarChar, adParamInput,50)
                .Parameters.Append objCmd.CreateParameter("CalculatedFieldName", adVarChar, adParamInput,50)
		        .Parameters.Append objCmd.CreateParameter("CalculatedFieldDesc", adLongVarChar, adParamInput,1000) 
                .Parameters.Append objCmd.CreateParameter("SQL", adLongVarChar, adParamInput,1000)
		        .Parameters.Append objCmd.CreateParameter("Active", adChar, adParamInput,1)
		        .Parameters.Append objCmd.CreateParameter("Formula", adChar, adParamInput,1)
		        .Parameters.Append objCmd.CreateParameter("Source", adVarChar, adParamInput,50)
                .Parameters.Append objCmd.CreateParameter("GLRangeStart", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("GLRangeEnd", adInteger, adParamInput)
                 .Parameters.Append objCmd.CreateParameter("ClassRangeStart", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("ClassRangeEnd", adInteger, adParamInput)      
                    
                .Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger)                                                                    
              			
		        .Parameters("BudgetID") = Session("BudgetID")		
                .Parameters("CalculatedField") = cstr(Request.Form("lngCalculatedField"))                     
                .Parameters("CalculatedFieldType") = cstr(Request.Form("lngCalculatedFieldType"))
                .Parameters("CalculatedFieldName") = cstr(Request.Form("lngCalculatedFieldName"))
		        .Parameters("CalculatedFieldDesc") = cstr(Request.Form("lngCalculatedFieldDesc"))
		.Parameters("SQL") = ""
		.Parameters("Active") = Request.Form("lngActive")
		.Parameters("Formula") = ""
		.Parameters("Source") = ""
        .Parameters("GLRangeStart") = Request.Form("GLRangeStart")
        .Parameters("GLRangeEnd") = Request.Form("GLRangeEnd")
        .Parameters("ClassRangeStart") = Request.Form("ClassRangeStart")
        .Parameters("ClassRangeEnd") = Request.Form("ClassRangeEnd")



                .Parameters("UpdatedBy") = Session("UserID")
                           
                .ActiveConnection = objCon
                
            End With
                
            objCmd.Execute          
  
      
	 ' lngCalculatedFieldName = Request.Form("lngCalculatedFieldName ")
	'response.write "UPDATE tblCalculatedFields SET CalculatedFieldType='" & Request.Form("lngCalculatedFieldType") & "', CalculatedFieldDesc  = '" & Request.Form("lngCalculatedFieldDesc") & "', CalculatedFieldName =" & request.form("lngCalculatedFieldName")  & ", UpdatedBy = " & clng(Session("UserID")) & ", Active='" & Request.Form("lngActive")  & "', DateUpdated = GetDate() WHERE  BudgetID = " & clng(Session("BudgetID")) & " And CalculatedField=" &   Request.Form("lngCalculatedField") & ""
	'response.End()
	  
	  'objCon.Execute("UPDATE tblCalculatedFields SET CalculatedFieldType='" & Request.Form("lngCalculatedFieldType") & "', CalculatedFieldDesc  = '" & Request.Form("lngCalculatedFieldDesc") & "', CalculatedFieldName ='" & request.form("lngCalculatedFieldName")  & "', UpdatedBy = " & clng(Session("UserID")) & ", Active='" & Request.Form("lngActive")  & "', DateUpdated = GetDate() WHERE  BudgetID = " & clng(Session("BudgetID")) & " And CalculatedField='" &   Request.Form("lngCalculatedField") & "'") 
	  
			'Return the result of the Save Function.
     		strMessage = "<B>Record Saved.</B>"
            strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"
     		
     
	 
	

	 
End Sub

Sub SaveAllRecord(StatusID)

  objCon.Execute("UPDATE tblCalculatedFields SET CalculatedFieldDesc  = '" & StatusID & "', UpdatedBy = " & clng(Session("UserID")) & ", DateUpdated = GetDate() WHERE  BudgetID = " & clng(Session("BudgetID"))  & "") 
 ' objCon.Execute("UPDATE tblCostCentreStatus SET StatusID = " & clng(StatusID) & ", UpdatedBy = " & clng(Session("UserID")) & ", DateUpdated = GetDate() WHERE  BudgetID = " & clng(Session("BudgetID")) & " and VersionID = " & clng(Session("VersionID")) & "") 
  strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"
  strMessage = "<B>All ALLOCATION METHOD HAVE BEEN UPDATED.</B>"

End Sub	


Set objRS = Nothing
Set objCon = Nothing


%>
