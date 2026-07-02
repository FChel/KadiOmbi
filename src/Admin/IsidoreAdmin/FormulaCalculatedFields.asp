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

Dim lngFormulaName
Dim lngCalculatedFieldName
Dim lngPeriod

Dim lngVersionID


Dim strSort
Dim strOrder
Dim strCalculatedField

Dim lngBM(24)
Dim lngOY(5)

Dim z
Dim BMUpdate
Dim OYUpdate
Dim BMInsert1
DIM	BMInsert2
DIM	OYInsert1
DIM	OYInsert2

'3. Capture Querystring variables

    If Not IsEmpty(Request("lngFormulaName")) Then
	   	lngFormulaName = Request("lngFormulaName")
		lngCalculatedFieldName=request("lngCalculatedFieldName")
		lngPeriod=request("lngPeriod")

	  ' Session("CalculatedFieldName") = lngCalculatedFieldName
	   'If lngCalculatedFieldName <> "" Then strCalculatedField = " AND CalculatedFieldName='" & lngCalculatedFieldName & "'"

	 
	Else 

		lngFormulaName="Actuals Trend"
	    lngCalculatedFieldName="AM1"
		lngPeriod="BM12"
    End If
    
    If Not IsEmpty(Request.QueryString("Sort")) Then
	   strSort = Request.QueryString("Sort")
    Else
	   strSort = "FormulaName"
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
	
	frm.submit();
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
    self.location="FormulaCalculatedFields.asp?lngFormulaName=" + frm.lngFormulaName.value	
}

function SaveData2(){
	var varSubmit = true
	if(document.frm.StatusID.value==0){
		alert("A Status must be selected!");
		varSubmit = false;
	}
	if(varSubmit == true){
	if ( confirm("Would you like to UPDATE all Allocation Method to Status " + document.frm.StatusID.options[document.frm.StatusID.selectedIndex].text + " ?"))
		self.location="FormulaCalculatedFields.asp?Action=SaveAll&StatusID=" + document.frm.StatusID.value;
	}else{
		//alert("Status NOT Updated!");
	}

}

//-->
</script>
</head>
<body onload=padding();>
<h3>Formula Calculated Fields Administration Screen
<%
	Response.write "for the currently selected Budget : <FONT color=Red>" &  Session("BudgetName") & "</FONT> and Version : <FONT color=Red>" & Session("VersionName") & "</FONT>"
%>
</h3>
<form action="FormulaCalculatedFields.asp?Action=Save" method="POST" id="frm" name="frm">

<table width="50%" align="left" border="1" cellspacing="1" cellpadding="1">

		
				
	<tr>
	<th style="text-align:left; height:20px; width:40%;">&nbsp;Formula Name:</th>
		<td style="text-align:left; height:20px; width:60%;">
		    <select Style="Width:100%" tabindex="20" id="lngFormulaName" name="lngFormulaName"><OPTION Value=0>Please Select..</OPTION>
	        <%		
		    objRS.Open "SELECT FormulaName FROM tblFormulas WHERE BudgetID =" & Session("BudgetID") ,objCon

		    Do until objRS.EOF
			    If trim(objRS("FormulaName")) = trim(lngFormulaName) Then
				    strSelected = " SELECTED "
			    Else
				    strSelected = ""
			    End if
				    Response.Write "<option Value=""" & objRS("FormulaName") & """" & strSelected & ">" & objRS("FormulaName")  & "</OPTION>"
			    objRS.Movenext
		    Loop
    		
		    objRS.Close

	        %></select>
	    </td>
	  	</tr>

	<tr>
	<th style="text-align:left; height:20px; width:40%;">&nbsp;Calculated Field Name:</th>
		<td style="text-align:left; height:20px; width:60%;">
		    <select Style="Width:100%" tabindex="20" id="lngCalculatedFieldName" name="lngCalculatedFieldName"><OPTION Value=0>Please Select..</OPTION>
	        <%		
		    objRS.Open "SELECT CalculatedFieldName FROM tblCalculatedFields WHERE BudgetID =" & Session("BudgetID") ,objCon

		    Do until objRS.EOF
			    If objRS("CalculatedFieldName") = lngCalculatedFieldName Then
				    strSelected = " SELECTED "
			    Else
				    strSelected = ""
			    End if
				    Response.Write "<option Value=""" & objRS("CalculatedFieldName") & """" & strSelected & ">" & objRS("CalculatedFieldName")  & "</OPTION>"
			    objRS.Movenext
		    Loop
    		
		    objRS.Close
		
	        %></select>
	    </td>
	  	</tr>

		<tr><th style="text-align:left; height:20px">&nbsp;Period</th>				
		<td align="left" >&nbsp;
		<select Style="Width:100%" id="lngPeriod" name="lngPeriod">

			<option Value="FVal" Selected>FVal</option>
			<option Value="CY">CY</option>
			<option Value="BM1">BM1</option>
			<option Value="BM2">BM2</option>
			<option Value="BM3">BM3</option>
			<option Value="BM4">BM4</option>
			<option Value="BM5">BM5</option>
			<option Value="BM6">BM6</option>
			<option Value="BM7">BM7</option>
			<option Value="BM8">BM8</option>
			<option Value="BM9">BM9</option>
			<option Value="BM10">BM10</option>
			<option Value="BM11">BM11</option>
			<option Value="BM12">BM12</option>
			<option Value="BM13">BM13</option>
			<option Value="BM14">BM14</option>
			<option Value="BM15">BM15</option>
			<option Value="BM16">BM16</option>
			<option Value="BM17">BM17</option>
			<option Value="BM18">BM18</option>
			<option Value="BM20">BM20</option>
			<option Value="BM21">BM21</option>
			<option Value="BM22">BM22</option>
			<option Value="BM23">BM23</option>
			<option Value="BM24">BM24</option>
			<option Value="OY1">OY1</option>
			<option Value="OY2">OY2</option>
			<option Value="OY3">OY3</option>
			<option Value="OY4">OY4</option>
			<option Value="OY5">OY5</option>
			<option Value="AM1">AM1</option>
			<option Value="AM2">AM2</option>
			<option Value="AM3">AM3</option>
			<option Value="AM4">AM4</option>
			<option Value="AM5">AM5</option>
			<option Value="AM6">AM6</option>
			<option Value="AM7">AM7</option>
			<option Value="AM8">AM8</option>
			<option Value="AM9">AM9</option>
			<option Value="AM10">AM10</option>
			<option Value="AM11">AM11</option>
			<option Value="AM12">AM12</option>			
			



		</select> </td>
		
		</tr>
		<script>
		//Get select object
		var objSelect = document.getElementById("lngPeriod");
		
		//Set selected
		setSelectedValue(objSelect, "<%=lngPeriod%>");
		
		function setSelectedValue(selectObj, valueToSet) {
			for (var i = 0; i < selectObj.options.length; i++) {
				if (selectObj.options[i].text== valueToSet) {
					selectObj.options[i].selected = true;
					return;
				}
			}
		}
	
		</script>
		

	
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
		"<th style=""height:20px""><B><A Target=""_self"" HREF=""FormulaCalculatedFields.asp?Sort=FormulaName&Ordered=" & strOrder & """>Formula Name"
		If strSort = "FormulaName" Then Response.write "<img height=20 =width=20 src=../images/" & strOrder & ".gif Align=absmiddle>"

	response.write"</A></B></th>" & _
		"<th style=""height:20px""><B><A Target=""_self"" HREF=""FormulaCalculatedFields.asp?Sort=CalculatedFieldName&Ordered=" & strOrder & """>Calculated Field Name"
		If strSort = "CalculatedFieldName" Then Response.write "<img height=20 =width=20 src=../images/" & strOrder & ".gif Align=absmiddle>"
	
	response.write"</A></B></th>"
		response.write"</A></B></th>" & _
		"<th><B><A Target=""_self"" HREF=""FormulaCalculatedFields.asp?Sort=Period&Ordered=" & strOrder & """>Period"
		If strSort = "Period" Then Response.write "<img height=20 =width=20 src=../images/" & strOrder & ".gif Align=absmiddle>"
	
	
	
	
	response.write"</A></B></th>" & _
		"<th>Updated By</th>" & _
		"<th>Date Updated</th></tr>"


		
	
    objRS.Open "SELECT * FROM tblFormulaCalculatedFields WHERE BudgetID = " & clng (Session("BudgetID"))  &  " AND VersionID = " & clng (Session("VersionID")) & " Order By " & strSort & " " & strOrder

	Do until objRS.eof	
		
   	    Response.Write "<TR><TD><A Target=""_self"" HREF=""FormulaCalculatedFields.asp?lngFormulaName=" & objRS("FormulaName") & "&lngCalculatedFieldName=" & objRS("CalculatedFieldName") & "&lngPeriod=" & objRS("Period") & "&Sort=" & strSort & """>&nbsp;" & objRS("FormulaName") & "</TD><TD>&nbsp;" & objRS("CalculatedFieldName") & "</TD>" & "</TD><TD>&nbsp;" & objRS("Period")  & "</TD><TD style=""text-align:center"">&nbsp;" & objRS("UpdatedBy") & "</TD><TD style=""text-align:center"">" & objRS("DateUpdated") & "</TD></TR>"
       
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

		objRS.Open "SELECT * FROM tblFormulaCalculatedFields WHERE BudgetID = " & Session("BudgetID")  & " And  VersionID = " & Session("VersionID")  & " And FormulaName='" & lngFormulaName  &  "' And CalculatedFieldName='" & lngCalculatedFieldName & "' And Period='" & lngPeriod & "'", objCon

			If Not objRS.EOF Then
			
				lngFormulaName=objRS("FormulaName")
				lngCalculatedFieldName=objRS("CalculatedFieldName")
				lngPeriod=objRS("Period")
				
				
			Else
			 lngFormulaName  = ""
				
			End If

		objRS.Close

End Sub

Sub SaveDetails()
	
	  objRS.Open "SELECT * FROM tblFormulaCalculatedFields WHERE BudgetID = " & Session("BudgetID")  & " And  VersionID = " & Session("VersionID")  & " And FormulaName='" & lngFormulaName  &  "' And CalculatedFieldName='" & lngCalculatedFieldName & "' And Period='" & lngPeriod & "'",objCon
	  If objRS.Eof=False Then
	  
		
		  objCon.Execute("UPDATE tblFormulaCalculatedFields SET UpdatedBy = " & clng(Session("UserID"))  &", DateUpdated = GetDate() WHERE  BudgetID = " & Session("BudgetID")  & " And  VersionID = " & Session("VersionID")  & " And FormulaName='" & lngFormulaName  &  "' And CalculatedFieldName='" & lngCalculatedFieldName & "' And Period='" & lngPeriod & "'") 
	Else


		 objCon.Execute("Insert into  tblFormulaCalculatedFields (BudgetID, VersionID, FormulaName, CalculatedFieldName, Period, UpdatedBy, DateUpdated) values (" & Session("BudgetID")  & "," & Session("VersionID") & ",'" & request.form("lngFormulaName")  & "','" & request.form("lngCalculatedFieldName") & "','" &  Request.Form("lngPeriod")  & "'," &   clng(Session("UserID")) & ",GetDate())")
	End If  
	objRS.close
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
