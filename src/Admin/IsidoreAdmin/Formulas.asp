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
Dim lngVersionID
Dim lngFormulaTitle
Dim lngFormulaDesc
Dim lngFormula 
Dim lngFormulaTypeID
Dim lngActive

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

		
	  ' Session("CalculatedFieldName") = lngCalculatedFieldName
	   'If lngCalculatedFieldName <> "" Then strCalculatedField = " AND CalculatedFieldName='" & lngCalculatedFieldName & "'"

	 
	Else 

		lngFormulaName="Phase by Quarter"
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
    self.location="Formulas.asp?lngFormulaName=" + frm.lngFormulaName.value	
}

function SaveData2(){
	var varSubmit = true
	if(document.frm.StatusID.value==0){
		alert("A Status must be selected!");
		varSubmit = false;
	}
	if(varSubmit == true){
	if ( confirm("Would you like to UPDATE all Allocation Method to Status " + document.frm.StatusID.options[document.frm.StatusID.selectedIndex].text + " ?"))
		self.location="Formulas.asp?Action=SaveAll&StatusID=" + document.frm.StatusID.value;
	}else{
		//alert("Status NOT Updated!");
	}

}

//-->
</script>
</head>
<body onload=padding();>
<h3>Formulas Administration Screen
<%
	Response.write "for the currently selected Budget : <FONT color=Red>" &  Session("BudgetName") & "</FONT> and Version : <FONT color=Red>" & Session("VersionName") & "</FONT>"
%>
</h3>
<form action="Formulas.asp?Action=Save" method="POST" id="frm" name="frm">

<table width="50%" align="left" border="1" cellspacing="1" cellpadding="1">

		<tr>
		<th style="text-align:left; height:20px">&nbsp;Formula Name:</th>
		<td align="left">&nbsp;<input style="text-align:left; height:20px"  name="lngFormulaName" type="text"  value="<%=lngFormulaName%>"></td>
		</tr>

		<tr>
		<th style="text-align:left; height:20px">&nbsp;Formula Title:</th>
		<td align="left">
		
		&nbsp;<input name="lngFormulaTitle" style="text-align:left; height:20px"  type="text" value="<%=lngFormulaTitle%>">
		</td>
		</tr>

		<tr>
		<th style="text-align:left; height:20px">&nbsp;Formula Desc:</th>
		<td align="left">
		<textarea name="lngFormulaDesc" cols="50" rows="5"><%=lngFormulaDesc%></textarea>
		</td>
		</tr>
		<tr>
		<th style="text-align:left; height:20px">&nbsp;Formula:</th>
		<td align="left">
		&nbsp;<input name="lngFormula" style="text-align:left; height:20px"  type="text" value="<%=lngFormula%>">
		</td>
		</tr>
				
	<tr>
	<th style="text-align:left; height:20px; width:40%;">&nbsp;Formula Type ID </th>
		<td style="text-align:left; height:20px; width:60%;">
		    <select Style="Width:100%" tabindex="20" id="lngFormulaTypeID" name="lngFormulaTypeID"><OPTION Value=0>Please Select..</OPTION>
	        <%		
		    objRS.Open "SELECT FormulaTypeID,FormulaTypeName FROM tblFormulaTypes WHERE Active = 'Y'",objCon

		    Do until objRS.EOF
			    If objRS("FormulaTypeID") = lngFormulaTypeID Then
				    strSelected = " SELECTED "
			    Else
				    strSelected = ""
			    End if
				    Response.Write "<option Value=""" & objRS("FormulaTypeID") & """" & strSelected & ">" & objRS("FormulaTypeName")  & "</OPTION>"
			    objRS.Movenext
		    Loop
    		
		    objRS.Close

	        %></select>
	    </td>
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


		

		<tr><th style="text-align:left; height:20px">&nbsp;BM 1 to BM 24 </th>				
		<td align="left" >
		<%For z=1 to 24%>
		<input style="text-align:left; height:20px" type="number" step="0.00" name="lngBM<%=z%>"  value="<%=lngBM(z)%>" placeholder="BM<%=z%>"  />
		<%Next%>
		</td>
	
		</tr>
		<tr><th style="text-align:left; height:20px">&nbsp;OY 1 to OY 5</th>				
		<td align="left" >
		<%For z=1 to 5%>
		<input style="text-align:left; height:20px" type="number" step="0.00" name="lngOY<%=z%>" value="<%=lngOY(z)%>"  placeholder="OY<%=z%>"  />
		<%Next%>
		</td>
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
		"<th style=""height:20px""><B><A Target=""_self"" HREF=""Formulas.asp?Sort=FormulaName&Ordered=" & strOrder & """>Formula Name"
		If strSort = "FormulaName" Then Response.write "<img height=20 =width=20 src=../images/" & strOrder & ".gif Align=absmiddle>"

	response.write"</A></B></th>" & _
		"<th style=""height:20px""><B><A Target=""_self"" HREF=""Formulas.asp?Sort=FormulaTitle&Ordered=" & strOrder & """>Formula Title"
		If strSort = "FormulaTitle" Then Response.write "<img height=20 =width=20 src=../images/" & strOrder & ".gif Align=absmiddle>"
	
	response.write"</A></B></th>"
		response.write"</A></B></th>" & _
		"<th><B><A Target=""_self"" HREF=""Formulas.asp?Sort=FormulaDesc&Ordered=" & strOrder & """>Formula Desc"
		If strSort = "FormulaDesc" Then Response.write "<img height=20 =width=20 src=../images/" & strOrder & ".gif Align=absmiddle>"
	
	
	response.write"</A></B></th>"
		response.write"</A></B></th>" & _
		"<th><B><A Target=""_self"" HREF=""Formulas.asp?Sort=Formula&Ordered=" & strOrder & """>Formula"
		If strSort = "Formula" Then Response.write "<img height=20 =width=20 src=../images/" & strOrder & ".gif Align=absmiddle>"
	
	response.write"</A></B></th>"
		response.write"</A></B></th>" & _
		"<th><B><A Target=""_self"" HREF=""Formulas.asp?Sort=Active&Ordered=" & strOrder & """>Active"
		If strSort = "Active" Then Response.write "<img height=20 =width=20 src=../images/" & strOrder & ".gif Align=absmiddle>"
	
	response.write"</A></B></th>" & _
		"<th>Updated By</th>" & _
		"<th>Date Updated</th></tr>"


    objRS.Open "SELECT * FROM tblFormulas WHERE BudgetID = " & clng (Session("BudgetID"))  &  " Order By " & strSort & " " & strOrder

	Do until objRS.eof	
		
   	    Response.Write "<TR><TD><A Target=""_self"" HREF=""Formulas.asp?lngFormulaName=" & objRS("FormulaName") & "&Sort=" & strSort & """>&nbsp;" & objRS("FormulaName") & "</TD><TD>&nbsp;" & objRS("FormulaTitle") & "</TD>" & "</TD><TD>&nbsp;" & objRS("FormulaDesc") & "</TD><TD style=""text-align:center"">" & objRS("Formula")  & "</TD><TD style=""text-align:center"">" & objRS("Active")  & "</TD><TD style=""text-align:center"">&nbsp;" & objRS("UpdatedBy") & "</TD><TD style=""text-align:center"">" & objRS("DateUpdated") & "</TD></TR>"
       
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

		objRS.Open "SELECT * FROM tblFormulas WHERE BudgetID = " & Session("BudgetID")  & " And FormulaName='" & lngFormulaName  &  "'",objCon

			If Not objRS.EOF Then
			
				lngFormulaName=objRS("FormulaName")
				lngFormulaTitle=objRS("FormulaTitle")
				lngFormulaDesc=objRS("FormulaDesc")
				lngFormula =objRS("Formula")
				lngFormulaTypeID=objRS("FormulaTypeID")

				lngActive=objRS("Active")
				 For z=1 to 24
					lngBM(z)=objRS("BM"&z&"Formula")
				 Next
				 
				 For z=1 to 5
					lngOY(z)=objRS("OY"&z&"Formula")
				 Next
				
			Else
			 lngFormulaName  = ""
				
			End If

		objRS.Close

End Sub

Sub SaveDetails()
	BMUpdate=""
	BMInsert1=""
	BMInsert2=""
	 For z=1 to 24
	 	BMInsert1=BMInsert1 & ",BM"&z&"Formula"
	 	If request.form("lngBM"&z)="" Then 
	 		BMUpdate=BMUpdate & " , BM"&z&"Formula" & "=0"
			BMInsert2=BMInsert2 & ",0" 
		Else
			BMUpdate=BMUpdate & " , BM"&z&"Formula" & "=" & request.form("lngBM"&z)
			BMInsert2=BMInsert2 & "," & request.form("lngBM"&z)
		End If
	 Next

	OYUpdate=""
	OYInsert1=""
	OYInsert2=""
	 For z=1 to 5
		 OYInsert1=OYInsert1 & ",OY"&z&"Formula"
	 	If request.form("lngOY"&z)="" Then 
	 		OYUpdate=OYUpdate & " , OY"&z&"Formula" & "=0"
			OYInsert2=OYInsert2 & ",0" 
		Else
			OYUpdate=OYUpdate & " , OY"&z&"Formula" & "=" & request.form("lngOY"&z)
			OYInsert2=OYInsert2 & "," & request.form("lngOY"&z)
		End If
	 Next
	  objRS.Open "SELECT * FROM tblFormulas WHERE BudgetID = " & Session("BudgetID")  & " And FormulaName='" & request.form("lngFormulaName") &"'" &  "",objCon
	  If objRS.Eof=False Then
	  
		
		  objCon.Execute("UPDATE tblFormulas SET FormulaTitle='" & Request.Form("lngFormulaTitle") & "', FormulaDesc  = '" & Request.Form("lngFormulaDesc") & "', Formula='" & Request.Form("lngFormula") & "', FormulaTypeID=" & request.form("lngFormulaTypeID") & ", Active='" & request.form("lngActive") & "', UpdatedBy = " & clng(Session("UserID")) &  BMUpdate &  OYUpdate &", DateUpdated = GetDate() WHERE  BudgetID = " & clng(Session("BudgetID"))  &  " And FormulaName='" &   Request.Form("lngFormulaName") & "'") 
	Else
		'response.write "Insert into  tblFormulas (BudgetID, FormulaName, FormulaTitle, FormulaDesc, Formula, FormulaTypeID, Active " & BMInsert1 & OYInsert1 & ", UpdatedBy,DateUpdated) values (" & Session("BudgetID")  & ",'" & request.form("lngFormulaName")  & "','" & request.form("lngFormulaTitle") & "','" &  Request.Form("lngFormulaDesc")   & "','" &  Request.Form("lngFormula") &"'," &  Request.Form("lngFormulaTypeID") &",'" & Request.form("lngActive") & "'" & BMInsert2 & OYInsert2 & "," &   clng(Session("UserID")) & ",GetDate())"
		

		 objCon.Execute("Insert into  tblFormulas (BudgetID, FormulaName, FormulaTitle, FormulaDesc, Formula, FormulaTypeID, Active " & BMInsert1 & OYInsert1 & ", UpdatedBy,DateUpdated) values (" & Session("BudgetID")  & ",'" & request.form("lngFormulaName")  & "','" & request.form("lngFormulaTitle") & "','" &  Request.Form("lngFormulaDesc")   & "','" &  Request.Form("lngFormula") &"'," &  Request.Form("lngFormulaTypeID") &",'" & Request.form("lngActive") & "'" & BMInsert2 & OYInsert2 & "," &   clng(Session("UserID")) & ",GetDate())")
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
