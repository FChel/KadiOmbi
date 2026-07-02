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

Dim lngCalculatedFieldName
Dim lngVersionID
Dim lngCostCentreID
Dim strSort
Dim strOrder
Dim strCalculatedField
Dim lngActive
Dim lngBM(12)
Dim lngOY(5)
Dim lngComments
Dim z
Dim BMUpdate
Dim OYUpdate
Dim BMInsert1
DIM	BMInsert2
DIM	OYInsert1
DIM	OYInsert2

'3. Capture Querystring variables

    If Not IsEmpty(Request("lngCalculatedFieldName")) Then
	   	lngCalculatedFieldName = Request("lngCalculatedFieldName")
		lngCostCentreID= Request("lngCostCentreID")
		
	  ' Session("CalculatedFieldName") = lngCalculatedFieldName
	   'If lngCalculatedFieldName <> "" Then strCalculatedField = " AND CalculatedFieldName='" & lngCalculatedFieldName & "'"

	 
	Else 

		lngCalculatedFieldName="CPI"
		lngCostCentreID=Session("CostCentreID")
    End If
    
    If Not IsEmpty(Request.QueryString("Sort")) Then
	   strSort = Request.QueryString("Sort")
    Else
	   strSort = "CalculatedFieldName"
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
    self.location="CalculatedFieldValues.asp?lngCalculatedFieldName=" + frm.CalculatedFieldName.value	
}

function SaveData2(){
	var varSubmit = true
	if(document.frm.StatusID.value==0){
		alert("A Status must be selected!");
		varSubmit = false;
	}
	if(varSubmit == true){
	if ( confirm("Would you like to UPDATE all Allocation Method to Status " + document.frm.StatusID.options[document.frm.StatusID.selectedIndex].text + " ?"))
		self.location="CalculatedFieldValues.asp?Action=SaveAll&StatusID=" + document.frm.StatusID.value;
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
<form action="CalculatedFieldValues.asp?Action=Save" method="POST" id="frm" name="frm">

<table width="50%" align="left" border="1" cellspacing="1" cellpadding="1">

	<th style="text-align:left; height:20px; width:40%;">&nbsp;Calculated Field Name</th>
		<td style="text-align:left; height:20px; width:60%;">
		    <select Style="Width:100%" tabindex="20" id="lngCalculatedFieldName" name="lngCalculatedFieldName"><OPTION Value=0>Please Select..</OPTION>
	        <%		
		    objRS.Open "SELECT CalculatedFieldName FROM tblCalculatedFields WHERE BudgetID = " & Session("BudgetID") & " AND Active = 'Y'",objCon

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

		<th style="text-align:left; height:20px; width:40%;">&nbsp;Cost Centre </th>
		<td style="text-align:left; height:20px; width:60%;">
		    <select Style="Width:100%" tabindex="20" id="lngCostCentreID" name="lngCostCentreID"><OPTION Value=0>Please Select..</OPTION>
	        <%		
		    objRS.Open "SELECT CostCentreID FROM tblCostCentres WHERE BudgetID = " & Session("BudgetID") & " AND Active = 'Y'",objCon

		    Do until objRS.EOF
			    If objRS("CostCentreID") = lngCostCentreID Then
				    strSelected = " SELECTED "
			    Else
				    strSelected = ""
			    End if
				    Response.Write "<option Value=""" & objRS("CostCentreID") & """" & strSelected & ">" & objRS("CostCentreID")  & "</OPTION>"
			    objRS.Movenext
		    Loop
    		
		    objRS.Close

	        %></select>
	    </td>

		<tr><th style="text-align:left; height:20px">&nbsp;BM 1 to BM 12 </th>				
		<td align="left" >
		<%For z=1 to 12%>
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
		<th style="text-align:left; height:20px">&nbsp;Comments</th>
		<td align="left" >
		<textarea name="lngComments" cols="50" rows="5"><%=lngComments%></textarea>

		
		</td>
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
		"<th style=""height:20px""><B><A Target=""_self"" HREF=""CalculatedFieldValues.asp?Sort=CalculatedFieldName&Ordered=" & strOrder & """>Calculated Field "
		If strSort = "CalculatedFieldName" Then Response.write "<img height=20 =width=20 src=../images/" & strOrder & ".gif Align=absmiddle>"

	response.write"</A></B></th>" & _
		"<th style=""height:20px""><B><A Target=""_self"" HREF=""CalculatedFieldValues.asp?Sort=CostCentreID&Ordered=" & strOrder & """>Cost Centre ID"
		If strSort = "CostCentreID" Then Response.write "<img height=20 =width=20 src=../images/" & strOrder & ".gif Align=absmiddle>"
	
	response.write"</A></B></th>"
		response.write"</A></B></th>" & _
		"<th><B><A Target=""_self"" HREF=""CalculatedFieldValues.asp?Sort=Comments&Ordered=" & strOrder & """>Comments"
		If strSort = "Comments" Then Response.write "<img height=20 =width=20 src=../images/" & strOrder & ".gif Align=absmiddle>"
	
	
	
	
	
	response.write"</A></B></th>" & _
		"<th>Updated By</th>" & _
		"<th>Date Updated</th></tr>"

    objRS.Open "SELECT * FROM tblCalculatedFieldValues WHERE BudgetID = " & clng (Session("BudgetID"))  & " And VersionID=" & Session("VersionID")  & " Order By " & strSort & " " & strOrder

	Do until objRS.eof	
		
   	    Response.Write "<TR><TD><A Target=""_self"" HREF=""CalculatedFieldValues.asp?lngCalculatedFieldName=" & objRS("CalculatedFieldName") & "&lngCostCentreID="& objRS("CostCentreID") & "&Sort=" & strSort & """>&nbsp;" & objRS("CalculatedFieldName") & "</TD><TD>&nbsp;" & objRS("CostCentreID") & "</TD><TD style=""text-align:center"">" & objRS("Comments")  & "</TD><TD style=""text-align:center"">&nbsp;" & objRS("UpdatedBy") & "</TD><TD style=""text-align:center"">" & objRS("DateUpdated") & "</TD></TR>"
       
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

		objRS.Open "SELECT * FROM tblCalculatedFieldValues WHERE BudgetID = " & Session("BudgetID")  & " And VersionID=" & session("VersionID") & " And CalculatedFieldName='" & lngCalculatedFieldName & "' And CostCentreID="& lngCostCentreID &  "",objCon
		
			If Not objRS.EOF Then
			
			  lngCalculatedFieldName = objRS("CalculatedFieldName")
              lngCostCentreID= objRS("CostCentreID")
              lngComments = objRS("Comments")

				 For z=1 to 12
					lngBM(z)=objRS("BM"&z)
				 Next
				 For z=1 to 5
					lngOY(z)=objRS("OY"&z)
				 Next
				
			Else
			 lngCalculatedFieldName  = ""
				
			End If

		objRS.Close
	

End Sub

Sub SaveDetails()
	BMUpdate=""
	BMInsert1=""
	BMInsert2=""
	 For z=1 to 12
	 	BMInsert1=BMInsert1 & ",BM"&z
	 	If request.form("lngBM"&z)="" Then 
	 		BMUpdate=BMUpdate & " , BM"&z & "=0"
			BMInsert2=BMInsert2 & ",0" 
		Else
			BMUpdate=BMUpdate & " , BM"&z & "=" & request.form("lngBM"&z)
			BMInsert2=BMInsert2 & "," & request.form("lngBM"&z)
		End If
	 Next

	OYUpdate=""
	OYInsert1=""
	OYInsert2=""
	 For z=1 to 5
		 OYInsert1=OYInsert1 & ",OY"&z
	 	If request.form("lngOY"&z)="" Then 
	 		OYUpdate=OYUpdate & " , OY"&z & "=0"
			OYInsert2=OYInsert2 & ",0" 
		Else
			OYUpdate=OYUpdate & " , OY"&z & "=" & request.form("lngOY"&z)
			OYInsert2=OYInsert2 & "," & request.form("lngOY"&z)
		End If
	 Next
	  objRS.Open "SELECT * FROM tblCalculatedFieldValues WHERE BudgetID = " & Session("BudgetID")  & " And VersionID=" & session("VersionID") & " And CalculatedFieldName='" & request.form("lngCalculatedFieldName") & "' And CostCentreID="&  Request.Form("lngCostCentreID") &  "",objCon
	  If objRS.Eof=False Then
	  
		
		  objCon.Execute("UPDATE tblCalculatedFieldValues SET CalculatedFieldName='" & Request.Form("lngCalculatedFieldName") & "', CostCentreID  = '" & Request.Form("lngCostCentreID") & "', UpdatedBy = " & clng(Session("UserID")) &  BMUpdate &  OYUpdate &", Comments='" & Request.Form("lngComments") & "', DateUpdated = GetDate() WHERE  BudgetID = " & clng(Session("BudgetID"))  & " And VersionID=" & session("VersionID") &  " And CalculatedFieldName='" &   Request.Form("lngCalculatedFieldName")  & "' And CostCentreID="&  Request.Form("lngCostCentreID") & "") 
	Else

		  objCon.Execute("Insert into  tblCalculatedFieldValues (BudgetID, VersionID, CalculatedFieldName, CostCentreID, Comments" & BMInsert1 & OYInsert1 & ", UpdatedBy,DateUpdated) values (" & Session("BudgetID")  & "," & session("VersionID")  & ",'" & request.form("lngCalculatedFieldName") & "'," &  Request.Form("lngCostCentreID")   & ",'" &  Request.Form("lngComments") &"'" & BMInsert2 & OYInsert2 & "," &   clng(Session("UserID")) & ",GetDate())")
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
