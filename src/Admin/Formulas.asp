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

Dim lngFormulaName
Dim lngVersionID
Dim lngFormulaTitle
Dim lngFormulaDesc
Dim lngFormula 
Dim lngFormulaTypeID
Dim lngActive
Dim lngFValActive

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
Dim strEnd

Dim strFormulaName1

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

    'Execute save 	
	If Request.QueryString("Action") = "Delete" Then
        
		DeleteRecord(Request.QueryString("FormulaName"))
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
	
	//frm.submit();
        var varSubmit = true
        var varAlert =""       

	if (isWhitespace(frm.lngFormulaName.value))
	    {
		    varAlert += "Please select a Formula Name. \n \n";
		    document.getElementById('lngFormulaName').style.backgroundColor="ff8080";
		    varSubmit = false;
	    }	
	    else document.getElementById('lngFormulaName').style.backgroundColor="ffffff";

	    //if(frm.lngCalculatedFieldType.value == "" )
	    //{
	//	    varAlert += "Please select a Field Type. \n \n";
	//	    document.getElementById('lngCalculatedFieldType').style.backgroundColor="ff8080";
	//	    varSubmit = false;
	  //  }	
	   // else document.getElementById('lngCalculatedFieldType').style.backgroundColor="ffffff";		   		  
	   	
	  if(varSubmit == true)
	  {

	        frm.submit();

	  }
	  else
	  {
	    window.alert ("" + varAlert);	
   
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

function DeleteData() {
    if (isWhitespace(frm.lngFormulaName.value)) {
        alert('Please select a record to DELETE!');
    } else {
        if (window.confirm('Would you like to DELETE the selected record?') == true) {

            self.location = "Formulas.asp?Action=Delete&FormulaName=" + frm.lngFormulaName.value;
        }

    }
}

//-->
</script>
</head>
<body >
<h3>Formulas Administration Screen
<%
	Response.write "for the currently selected Budget : <FONT color=Red>" &  Session("BudgetName") & "</FONT> and Version : <FONT color=Red>" & Session("VersionName") & "</FONT>"
%>
</h3>
<form action="Formulas.asp?Action=Save" method="POST" id="frm" name="frm">

<table width="100%" align="left" border="1" cellspacing="1" cellpadding="1">

		<tr>
		<th style="text-align:left; height:20px; width:10%;">&nbsp;Formula Name:</th>
		<td align="left" style="width:40%;">&nbsp;<input style="text-align:left; height:20px"  name="lngFormulaName" id="lngFormulaName" type="text"  value="<%=lngFormulaName%>"></td>
		<td colspan="2" width="50%"></td>
		</tr>
		<tr>
		<th style="text-align:left; height:20px">&nbsp;Formula Title:</th>
		<td align="left">
		
		&nbsp;<input name="lngFormulaTitle" style="text-align:left; height:20px"  type="text" value="<%=lngFormulaTitle%>">
		</td><td colspan="2"></td>
		</tr>
		<tr>
		<th style="text-align:left; height:20px">&nbsp;Formula Desc:</th>
		<td align="left">
		<textarea name="lngFormulaDesc" cols="50" rows="5"><%=lngFormulaDesc%></textarea>
		</td><td colspan="2"></td>
		</tr>
		<tr>
		<th style="text-align:left; height:20px">&nbsp;Formula:</th>
		<td align="left">
		&nbsp;<input name="lngFormula" style="text-align:left; height:20px"  type="text" value="<%=lngFormula%>">
		</td><td colspan="2"></td>
		</tr>

		<tr>
		<th style="text-align:left; height:20px; width:10%;">&nbsp;Formula Type ID </th>
		<td style="text-align:left; height:20px; width:40%;">
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
	    </td><td colspan="2"></td>
	  	</tr>
		<tr><th style="text-align:left; height:20px">&nbsp;Active</th>				
		<td align="left" >&nbsp;
		<select Style="Width:95%" id="lngActive" name="lngActive">
		<%If lngActive="Y" Or lngActive="" Then%>
			<option Value="Y" Selected>Yes</option>
			<option Value="N">No</option>
		<%Else%>
			<option Value="Y">Yes</option>
			<option Value="N"  Selected>No</option>
		<%End If%>
		</select> </td>
		<th style="text-align:left; height:20px">&nbsp;FVal Active</th>				
		<td align="left" >&nbsp;
		<select Style="Width:95%" id="lngFValActive" name="lngFValActive">
		<%If lngFValActive="Y" Or lngFValActive="" Then%>
			<option Value="Y" Selected>Yes</option>
			<option Value="N">No</option>
		<%Else%>
			<option Value="Y">Yes</option>
			<option Value="N"  Selected>No</option>
		<%End If%>
		</select> </td>
		</tr>

		<%
		'Write all of the BM months
		For z = 1 to 24
		
			'Write a new row every second z (two BMs)
			If z Mod 2 = 0 Then
				'Response.write "</tr>"
				strEnd = "</tr>"
			Else
				Response.Write "<tr>"
				strEnd = ""
			End If
			Response.Write "<th style=""text-align:left;"" width=""10%"">&nbsp;BM " & z  & "</th>" & _
					"<td align=""left"" width=""40%""><input style=""text-align:left; width:100%;"" type=""text"" name=""lngBM" & z & """ value=""" & lngBM(z) & """ placeholder=""BM " & z & """  />" & strEnd 
			
			If z = 12 Then Response.write "<tr><td style=""height:0px"" colspan=""4"" ></td></tr>"
				
		Next

		Response.write "<tr><td style=""height:0px"" colspan=""4"" ></td></tr>"

		'Write all of the Out Years
		For z = 1 To 5
		
			'Write a new row every second z (two OYs)
			If z Mod 2 = 0 Then
				'Response.write "</tr>"
				strEnd = "</tr>"
			Else
				Response.Write "<tr>"
				If z = 5 then
					strEnd = "<td colspan=""2""></td>"
				Else
					strEnd = ""
				End If
			End If
			Response.Write "<th style=""text-align:left;"" width=""10%"">&nbsp;OY " & z  & "</th>" & _
					"<td align=""left"" width=""40%""><input style=""text-align:left; width:100%;"" type=""text"" name=""lngOY" & z & """ value=""" & lngOY(z) & """ placeholder=""OY " & z & """  />" & strEnd 

		Next
		%>
		

	
	<tr>
		<td style="height:20px" colspan="4" align="left">&nbsp;</td>
	</tr>
	
</table>
<br />
<br />
<table WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
	<tr>
	    <td class='locked' Width="100px" style="border-right:0px"><button type="button" tabindex="8" onClick="self.location='AdminMenu.asp';"><img src="../images/door.png" alt="" /> Close </button></td>    
        <td class='locked' Width="100px" style="border-right:0px"><button type="button" tabindex="9" onClick="SaveData()";><img src="../images/tick.png" alt="" /> Save </button></td>  
		<td class='locked' Width="100px"><button type="button" tabindex="23" onClick="self.location='Formulas.asp?lngFormulaName=';"><img src="../images/page_white_stack.png" alt="" /> New&nbsp;&nbsp; </button></td>
		<td class='locked' Width="100px"><button type="button" tabindex="19" onclick="DeleteData()";><img src="../images/cross.png" alt="" /> Delete </button></td>
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
	
		If IsNull(objRS("FormulaName")) or objRS("FormulaName") = "" Then
			strFormulaName1 = ""
		Else
			'Remove the URL reserved characters for the HREF string
			strFormulaName1 = objRS("FormulaName")
			'strFormulaName1 = Replace(strFormulaName1," ","%20")
			strFormulaName1 = Replace(strFormulaName1,"%","%25")
			strFormulaName1 = Replace(strFormulaName1,"+","%2B")
		End If

   	    Response.Write "<TR><TD><A Target=""_self"" HREF=""Formulas.asp?lngFormulaName=" & strFormulaName1 & "&Sort=" & strSort & """>&nbsp;" & objRS("FormulaName") & "</TD><TD>&nbsp;" & objRS("FormulaTitle") & "</TD>" & "</TD><TD>&nbsp;" & objRS("FormulaDesc") & "</TD><TD style=""text-align:center"">" & objRS("Formula")  & "</TD><TD style=""text-align:center"">" & objRS("Active")  & "</TD><TD style=""text-align:center"">&nbsp;" & objRS("UpdatedBy") & "</TD><TD style=""text-align:center"">" & objRS("DateUpdated") & "</TD></TR>"
       
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

                 lngFValActive=objRS("FVal")
				
				
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
	 		BMUpdate=BMUpdate & " , BM"&z&"Formula" & "=''"
			BMInsert2=BMInsert2 & ",''" 
		Else
			BMUpdate=BMUpdate & " , BM"&z&"Formula" & "='" & request.form("lngBM"&z) & "'"
			BMInsert2=BMInsert2 & ",'" & request.form("lngBM"&z) & "'"
		End If

		'If request.form("lngBM"&z)="" Then 
	 	'	BMUpdate=BMUpdate & " , BM"&z&"Formula" & "=0"
		'	BMInsert2=BMInsert2 & ",0" 
		'Else
		'	BMUpdate=BMUpdate & " , BM"&z&"Formula" & "=" & request.form("lngBM"&z)
		'	BMInsert2=BMInsert2 & "," & request.form("lngBM"&z)
		'End If

         'objCon.Execute "spExtractCalculatedField " & Session("BudgetID") & "," & Session("VersionID") & ",'" & request.form("lngFormulaName") & "','" &  Request.Form("lngFormula") & "'," & "BM" & z & "
         'If Len((request.form("lngBM"&z))) > 0 Then
            'Response.Write "spExtractCalculatedField " & Session("BudgetID") & "," & Session("VersionID") & ",'" & request.form("lngFormulaName") & "','" &  request.form("lngBM"&z) & "','" & "BM" & z & "'," & Session("UserID") & ""
	        objCon.Execute "spExtractCalculatedField " & Session("BudgetID") & "," & Session("VersionID") & ",'" & request.form("lngFormulaName") & "','" &  request.form("lngBM"&z) & "','" & "BM" & z & "'," & Session("UserID") & ""
         'End If
     Next

	OYUpdate=""
	OYInsert1=""
	OYInsert2=""
	 For z=1 to 5
		 OYInsert1=OYInsert1 & ",OY"&z&"Formula"
	 	'If request.form("lngOY"&z)="" Then 
	 	'	OYUpdate=OYUpdate & " , OY"&z&"Formula" & "=0"
		'	OYInsert2=OYInsert2 & ",0" 
		'Else
		'	OYUpdate=OYUpdate & " , OY"&z&"Formula" & "=" & request.form("lngOY"&z)
		'	OYInsert2=OYInsert2 & "," & request.form("lngOY"&z)
		'End If

		If request.form("lngOY"&z)="" Then 
	 		OYUpdate=OYUpdate & " , OY"&z&"Formula" & "=''"
			OYInsert2=OYInsert2 & ",''" 
		Else
			OYUpdate=OYUpdate & " , OY"&z&"Formula" & "='" & request.form("lngOY"&z) & "'"
			OYInsert2=OYInsert2 & ",'" & request.form("lngOY"&z) & "'"
		End If

        ' If Len((request.form("lngOY"&z))) > 0 Then
            'Response.Write "spExtractCalculatedField " & Session("BudgetID") & "," & Session("VersionID") & ",'" & request.form("lngFormulaName") & "','" &  request.form("lngBM"&z) & "','" & "BM" & z & "'," & Session("UserID") & ""
	        objCon.Execute "spExtractCalculatedField " & Session("BudgetID") & "," & Session("VersionID") & ",'" & request.form("lngFormulaName") & "','" &  request.form("lngOY"&z) & "','" & "OY" & z & "'," & Session("UserID") & ""
        ' End If

	 Next
	  objRS.Open "SELECT * FROM tblFormulas WHERE BudgetID = " & Session("BudgetID")  & " And FormulaName='" & request.form("lngFormulaName") &"'" &  "",objCon
	  If objRS.Eof=False Then
	  
		
		  objCon.Execute("UPDATE tblFormulas SET FormulaTitle='" & Request.Form("lngFormulaTitle") & "', FormulaDesc  = '" & Request.Form("lngFormulaDesc") & "', Formula='" & Request.Form("lngFormula") & "', FormulaTypeID=" & request.form("lngFormulaTypeID") & ", Active='" & request.form("lngActive") & "', FVal='" & request.form("lngFValActive") & "', UpdatedBy = " & clng(Session("UserID")) &  BMUpdate &  OYUpdate &", DateUpdated = GetDate() WHERE  BudgetID = " & clng(Session("BudgetID"))  &  " And FormulaName='" &   Request.Form("lngFormulaName") & "'") 
	Else
		'response.write "Insert into  tblFormulas (BudgetID, FormulaName, FormulaTitle, FormulaDesc, Formula, FormulaTypeID, Active " & BMInsert1 & OYInsert1 & ", UpdatedBy,DateUpdated) values (" & Session("BudgetID")  & ",'" & request.form("lngFormulaName")  & "','" & request.form("lngFormulaTitle") & "','" &  Request.Form("lngFormulaDesc")   & "','" &  Request.Form("lngFormula") &"'," &  Request.Form("lngFormulaTypeID") &",'" & Request.form("lngActive") & "'" & BMInsert2 & OYInsert2 & "," &   clng(Session("UserID")) & ",GetDate())"
		

		 
        objCon.Execute("Insert into  tblFormulas (BudgetID, FormulaName, FormulaTitle, FormulaDesc, Formula, FormulaTypeID, Active " & BMInsert1 & OYInsert1 & ", UpdatedBy,DateUpdated,FVal) values (" & Session("BudgetID")  & ",'" & request.form("lngFormulaName")  & "','" & request.form("lngFormulaTitle") & "','" &  Request.Form("lngFormulaDesc")   & "','" &  Request.Form("lngFormula") &"'," &  Request.Form("lngFormulaTypeID") &",'" & Request.form("lngActive") & "'" & BMInsert2 & OYInsert2 & "," &   clng(Session("UserID")) & ",GetDate(),'" & request.form("lngFValActive") & "')")
	     'Response.Write "Insert into  tblFormulas (BudgetID, FormulaName, FormulaTitle, FormulaDesc, Formula, FormulaTypeID, Active " & BMInsert1 & OYInsert1 & ", UpdatedBy,DateUpdated,FVal) values (" & Session("BudgetID")  & ",'" & request.form("lngFormulaName")  & "','" & request.form("lngFormulaTitle") & "','" &  Request.Form("lngFormulaDesc")   & "','" &  Request.Form("lngFormula") &"'," &  Request.Form("lngFormulaTypeID") &",'" & Request.form("lngActive") & "'" & BMInsert2 & OYInsert2 & "," &   clng(Session("UserID")) & ",GetDate(),'" & request.form("lngFValActive") & "'"
    End If  
	objRS.close

   
			'Return the result of the Save Function.
     		strMessage = "<B>Record Saved.</B>"
            strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"
     		
     
	 
	

	 
End Sub

Sub DeleteRecord(FormulaName)

  objCon.Execute("DELETE tblFormulas WHERE BudgetID = " & Session("BudgetID") & " AND FormulaName = '" & FormulaName & "'") 
  objCon.Execute("DELETE tblFormulaCalculatedFields WHERE BudgetID = " & Session("BudgetID") & " AND FormulaName = '" & FormulaName & "'") 
 ' objCon.Execute("UPDATE tblCostCentreStatus SET StatusID = " & clng(StatusID) & ", UpdatedBy = " & clng(Session("UserID")) & ", DateUpdated = GetDate() WHERE  BudgetID = " & clng(Session("BudgetID")) & " and VersionID = " & clng(Session("VersionID")) & "") 
  strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"
  strMessage = "<B>FORMULA HAS BEEN DELETED.</B>"

End Sub	


Set objRS = Nothing
Set objCon = Nothing


%>
