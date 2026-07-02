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
Dim arrYesNo(3)
Dim arrFundingType(4)
Dim strActive
Dim strSort
Dim strOrder
Dim arrMFFMode(2)
Dim strMFFMode
Dim strLevel2

arrYesNo(1) = "Y"
arrYesNo(2) = "N"
arrYesNo(3) = "X"

arrFundingType(1) = "RR"
arrFundingType(2) = "RD"
arrFundingType(3) = "ER"
arrFundingType(4) = "DR"

arrMFFMode(1) = "M"
arrMFFMode(2) = "R"

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
Dim lngFundingPurposeID
Dim lngBudgetID
Dim strFundingPurposeName
Dim strFundingPurposeDesc
Dim lngAllocationMethod
Dim lngCeilingLevelID
Dim intSortOrder
Dim strCalculatedField
Dim strFundingType
Dim strAllocationMethod
Dim strDefaultFund
Dim dblAlloRate

If IsEmpty(Session("FundingPurposeID")) Then Session("FundingPurposeID") = 0

'3. Capture Querysring variables

    If Not IsEmpty(Request.QueryString("FundingPurposeID")) Then

	   Session("FundingPurposeID") = clng(Request.QueryString("FundingPurposeID"))	  

    End If

	'Execute save 	
	If Request.QueryString("Action") = "Save" Then
		SaveDetails()
	End If

	  'Execute save 	
	If Request.QueryString("Action") = "Delete" Then
        
		DeleteRecord(Request.QueryString("FundingPurposeID"))
	End If

	'Load page details
	LoadDetails()

    If Request.QueryString("Action") = "BuildMTFF" Then
        Call Build_MTFF()
    End If

    If Request.QueryString("Action") = "ResetMTFF" Then
        Call Reset_MTFF()
    End If

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

	    if(frm.FundingPurposeID.value == 0 )
	    {
		    varAlert += "Please enter a Funding Purpose ID. \n \n";
		    document.getElementById('FundingPurposeID').style.backgroundColor="ff8080";
		    varSubmit = false;
	    }	
	    else document.getElementById('FundingPurposeID').style.backgroundColor="ffffff";		   		  
	   	
	  if(varSubmit == true)
	  {
	        frm.submit();
	  }
	  else
	  {
	    window.alert ("" + varAlert);	    
	  }
	  
    }

    function DeleteData() {
        if (isWhitespace(frm.FundingPurposeID.value)) {
            alert('Please select a record to DELETE!');
        } else {
            if (window.confirm('Would you like to DELETE the selected record?') == true) {

                self.location = "FundingPurposes.asp?Action=Delete&FundingPurposeID=" + frm.FundingPurposeID.value;
            }

        }
    }

    function BuildMTFF() {

        if (window.confirm('Would you like to Build Macro Fiscal Framework?') == true) {
            document.getElementById('Progress').style.display = "inline";
            self.location = "FundingPurposes.asp?Action=BuildMTFF"

        }

    }

    function ResetMTFF() {

        if (window.confirm('Would you like to Reset Macro Fiscal Framework?') == true) {
            document.getElementById('Progress').style.display = "inline";
            self.location = "FundingPurposes.asp?Action=ResetMTFF"

        }

    }
//-->
</script>
</head>
<body onload=padding();>
<h3>Funding Purpose Administration Screen</h3>
<form action="FundingPurposes.asp?Action=Save" method="POST" id="frm" name="frm">

<table width="100%" align="left" border="1" cellspacing="1" cellpadding="1">
<tr>
    <th style="text-align:left; height:20px; width:10%;">&nbsp;Funding Purpose ID</th>				
	<td style="text-align:left; height:20px; width:20%;" >&nbsp;<input style="text-align:left; height:20px" type="number"  id="FundingPurposeID" name="FundingPurposeID" value="<%=Session("FundingPurposeID")%>"  /></td>
    <td style="text-align:left; height:20px; width:70%;" rowspan="12"><iframe id="framecontent" name="framecontent" src="../Indexation/MTFF.asp" Style="height:400px; width:600px;" frameborder="0"></iframe></td>
</tr>
<tr>
    <th style="text-align:left; height:20px">&nbsp;Funding Purpose Name</th>				
	<td align="left" >&nbsp;<input style="text-align:left; height:20px" type="text" id="FundingPurposeName" name="FundingPurposeName" value="<%=strFundingPurposeName%>"  /></td>
</tr>
<tr>
    <th style="text-align:left; height:20px">&nbsp;Funding Purpose Desc</th>				
	<td align="left" >&nbsp;<input style="text-align:left; height:20px" type="text" id="FundingPurposeDesc" name="FundingPurposeDesc" value="<%=strFundingPurposeDesc%>"  /></td>
</tr>
<tr>
	<th style="text-align:left; height:20px; width:40%;">&nbsp;Calculated Fields</th>
		<td style="text-align:left; height:20px; width:60%;">
		    <select Style="Width:100%" tabindex="20" id="CalculatedField" name="CalculatedField"><OPTION Value="">Please Select..</OPTION>
	        <%	
		    objRS.Open "SELECT * FROM tblCalculatedFields WHERE BudgetID = " & Session("BudgetID") & " AND CalculatedFieldType = 'MTFF' AND Active = 'Y'",objCon

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
    </tr>
    <tr>
    <th style="text-align:left; height:20px">&nbsp;Sort Order</th>				
	<td align="left" >&nbsp;<input style="text-align:left; height:20px" type="number"  id="SortOrder" name="SortOrder" value="<%=intSortOrder%>" /></td>
</tr>
<tr>        
      <th height="20px" style="text-align:left">&nbsp;Funding Type</th>		
		<td><select Style="Width:40%" tabindex="9" id="FundingType" name="FundingType"><option Value="0">Please Select....</option>
		<%
		For x = 1 to 4
			If strFundingType = arrFundingType(x)Then
				strSelected = " SELECTED " 
			Else
				strSelected = ""	
			End If
			
			Response.Write "<option Value=""" & arrFundingType(x) & """" & strSelected & ">" & arrFundingType(x) & "</OPTION>"
		Next
		%>
		</select> </td>
		
	</tr>
    <tr>
	
		<th style="text-align:left; height:20px">&nbsp;Allocation Method</th>		
		<td><select Style="Width:100%" tabindex="6" id="AllocationMethod" name="AllocationMethod"><option Value="">Please Select....</option>
		<%
		objRS.Open "SELECT * FROM tblAllocationMethods WHERE BudgetID = " & Session("BudgetID") & " AND Active = 'Y'",objCon
		
		Do until objRS.EOF
			If objRS("AllocationMethodName") = strAllocationMethod Then
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
        <tr>   <th Style="Height:20px; text-align:left;">&nbsp;Default Fund</th>
        <td style="background-color:FFFFFF;"><select Style="Width:95%; background-color:FFFFFF;" tabindex="2" id="DefaultFund" name="DefaultFund"><option Value="0">Please Select....</option>
		<%
		objRS.Open "SELECT * FROM tblSegmentValues WITH(NOLOCK) WHERE BudgetID = " & Session("BudgetID") & " AND BusinessAreaID = 1000 AND SegmentNo = 9 AND Active = 'Y'",objCon,0,1
		
		Do until objRS.EOF
			If objRS("SegmentCode") = strDefaultFund Then
				strSelected = " SELECTED "
			Else
				strSelected = ""
			End if
				Response.Write "<option Value=""" & objRS("SegmentCode") & """" & strSelected & ">" & objRS("SegmentCode") & " : " & objRS("SegmentName") & "</OPTION>"
			objRS.Movenext
		Loop
		
		objRS.Close
		%>
		</select> </td>	
        </tr>
 <tr>
    <th style="text-align:left; height:20px">&nbsp;Allocation Rate</th>				
	<td align="left" >&nbsp;<input style="text-align:left; height:20px" type="number"  id="AllocationRate" name="AllocationRate" value="<%=dblAlloRate%>" /></td>
</tr>
    <tr>
         <th height="20px" style="text-align:left">&nbsp;Active</th>		
		<td><select Style="Width:40%" tabindex="9" id="Active" name="Active"><option Value="0">Please Select....</option>
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
    </tr>
      <tr>
         <th height="20px" style="text-align:left">&nbsp;MFF Mode</th>		
		<td><select Style="Width:40%" tabindex="9" id="MFFMode" name="MFFMode"><option Value="0">Please Select....</option>
		<%
		For x = 1 to 2
			If strMFFMode = arrMFFMode(x)Then
				strSelected = " SELECTED " 
			Else
				strSelected = ""	
			End If
			
			Response.Write "<option Value=""" & arrMFFMode(x) & """" & strSelected & ">" & arrMFFMode(x) & "</OPTION>"
		Next
		%>
		</select> </td>
    </tr>
      <tr>
         <th height="20px" style="text-align:left">&nbsp;Level 2</th>		
		<td><select Style="Width:40%" tabindex="9" id="Level2" name="Level2"><option Value="0">Please Select....</option>
		<%
		For x = 1 to 3
			If strLevel2 = arrYesNo(x)Then
				strSelected = " SELECTED " 
			Else
				strSelected = ""	
			End If
			
			Response.Write "<option Value=""" & arrYesNo(x) & """" & strSelected & ">" & arrYesNo(x) & "</OPTION>"
		Next
		%>
		</select> </td>
    </tr>
    <tr>
		<td style="height:20px" colspan="4" align="left">&nbsp;</td>
	</tr>
</table>
<br/>

<table WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
	<tr>
	    <td Width="100px" style="border-right:0px"><button type="button" tabindex="8" onClick="self.location='AdminMenu.asp';"><img src="../images/door.png" alt="" /> Close </button></td> 
        <td Width="100px" style="border-right:0px"><button type="button" tabindex="9" onClick="SaveData()";><img src="../images/tick.png" alt="" /> Save </button></td>  
		<td Width="100px" style="border-right:0px"><button type="button" tabindex="23" onClick="self.location='FundingPurposes.asp?FundingPurposeID=0';"><img src="../images/page_white_stack.png" alt="" /> New&nbsp;&nbsp; </button></td>
		<td Width="100px" style="border-right:0px"><button type="button" tabindex="19" onclick="DeleteData()";><img src="../images/cross.png" alt="" /> Delete </button></td>
        <td Width="100px" style="border-right:0px"><button type="button" tabindex="13" onclick="javascript:BuildMTFF();"><img src="../images/bricks.png" alt="" /> Build MTFF</button></td>
        <td Width="120px" style="border-right:0px"><button type="button" tabindex="13" onclick="javascript:ResetMTFF();"><img src="../images/wrench.png" alt="" /> Reset MTFF</button></td>
        <TD Width="120px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=strMessageIcon %></TD>
        <TD Width="200px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=strMessage %></TD>
	</tr>
    <tr>
    <td Colspan="8"><span id="Progress" style="display:none"><img src=../Images/progress.gif />  &nbsp;&nbsp;&nbsp; <b><FONT Face="Arial">Applying Framework....</FONT></b></span></td>	
</TR>
</table>
<hr />
</form>
<h3>Funding Purposes</h3>
<table WIDTH="100%" BORDER="1" CELLSPACING="1" CELLPADDING="1">
<tr><th>ID</th><th>Name</th><th>Calculated Field</th><th>Allocation Method</th><th>Allocation Rate</th><th>Default Fund</th><th>Type</th><th>MFF Mode</th><th>Level 2</th></tr>
	<tr>
<%
    objRS.Open "SELECT * FROM tblFundingPurposes WHERE BudgetID = " & clng (Session("BudgetID"))  & " Order By SortOrder"

	Do until objRS.eof	
		
   	    Response.Write "<TR><TD><A Target=""_self"" HREF=""FundingPurposes.asp?FundingPurposeID=" & objRS("FundingPurposeID") & """>&nbsp;" & objRS("FundingPurposeID") & "</TD><TD>&nbsp;" & objRS("FundingPurposeName") & "</TD><TD style=""text-align:center"">" & objRS("CalculatedField") & "</TD><TD style=""text-align:center"">" & objRS("AllocationMethod")  & "</TD><TD style=""text-align:center"">&nbsp;" & objRS("AllocationRate") & "</TD><TD style=""text-align:center"">&nbsp;" & objRS("DefaultFundID") & "</TD><TD style=""text-align:center"">&nbsp;" & objRS("FundingType") & "</TD><TD style=""text-align:center"">" & objRS("MFFMode") & "</TD><TD style=""text-align:center"">" & objRS("Level2") & "</TD></TR>"
       
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
	
		objRS.Open "SELECT * FROM tblFundingPurposes WHERE BudgetID = " & Session("BudgetID")  & " And FundingPurposeID = " & Session("FundingPurposeID") & "",objCon
		
			If Not objRS.EOF Then
			
			  Session("FundingPurposeID") = objRS("FundingPurposeID")
			  strFundingPurposeName = objRS("FundingPurposeName")
              strFundingPurposeDesc = objRS("FundingPurposeDesc")
 			  strCalculatedField = objRS("CalculatedField")
			  strActive= objRS("Active")
              intSortOrder = objRS("SortOrder")
              strFundingType = objRS("FundingType")
              strAllocationMethod = objRS("AllocationMethod")
              strDefaultFund = objRS("DefaultFundID")
              dblAlloRate   = objRS("AllocationRate")
              strMFFMode = objRS("MFFMode")
              strLevel2 = objRS("Level2")
				
			End If

		objRS.Close


End Sub

Sub SaveDetails()

    objCon.Execute "spFundingPurposeSave " & Session("BudgetID") & "," & Request.Form("FundingPurposeID") & ",'" & Request.Form("FundingPurposeName") & "','" & Request.Form("FundingPurposeDesc") & "', '" & Request.Form("CalculatedField") & "', '" & Request.Form("SortOrder")  & "', '" & Request.Form("Active") & "','" & Request.Form("FundingType") & "'," & Session("UserID") & ",'" & Request.Form("AllocationMethod") & "','" & Request.Form("DefaultFund") & "'," & Request.Form("AllocationRate") & ",'" & Request.Form("MFFMode") & "','" & Request.Form("Level2") & "'"
	'Response.Write "spFundingPurposeSave " & Session("BudgetID") & "," & Request.Form("FundingPurposeID") & ",'" & Request.Form("FundingPurposeName") & "','" & Request.Form("FundingPurposeDesc") & "', '" & Request.Form("CalculatedField") & "', '" & Request.Form("SortOrder")  & "', '" & Request.Form("Active") & "','" & Request.Form("FundingType") & "'," & Session("UserID") & ""   
	'Return the result of the Save Function.
   	strMessage = "<B>Record Saved.</B>"
    strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"
   
	 
End Sub

Sub DeleteRecord(FundingPurposeID)

  objCon.Execute("DELETE tblFundingPurposes WHERE BudgetID = " & Session("BudgetID") & " AND FundingPurposeID = " & FundingPurposeID & "") 
 
  strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"
  strMessage = "<B>RECORD HAS BEEN DELETED.</B>"

End Sub	

Sub Build_MTFF()

    objCon.Execute "spBuildMTFF " & Session("BudgetID") & "," & Session("VersionID") & ",1000,'N'," & Session("UserID") & ""
    strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"
    strMessage = "<B>MACRO FISCAL FRAMEWORK HAS BEEN BUILT.</B>"

End Sub

Sub Reset_MTFF()

    objCon.Execute "spBuildMTFF " & Session("BudgetID") & "," & Session("VersionID") & ",1000,'Y'," & Session("UserID") & ""
    strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"
    strMessage = "<B>MACRO FISCAL FRAMEWORK HAS BEEN RESET.</B>"

End Sub


Set objRS = Nothing
Set objCon = Nothing


%>
