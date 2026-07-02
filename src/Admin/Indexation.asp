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

Session("CurrentPage") = "Admin/Indexation.asp"

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
Dim lngReportID

'3. Capture Querystring variables

    If Not IsEmpty(Request.Querystring("Level1ID")) Then
	   	Session("Level1ID") = Request.Querystring("Level1ID")
	End If

    If Not IsEmpty(Request.Querystring("GLCode")) Then
	   	Session("GLCode") = Request.Querystring("GLCode")
	End If

    If Not IsEmpty(Request.QueryString("CostCentreID")) Then
	   Session("CostCentreID") = Request.QueryString("CostCentreID")
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
		DeleteRecord()
	End If

    'Execute Reset 	
	If Request.QueryString("Action") = "Reset" Then
		Reset()
	End If

	'Call the Save procedure to Update ALL Business Areas.
	If Request.QueryString("Action") = "SaveAll" Then
		SaveAllRecord(Request.QueryString("StatusID"))
	End If

	'Load page details

    If IsNull(Session("CostCentreID")) or Session("CostCentreID") = "" Then Session("CostCentreID") = 0 End If
    If IsNull(Session("GLCode")) or Session("GLCode") = "" Then Session("GlCode") = 0 End If
    If IsNull(Session("Level1ID")) or Session("Level1ID") = "" Then Session("Level1ID") = 0 End If
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
        alert('DDDD');
 	if(window.confirm('Confirm delete')==true){
	self.location="Indexation.asp?Action=Delete"
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

function CopyTo() {
  
    var mynum = frm.lngBM1.value;

    //mynum = mynum.replace(/,/g, "");
    //mynum = FormatNumber(mynum, 0, 0, 0, 1);
    document.getElementById('lngBM2').value = mynum;
    document.getElementById('lngBM3').value = mynum;
    document.getElementById('lngBM4').value = mynum;
    document.getElementById('lngBM5').value = mynum;
    document.getElementById('lngBM6').value = mynum;
    document.getElementById('lngBM7').value = mynum;
    document.getElementById('lngBM8').value = mynum;
    document.getElementById('lngBM9').value = mynum;
    document.getElementById('lngBM10').value = mynum;
    document.getElementById('lngBM11').value = mynum;
    document.getElementById('lngBM12').value = mynum;
}

function DeleteData() {
    if (isWhitespace(frm.GLCode.value)) {
        alert('Please select a record to DELETE!');
    } else {
        if (window.confirm('Would you like to DELETE the selected record?') == true) {

            self.location = "Indexation.asp?Action=Delete&GLCode=" + frm.GLCode.value + "&Level1ID=" + frm.Level1ID.value + "&CostCentreID=" + frm.CostCentreID.value;
        }

    }
}

function Reset() {
   
        if (window.confirm('Would you like to RESET the indexation for selected Business Area?') == true) {

            self.location = "Indexation.asp?Action=Reset"
        }

    }


//-->
</script>
</head>
<body onload=padding();>
<h3>Indexation Administration Screen
<%
	Response.write "for the currently selected Budget : <FONT color=Red>" &  Session("BudgetName") & "</FONT> and Version : <FONT color=Red>" & Session("VersionName") & "</FONT>"
%>
</h3>
<form action="Indexation.asp?Action=Save" method="POST" id="frm" name="frm">

<table width="50%" align="left" border="1" cellspacing="1" cellpadding="1">
<tr>

		<th style="text-align:left; height:20px; width:20%;">&nbsp;Cost Centre </th>
		<td style="text-align:left; height:20px; width:80%;">
		    <select Style="Width:100%" tabindex="20" id="CostCentreID" name="CostCentreID" onchange="self.location='Indexation.asp?Level1ID=' + frm.Level1ID.value + '&CostCentreID=' + frm.CostCentreID.value"><OPTION Value=0>Apply to All for selected Business Area..</OPTION>
	        <%		
		    objRS.Open "SELECT * FROM qryCostCentresByBusinessArea WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND BusinessAreaID = " & Session("BusinessAreaID") & " AND Active = 'Y'",objCon

		    Do until objRS.EOF
			    If clng(objRS("CostCentreID")) = clng(Session("CostCentreID")) Then
				    strSelected = " SELECTED "
			    Else
				    strSelected = ""
			    End if
				    Response.Write "<option Value=""" & objRS("CostCentreID") & """" & strSelected & ">" & objRS("DivisionCode") & " - " & objRS("ProgramCode")  & " : " & objRS("CostCentreName") & "</OPTION>"
			    objRS.Movenext
		    Loop
    		
		    objRS.Close

	        %></select>
	    </td>
</tr>
	<th style="text-align:left; height:20px; width:50%;">&nbsp;Account Class</th>
		<td style="text-align:left; height:20px; width:50%;">
		    <select Style="Width:100%" tabindex="20" id="Level1ID" name="Level1ID" onchange="self.location='Indexation.asp?GLCode=' + frm.GLCode.value + '&Level1ID=' + frm.Level1ID.value + '&CostCentreID=' + frm.CostCentreID.value"><OPTION Value=0>Please Select..</OPTION>
	        <%		
		    objRS.Open "SELECT DISTINCT Level1ID,Level1Name FROM tblReportLayoutLevel1 WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND ReportID = 1 Order By Level1ID",objCon

		    Do until objRS.EOF
				If clng(objRS("Level1ID")) = clng(Session("Level1ID")) Then
				   strSelected = " SELECTED "
			    Else
				    strSelected = ""
			   End if
				    Response.Write "<option Value=""" & objRS("Level1ID") & """" & strSelected & ">" & objRS("Level1ID")  & " - " & objRS("Level1Name") & "</OPTION>"
			    objRS.Movenext
		    Loop
    		
		    objRS.Close

	        %></select>
	    </td>
</tr>
<tr>
	<th style="text-align:left; height:20px; width:50%;">&nbsp;GL Code</th>
		<td style="text-align:left; height:20px; width:50%;">
		    <select Style="Width:100%" tabindex="20" id="GLCode" name="GLCode" onchange="self.location='Indexation.asp?GLCode=' + frm.GLCode.value + '&Level1ID=' + frm.Level1ID.value + '&CostCentreID=' + frm.CostCentreID.value"><OPTION Value=0>Apply to All for selected Account Class..</OPTION>
	        <%		
		    objRS.Open "SELECT * FROM qryReportLayoutLevel2List WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND Level1ID = " & Session("Level1ID") & " AND BusinessAreaID = " & Session("BusinessAreaID") & "",objCon

		    Do until objRS.EOF
			    If clng(objRS("GLCode")) = clng(Session("GLCode")) Then
				    strSelected = " SELECTED "
			    Else
				    strSelected = ""
			    End if
				    Response.Write "<option Value=""" & objRS("GLCode") & """" & strSelected & ">" & objRS("SegmentValue")  & " - " & objRS("GLCodeName") & "</OPTION>"
			    objRS.Movenext
		    Loop
    		
		    objRS.Close

	        %></select>
	    </td>
</tr>

      
        </table>
        <br>
        <br />
        <br />
        <br />

             
        <table width="100%" align="left" border="1" cellspacing="1" cellpadding="1">
		<tr>
        			
		
		<%For z=1 to 6%>
		<TH>BM<%=z%></TH><TD>&nbsp;<input style="text-align:left; height:20px" type="number" step="0.00" id="lngBM<%=z%>"  name="lngBM<%=z%>"  value="<%=lngBM(z)%>" placeholder="BM<%=z%>" ondblclick="CopyTo()" />
		<%Next%>
		</td>
		
		</tr>
        <tr>
     
		<%For z=7 to 12%>
		<TH>BM<%=z%></TH><TD>&nbsp;<input style="text-align:left; height:20px" type="number" step="0.00" id="lngBM<%=z%>" name="lngBM<%=z%>"  value="<%=lngBM(z)%>" placeholder="BM<%=z%>"  />
		<%Next%>
		</td>
		
		</tr>
        </table>
        <br>
        <br />
        <br />
          <table width="100%" align="left" border="1" cellspacing="1" cellpadding="1">
		<tr>				
	
		<%For z=1 to 5%>
		<TH style="width:20%;">OY<%=z%></TH><td>&nbsp;<input style="text-align:left; height:20px" type="number" step="0.00" name="lngOY<%=z%>" value="<%=lngOY(z)%>"  placeholder="OY<%=z%>"  />
		<%Next%>
		</td>
		</tr>
		<tr>
		<th style="text-align:left; height:20px">&nbsp;Comments</th>
		<td colspan="9" >
		    <textarea name="lngComments" cols="200" rows="5"><%=lngComments%></textarea>
        </td>
		</tr>
		
	</tr>

	<tr>
		<td style="height:20px" colspan="10" align="left">&nbsp;</td>
	</tr>
	
</table>
<br/>
<br/>
<br/>
<br/>
<br/>
</br>
<table Class="CallCentre" WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
	<tr>
	    <td class='locked' Width="100px" style="border-right:0px"><button type="button" tabindex="8" onClick="self.location='AdminMenu.asp';"><img src="../images/door.png" alt="" /> Close </button></td>    
        <td class='locked' Width="100px" style="border-right:0px"><button type="button" tabindex="9" onClick="SaveData()";><img src="../images/tick.png" alt="" /> Save </button></td>  
		<td class='locked' Width="100px" style="border-right:0px"><button type="button" tabindex="9" onClick="DeleteData()";><img src="../images/cross.png" alt="" /> Delete </button></td> 
        <td class='locked' Width="100px" style="border-right:0px"><button type="button" tabindex="9" onClick="Reset()";><img src="../images/rubber.jpg" alt="" /> Reset </button></td> 

		<TD class='locked' Width="100px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=strMessageIcon %></TD>
        <TD class='locked' Width="400px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=strMessage %></TD>
	</tr>
</table>
<hr />
</form>

<table WIDTH="100%" BORDER="1" CELLSPACING="1" CELLPADDING="1">

	<tr>
        <th>Cost Centre</th>
        <th>Level</th>
        <th>GL Code</th>
        <th>BM1</th>
        <th>BM2</th>
        <th>BM3</th>
        <th>BM4</th>
        <th>BM5</th>
        <th>BM6</th>
        <th>BM7</th>
        <th>BM8</th>
        <th>BM9</th>
        <th>BM10</th>
        <th>BM11</th>
        <th>BM12</th>
        <th>Comments</th>
    </tr>
    <tr>
<%
	'Dynamically build the menu items depending on the sort selection 

    'If IsNull(Session("Level1ID")) = False AND IsNull(Session("CostCentreID")) = False AND IsNull(Session("GLCode")) = False AND Session("GLCode") <> 0 AND Session("CostCentreID") <> 0 Then
        'objRS.Open "SELECT * FROM tblIndexation WHERE BudgetID = " & clng (Session("BudgetID"))  & " And VersionID=" & Session("VersionID")  & " AND CostCentreID = " & Session("CostCentreID") & " AND Level1ID = " & Session("Level1ID") & "  AND GLCode = " & Session("GLCode") & ""

           ' If objRS.EOF Then
                'Response.Write "INSERT INTO tblIndexation VALUES (" & Session("BudgetID") & "," & Session("VersionID") & "," & Session("BusinessAreaID") & "," & Session("CostCentreID") & "," & Session("Level1ID") & "," & Session("GLCode") & ",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,''," & Session("UserID") & ",'" & now() & "')"
                'objCon.Execute "INSERT INTO tblIndexation VALUES (" & Session("BudgetID") & "," & Session("VersionID") & "," & Session("BusinessAreaID") & "," & Session("CostCentreID") & "," & Session("Level1ID") & "," & Session("GLCode") & ",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,''," & Session("UserID") & ",'" & now() & "')"
           ' End If

        'objRS.Close
    'End If

    objRS.Open "SELECT * FROM tblIndexation WHERE BudgetID = " & clng (Session("BudgetID"))  & " And VersionID=" & Session("VersionID")  & " AND CostCentreID = " & Session("CostCentreID") & " Order By CostCentreID,Level1ID,GlCode"
    'Response.Write "SELECT * FROM tblIndexation WHERE BudgetID = " & clng (Session("BudgetID"))  & " And VersionID=" & Session("VersionID")  & " AND CostCentreID = " & Session("CostCentreID") & " Order By CostCentreID,Level1ID,GlCode"
	Do until objRS.eof	
		
   	    Response.Write "<TR><TD><A Target=""_self"" HREF=""Indexation.asp?GLCode=" & objRS("GLCode") & "&=Level1ID=" & objRS("Level1ID") & "&CostCentreID=" & objRS("CostCentreID") & """>&nbsp;" & objRS("CostCentreID") & "</TD><TD>&nbsp;" & objRS("Level1ID") & "</TD><TD>&nbsp;" & objRS("GLCode") & "</TD><TD>&nbsp;" & objRS("BM1") & "</TD><TD>&nbsp;" & objRS("BM2") & "</TD><TD>&nbsp;" & objRS("BM3") & "</TD><TD>&nbsp;" & objRS("BM4") & "</TD><TD>&nbsp;" & objRS("BM5") & "</TD><TD>&nbsp;" & objRS("BM6") & "</TD><TD>&nbsp;" & objRS("BM7") & "</TD><TD>&nbsp;" & objRS("BM8") & "</TD><TD>&nbsp;" & objRS("BM9") & "</TD><TD>&nbsp;" & objRS("BM10") & "</TD><TD>&nbsp;" & objRS("BM11") & "</TD><TD>&nbsp;" & objRS("BM12") & "</TD><TD style=""text-align:center"">" & objRS("Comments")  & "</TD></TR>"
       
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
       ' Response.Write "SELECT * FROM tblIndexation WHERE BudgetID = " & Session("BudgetID")  & " And VersionID=" & session("VersionID") & " And Level1ID=" & Session("Level1ID") & " And CostCentreID= " & Session("CostCentreID") &  " AND GLCode = " & Session("GLCode") & "" 
		objRS.Open "SELECT * FROM tblIndexation WHERE BudgetID = " & Session("BudgetID")  & " And VersionID=" & session("VersionID") & " And Level1ID=" & Session("Level1ID") & " And CostCentreID= " & Session("CostCentreID") &  " AND GLCode = " & Session("GLCode") & "",objCon
		
			If Not objRS.EOF Then
			
			  Session("Level1ID") = objRS("Level1ID")
              Session("GLCode") = objRS("GLCode")
              Session("CostCentreID") = objRS("CostCentreID")
              lngComments = objRS("Comments")

				 For z=1 to 12
					lngBM(z)=objRS("BM"&z)
				 Next
				 For z=1 to 5
					lngOY(z)=objRS("OY"&z)
				 Next
				
			Else
			 'lngCalculatedFieldName  = ""
				
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
	  objRS.Open "SELECT * FROM tblIndexation WHERE BudgetID = " & Session("BudgetID")  & " And VersionID=" & session("VersionID") & " And Level1ID='" & request.form("Level1ID") & "' And CostCentreID="&  Request.Form("CostCentreID") &  " AND GLCode = " & Request.Form("GLCode") & "",objCon
	  If objRS.Eof=False Then
	  
		
          objCon.Execute("UPDATE tblIndexation SET UpdatedBy = " & clng(Session("UserID")) &  BMUpdate &  OYUpdate &", Comments='" & Request.Form("lngComments") & "', DateUpdated = GetDate() WHERE  BudgetID = " & clng(Session("BudgetID"))  & " And VersionID=" & session("VersionID") &  " And Level1ID='" &   Request.Form("Level1ID")  & "' And CostCentreID= "&  Request.Form("CostCentreID") & " AND GLCode = " & Request.Form("GLCode") & "") 
	     ' Response.Write "UPDATE tblIndexation SET BM24=BM12 WHERE  BudgetID = " & clng(Session("BudgetID"))  & " And VersionID=" & session("VersionID") &  " And Level1ID='" &   Request.Form("Level1ID")  & "' And CostCentreID= "&  Request.Form("CostCentreID") & " AND GLCode = " & Request.Form("GLCode") & ""
      Else

          objCon.Execute "INSERT INTO tblIndexation VALUES (" & Session("BudgetID") & "," & Session("VersionID") & "," & Session("BusinessAreaID") & "," & Session("CostCentreID") & "," & Request.Form("Level1ID") & "," & Request.Form("GLCode") & "" & BMInsert2 & OYInsert2 & ",'" &  Request.Form("lngComments") &"'," & Session("UserID") & ",GetDate())"
		  'Response.Write "INSERT INTO tblIndexation VALUES (" & Session("BudgetID") & "," & Session("VersionID") & "," & Session("BusinessAreaID") & "," & Session("CostCentreID") & "," & Request.Form("Level1ID") & "," & Request.Form("GLCode") & "" & BMInsert2 & OYInsert2 & ",'" &  Request.Form("lngComments") &"'," & Session("UserID") & ",GetDate())"
         'objCon.Execute("Insert into  tblCalculatedFieldValues (BudgetID, VersionID, CalculatedFieldName, CostCentreID, Comments" & BMInsert1 & OYInsert1 & ", UpdatedBy,DateUpdated) values (" & Session("BudgetID")  & "," & session("VersionID")  & ",'" & request.form("lngCalculatedFieldName") & "'," &  Request.Form("lngCostCentreID")   & ",'" &  Request.Form("lngComments") &"'" & BMInsert2 & OYInsert2 & "," &   clng(Session("UserID")) & ",GetDate())")
	      'objCon.Execute("UPDATE tblCalculatedFieldValues SET BM24=BM12 WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND CostCentreID = " & Session("CostCentreID") & " AND CalculatedFieldName = '" & request.form("lngCalculatedFieldName") & "'") 
       ' Response.Write "UPDATE tblCalculatedFieldValues SET BM24=BM12 WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND CostCentreID = " & Session("CostCentreID") & " AND CalculatedFieldName = '" & request.form("lngCalculatedFieldName") & "'"
  
    End If  
	objRS.close

    If Session("CostCentreID") = 0 or Session("GLCode") = 0 Then
        objCon.Execute "spApplyGlobalIndexation " & Session("BudgetID") & "," & Session("VersionID") & "," & Session("BusinessAreaID") & "," & Session("CostCentreID") & "," & Session("Level1ID") & "," & Session("GLCode") & "," & Session("UserID") & ""
        Response.Write "spApplyGlobalIndexation " & Session("BudgetID") & "," & Session("VersionID") & "," & Session("BusinessAreaID") & "," & Session("CostCentreID") & "," & Session("Level1ID") & "," & Session("GLCode") & "," & Session("UserID") & ""
	End If
    		'Return the result of the Save Function.
     		strMessage = "<B>RECORD SAVED.</B>"
            strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"
	 
End Sub

Sub Reset()

  objCon.Execute("DELETE tblIndexation WHERE  BudgetID = " & clng(Session("BudgetID"))  & " AND BusinessAreaID = " & Session("BusinessAreaID") & "") 
 ' objCon.Execute("UPDATE tblCostCentreStatus SET StatusID = " & clng(StatusID) & ", UpdatedBy = " & clng(Session("UserID")) & ", DateUpdated = GetDate() WHERE  BudgetID = " & clng(Session("BudgetID")) & " and VersionID = " & clng(Session("VersionID")) & "") 
  strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"
  strMessage = "<B>Business Area Indexation has been reset.</B>"
  strMessage = UCASE(strMessage)

End Sub	

Sub DeleteRecord()

  objCon.Execute "DELETE tblIndexation WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND CostCentreID = " & Session("CostCentreID") & " AND GLCode = " & Session("GLCode") & " AND Level1ID = " & Session("Level1ID") & ""
  Session("GLCode") = 0

  strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"
  strMessage = "<B>RECORD HAS BEEN DELETED.</B>"
 
End Sub	


Set objRS = Nothing
Set objCon = Nothing


%>
