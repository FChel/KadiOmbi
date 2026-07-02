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
Dim arrYesNo(2)
Dim arrFundingType(2)
Dim strActive
Dim strSort
Dim strOrder

arrYesNo(1) = "Y"
arrYesNo(2) = "N"

arrFundingType(1) = "R"
arrFundingType(2) = "E"


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
Dim lngGoalID
Dim lngBudgetID
Dim strClusterID
Dim strGoalName
Dim strGoalDesc
Dim strGoalNameL2
Dim strGoalDescL2
Dim lngAllocationMethod
Dim lngCeilingLevelID
Dim intSortOrder
Dim strCalculatedField
Dim strFundingType
'3. Capture Querysring variables

    If Not IsEmpty(Request.QueryString("GoalID")) Then

	   lngGoalID = clng(Request.QueryString("GoalID"))	 
    Else
        lngGoalID = 0

    End If

	'Execute save 	
	If Request.QueryString("Action") = "Save" Then
		SaveDetails()
	End If

	  'Execute save 	
	If Request.QueryString("Action") = "Delete" Then
        
		DeleteRecord(Request.QueryString("GoalID"))
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

	    if(frm.GoalID.value == 0 )
	    {
		    varAlert += "Please enter a Goal ID. \n \n";
		    document.getElementById('GoalID').style.backgroundColor="ff8080";
		    varSubmit = false;
	    }	
	    else document.getElementById('GoalID').style.backgroundColor="ffffff";		   		  
	   	
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
        if (isWhitespace(frm.GoalID.value)) {
            alert('Please select a record to DELETE!');
        } else {
            if (window.confirm('Would you like to DELETE the selected record?') == true) {

                self.location = "Goals.asp?Action=Delete&GoalID=" + frm.GoalID.value;
            }

        }
    }




//-->
</script>
</head>
<body onload=padding();>
<h3>Goal Administration Screen</h3>
<form action="Goals.asp?Action=Save" method="POST" id="frm" name="frm">

<table width="100%" align="left" border="1" cellspacing="1" cellpadding="1">
<tr>
    <th style="text-align:left; height:20px; width:20%;">&nbsp;Goal ID</th>				
	<td style="text-align:left; height:20px; width:30%;">&nbsp;<input style="text-align:left; height:20px" type="number"  id="GoalID" name="GoalID" value="<%=lngGoalID%>"  /></td>
    <td style="width:50%;"></td>
</tr>
<tr>
    <th style="text-align:left; height:20px; width:20%;">&nbsp;Goal Name</th>				
	<td style="text-align:left; height:20px; width:30%;">&nbsp;<input style="text-align:left; height:20px" type="text" id="GoalName" name="GoalName" value="<%=strGoalName%>"  /></td>
    <td style="width:50%;"></td>
</tr>
<tr>
    <th style="text-align:left; height:20px; width:20%;">&nbsp;Goal Desc</th>				
	<td style="text-align:left; height:20px; width:30%;" >&nbsp;<input style="text-align:left; height:20px" type="text" id="GoalDesc" name="GoalDesc" value="<%=strGoalDesc%>"  /></td>
    <td style="width:50%;"></td>
</tr>
<tr>
    <th style="text-align:left; height:20px; width:20%;">&nbsp;Goal Name L2</th>				
	<td style="text-align:left; height:20px; width:30%;" >&nbsp;<input style="text-align:left; height:20px" type="text" id="GoalNameL2" name="GoalNameL2" value="<%=strGoalNameL2%>"  /></td>
    <td style="width:50%;"></td>
</tr>
<tr>
    <th style="text-align:left; height:20px; width:20%;">&nbsp;Goal Desc L2</th>				
	<td style="text-align:left; height:20px; width:30%;" >&nbsp;<input style="text-align:left; height:20px" type="text" id="GoalNameL2" name="GoalDescL2" value="<%=strGoalDescL2%>"  /></td>
    <td style="width:50%;"></td>
</tr>
<tr>
	<th style="text-align:left; height:20px; width:20%;">&nbsp;Cluster</th>
		<td style="text-align:left; height:20px; width:30%;">
		    <select Style="Width:100%" tabindex="20" id="ClusterID" name="ClusterID"><OPTION Value="">Please Select..</OPTION>
	        <%	
		    objRS.Open "SELECT * FROM tblClusters WHERE BudgetID = " & Session("BudgetID") & " AND Active = 'Y'",objCon

		    Do until objRS.EOF
			    If objRS("ClusterID") = strClusterID  Then
				    strSelected = " SELECTED "
			    Else
				    strSelected = ""
			    End if
				    Response.Write "<option Value=""" & objRS("ClusterID") & """" & strSelected & ">" & objRS("ClusterID") & " : " & objRS("ClusterName")  & "</OPTION>"
			    objRS.Movenext
		    Loop
    		
		    objRS.Close
    		
	        %></select>
	    </td><td style="width:50%;"></td>
  </tr>
  <tr>
  
         <th style="text-align:left; height:20px; width:20%;">&nbsp;Active</th>		
		<td style="text-align:left; height:20px; width:30%;"><select Style="Width:40%" tabindex="9" id="Active" name="Active"><option Value="0">Please Select....</option>
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
        <td style="width:50%;"></td>	
        <tr><td colspan="5">&nbsp;  </td></tr>
</table>
<br></br>


<table  WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
	<tr>
	    <td class='locked' Width="100px" style="border-right:0px"><button type="button" tabindex="8" onClick="self.location='AdminMenu.asp';"><img src="../images/door.png" alt="" /> Close </button></td> 
        <td class='locked' Width="100px" style="border-right:0px"><button type="button" tabindex="9" onClick="SaveData()";><img src="../images/tick.png" alt="" /> Save </button></td>  
		<td class='locked' Width="100px"><button type="button" tabindex="23" onClick="self.location='Goals.asp?GoalID=0';"><img src="../images/page_white_stack.png" alt="" /> New&nbsp;&nbsp; </button></td>
		<td class='locked' Width="100px"><button type="button" tabindex="19" onclick="DeleteData()";><img src="../images/cross.png" alt="" /> Delete </button></td>
        <TD class='locked' Width="100px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=strMessageIcon %></TD>
        <TD class='locked' Width="400px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=strMessage %></TD>
	</tr>
</table>

</form>
<h3>Goals</h3>
<table WIDTH="100%" BORDER="1" CELLSPACING="1" CELLPADDING="1">
	<tr>
        <th>Goal ID</th>
        <th>Goal Name</th>
        <th>Goal Desc</th>
        <th>Goal Desc L2</th>
        <th>Goal Name L2</th>
        <th>Cluster</th>
        <th>Active</th>
    </tr>
    <tr>
<%
    objRS.Open "SELECT * FROM tblGoals WHERE BudgetID = " & clng (Session("BudgetID"))  & " Order By GoalID"

	Do until objRS.eof	
		
   	    Response.Write "<TR><TD><A Target=""_self"" HREF=""Goals.asp?GoalID=" & objRS("GoalID") & """>&nbsp;" & objRS("GoalID") & "</TD><TD>&nbsp;" & objRS("GoalName") & "</TD><TD style=""text-align:center"">" & objRS("GoalDesc") & "</TD><TD style=""text-align:center"">" & objRS("GoalNameL2")  & "</TD><TD style=""text-align:center"">" & objRS("GoalDescL2")  & "</TD><TD style=""text-align:center"">&nbsp;" & objRS("ClusterID") & "</TD><TD style=""text-align:center"">" & objRS("Active") & "</TD></TR>"
       
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
	
		objRS.Open "SELECT * FROM tblGoals WHERE BudgetID = " & Session("BudgetID")  & " And GoalID =" & lngGoalID & "",objCon
		
			If Not objRS.EOF Then
			
			  lngGoalID = objRS("GoalID")
			  strGoalName = objRS("GoalName")
              strGoalDesc = objRS("GoalDesc")
              strGoalNameL2 = objRS("GoalNameL2")
              strGoalDescL2 = objRS("GoalDescL2")
 			  strClusterID = objRS("ClusterID")
			  strActive= objRS("Active")
          
			Else
			 lngGoalID = ""
				
			End If

		objRS.Close


End Sub

Sub SaveDetails()

    objCon.Execute "spGoalSave " & Session("BudgetID") & "," & Request.Form("GoalID") & ",'" & Request.Form("GoalName") & "','" & Request.Form("GoalDesc") & "','" & Request.Form("GoalNameL2") & "', '" & Request.Form("GoalDescL2") & "','" & Request.Form("ClusterID") & "','" & Request.Form("Active") & "'," & Session("UserID") & ""
 
	'Return the result of the Save Function.
   	strMessage = "<B>Record Saved.</B>"
    strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"
   
	 
End Sub

Sub DeleteRecord(GoalID)

  objCon.Execute("DELETE tblGoals WHERE BudgetID = " & Session("BudgetID") & " AND GoalID = " & GoalID & "") 
 
  strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"
  strMessage = "<B>RECORD HAS BEEN DELETED.</B>"

End Sub	


Set objRS = Nothing
Set objCon = Nothing


%>
