<%@ Language=VBScript %>
<% Option Explicit %>
<!-- #Include file=../ADOVBS.inc -->

<%

Response.Expires = -1500

If IsEmpty(Session("BudgetID")) Then Response.Redirect("../Timeout.asp")
 
'Description:	Student Screen
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
Dim strVote

'Set Database objects

Set objCon = Server.CreateObject("ADODB.Connection")
Set objCmd = Server.CreateObject("ADODB.Command")
Set objRS = Server.CreateObject("ADODB.Recordset")

'Open database connection

objCon.Open Session("DBConnection")

'1. Declare screen specific variables naming convention = variable type prefix + database field name

Dim lngBusinessAreaID
Dim lngCostCentreID
Dim dblWeighting

lngBusinessAreaID = 0

'2. Capture Querystring variables	
If Not IsEmpty(Request.QueryString("BusinessAreaID")) Then		
    'Session("BusinessAreaID") = Request.QueryString("BusinessAreaID")
    lngBusinessAreaID = clng(Request.QueryString("BusinessAreaID"))	
Else
	'lngBusinessAreaID = 0'Session("BusinessAreaID")		    	
End If

   		

If Not IsEmpty(Request.QueryString("CostCentreID")) Then		
    'Session("CostCentreID") = Request.QueryString("CostCentreID")
    lngCostCentreID = clng(Request.QueryString("CostCentreID"))	
Else
    'lngCostCentreID = 0'Session("CostCentreID")				
End If		


'Execute save 	
If Request.QueryString("Action") = "Save" Then
    SaveDetails()
End If

'Execute save 	
If Request.QueryString("Action") = "Delete" Then
    deleteRecord()
End If

'Load page details
if IsEmpty(Request.QueryString("Show")) then
    'LoadDetails()
End if	

objRS.Open "SELECT BusinessAreaCode FROM tblBusinessArea WHERE BudgetID = " & Session("BudgetID") & " AND BusinessAreaID = " & lngBusinessAreaID & "",objCon

    If Not objRS.EOF Then
        strVote = objRS("BusinessAreaCode")
    Else
        strVote = ""
    End If

objRS.Close

%>

<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="../BERTStyle.css">
	<script src="../formChek.js">
	</script>
	<script src="../ButtonRollOver.js">
	</script>

<script LANGUAGE="javascript">
<!--
function SaveData()
{	    
    var varSubmit = true						
    var varAlert="";	
   
	if(frm.BusinessAreaID.value=="0")
	{
       varAlert += "Business Area Cannot Be Blank. \n \n";
       document.getElementById('BusinessAreaID').style.backgroundColor="ff8080";
       varSubmit = false;
    }   
    else document.getElementById('BusinessAreaID').style.backgroundColor="ffffff";
    
    if(frm.CostCentreID.value=="0")
	{
       varAlert += "Cost Centre Cannot Be Blank. \n \n";
       document.getElementById('CostCentreID').style.backgroundColor="ff8080";
       varSubmit = false;
    }   
    else document.getElementById('CostCentreID').style.backgroundColor="ffffff";  
    
   if(varSubmit == true)
	{
	    frm.submit();		
	}
	else
	{
	    alert(varAlert);
	}
}


function BAIDSearch()
{	        
	self.location="CCBARelationship.asp?Show=True&BusinessAreaID=" + frm.BusinessAreaID.value + "&CostCentreID=" + frm.CostCentreID.value	
}

function deleteRecord()
{
     if(frm.CostCentreID.value=="0" || frm.BusinessAreaID.value=="0")
	 {       
	    alert("Please select a record to delete");
     }   
     else
     {
        self.location="CCBARelationship.asp?Action=Delete&BusinessAreaID=" + frm.BusinessAreaID.value + "&CostCentreID=" + frm.CostCentreID.value
     }
}
//-->
</script>
</head>
<body>
<h3><%=Session("BAName") %> / <%=Session("CCName") %> Administration Screen
<%
	Response.write "for the currently selected Budget : <FONT color=Red>" &  Session("BudgetName") & "</FONT> and Version : <FONT color=Red>" & Session("VersionName") & "</FONT></H3>"
%>
<form action="CCBARelationship.asp?Action=Save&?BusinessAreaID=" & lngBusinessAreaID & "&CostCentreID=" & lngCostCentreID & "" method="POST" id="frm" name="frm">

<TABLE WIDTH="50%" ALIGN="left" BORDER="1" CELLSPACING="1" CELLPADDING="1">
<tr>
		<th style="text-align:left; height:20px; width:20%;">&nbsp;<%=Session("BANAme")%></th>
		<td style="text-align:left; height:20px; width:30%;">
		    <select Style="Width:40%" tabindex="20" id="BusinessAreaID" name="BusinessAreaID" onchange="BAIDSearch()"><OPTION Value=0>Please Select..</OPTION>
	        <%	
		    objRS.Open "SELECT * FROM tblBusinessArea WHERE BudgetID = " & Session("BudgetID") & " AND Active = 'Y'",objCon
    		
		    Do until objRS.EOF
			    If clng(objRS("BusinessAreaID")) = clng(lngBusinessAreaID) Then
				    strSelected = " SELECTED "
			    Else
				    strSelected = ""
			    End if
				    Response.Write "<option Value=""" & objRS("BusinessAreaID") & """" & strSelected & ">" & objRS("BusinessAreaCode") & " - " & objRS("BusinessAreaName") & "</OPTION>"
			    objRS.Movenext
		    Loop
    		
		    objRS.Close
    		
	        %></select>
	    </td>
	</tr>		
	
	<tr>
        <th style="text-align:left; height:20px; width:20%;">&nbsp;<%=Session("CCName") %></th>
		<td style="text-align:left; height:20px; width:30%;">
		    <select Style="Width:40%" tabindex="20" id="CostCentreID" name="CostCentreID" onchange="BAIDSearch()"><OPTION Value=0>Please Select..</OPTION>
	        <%	
            If lngBusinessAreaID <> 1000 Then
		        objRS.Open "SELECT * FROM tblCostcentres WHERE BudgetID = " & Session("BudgetID") & " AND CostCentreID > 0 AND Left(CostCentreID,4) = " & lngBusinessAreaID & " AND Active = 'Y'",objCon
    		Else
                objRS.Open "SELECT * FROM tblCostcentres WHERE BudgetID = " & Session("BudgetID") & " AND CostCentreID > 0 AND Active = 'Y'",objCon
            End If

		    Do until objRS.EOF
			    If clng(objRS("CostCentreID")) = clng(lngCostCentreID) Then
				    strSelected = " SELECTED "
			    Else
				    strSelected = ""
			    End if
				    Response.Write "<option Value=""" & objRS("CostCentreID") & """" & strSelected & ">" & objRS("CostCentreID") & " - " & objRS("CostCentreName") &"</OPTION>"
			    objRS.Movenext
		    Loop
    		
		    objRS.Close
    		
	        %></select>
	    </td>
	</tr>
	<tr><td style="text-align:left;height:20px" colspan="4">&nbsp;</td></tr>
</table>
<br />
<br />
<br />
<br />
<table Class="CallCentre" WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
	<tr>
	    <td class='locked' Width="100px" style="border-right:0px"><button type="button" tabindex="8" onClick="self.location='AdminMenu.asp';"><img src="../images/door.png" alt="" /> Close </button></td>    
        <td class='locked' Width="100px" style="border-right:0px"><button type="button" tabindex="9" onclick="SaveData()";><img src="../images/tick.png" alt="" /> Save </button></td>  
        <td class='locked' Width="100px"><button type="button" tabindex="10" onClick="self.location='CCBARelationship.asp?BusinessAreaID=0&CostCentreID=0'" )""><img src="../images/page_white_stack.png" alt="" /> New&nbsp;&nbsp; </button></td>
        <td class='locked' Width="100px"><button type="button" tabindex="11" onclick="deleteRecord()";><img src="../images/cross.png" alt="" /> Delete </button></td>
		<TD class='locked' Width="100px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=strMessageIcon%></TD>
        <TD class='locked' Width="300px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=strMessage%></TD>
	</tr>
</table>

<hr>
</form>

<table WIDTH="100%" BORDER="1" CELLSPACING="1" CELLPADDING="1">
	<tr>
	    <th style="text-align:center;height:20px">Edit</th>		
		<th align="center"><%=Session("BAName") %></th>				
		<th align="center"><%=Session("CCName") %></th>
	    <th align="center">Updated By</th>
	    <th align="center">Date Updated</th>						
	</tr>
	
<%
    Dim strSQL 
    Dim dblWTotal
       
    dblWTotal = 0
    
    If lngCostCentreID = 0 Then
		strSQL =  "SELECT * FROM qryCCBARelationship WHERE BudgetID = " & Session("BudgetID") & " AND BusinessAreaID=" & lngBusinessAreaID
    Else
        strSQL =  "SELECT * FROM qryCCBARelationship WHERE BudgetID = " & Session("BudgetID") & " AND CostCentreID=" & lngCostCentreID
    End If
    
    objRS.Open strSQL,objCon
    
		Do until objRS.eof
		    dblWTotal = dblWTotal + objRS("Weighting")
			Response.Write "<TR><TD><A Target='_self' HREF='CCBARelationship.asp?BusinessAreaID=" & objRS("BusinessAreaID") & "&CostCentreID="& objRS("CostCentreID") & "'><IMG SRC=""../images/edit.jpg""></TD><TD><B>&nbsp;" & objRS("BusinessAreaCode") & " - " & objRS("BusinessAreaName") & "</TD><TD>&nbsp;" & objRS("CostCentreID") & " - " & objRS("CostCentreName") & "</B></TD><TD style=""text-align:center"">" & objRS("UpdatedBy") & "</TD><TD style=""text-align:center"">" & objRS("DateUpdated") & "</TD></TR>"
			objRS.movenext
		Loop
			
	objRS.Close
	
	    If IsNull(dblWTotal) Then 
	        dblWTotal = 0
	    End If 
	    
	    Response.Write "<TR><TH Style=""Height:20px""Colspan=""5""></TH></TR>"

%>
</table>
</body>

</html>

<% 

Sub LoadDetails()

'Description:	Loads Caller's details into page if applicable.
		
		objRS.Open "SELECT * FROM tblCCBARelationship WHERE BudgetID = " & Session("BudgetID") & " AND BusinessAreaID = " & clng(Session("BusinessAreaID")) & " AND CostCentreID=" & clng(Session("CostCentreID")),objCon							
		If Not objRS.EOF Then
		    lngBusinessAreaID = objRS("BusinessAreaID")
            lngCostCentreID = objRS("CostCentreID") 
            dblWeighting = objRS("Weighting")        						
		Else		                
		    'Do nothing                         
		End if

		objRS.Close	
End Sub

Sub DeleteRecord()
  
        objCon.Execute("Delete from tblCCBARelationship where BusinessAreaID = " & lngBusinessAreaID & " AND CostCentreID = " & lngCostCentreID & " AND BudgetID = " & Session("BudgetID"))    
        strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"	
        strMessage = "<B>RECORD DELETED.</B>"

End Sub

Sub LoadERPData()
    
    objCon.Execute "spLoadERPCCBARelationship " & Session("BudgetID") & ",'N'," & Session("UserID") & ""
     
End Sub

Sub SaveDetails()	
		
		 With objCmd
                .CommandType = 4
                .CommandText = "spCCBARelationshipSave"                  
                
                .Parameters.Append objCmd.CreateParameter("BudgetID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("CostCentreID", adInteger, adParamInput)                                
                .Parameters.Append objCmd.CreateParameter("BusinessAreaID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("Weighting", adDecimal, adParamInput,15)
                
                    With .Parameters.Item("Weighting")
                        .Precision = 10
                        .NumericScale = 7
                    End With

                .Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger)                                                                    
              			
				.Parameters("BudgetID") = Session("BudgetID")		
                .Parameters("CostCentreID") = Request.Form("CostCentreID")                      
                .Parameters("BusinessAreaID") = Request.Form("BusinessAreaID")
                .Parameters("Weighting") = 0
                .Parameters("UpdatedBy") = Session("UserID")
                           
                .ActiveConnection = objCon
                
            End With
                
            objCmd.Execute          
               			     				
     		strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"	
            strMessage = "<B>RECORD SAVED.</B>"								
     		lngBusinessAreaID =  Request.Form("BusinessAreaID")
     		lngCostCentreID = Request.Form("CostCentreID")					
End Sub	

Function MediumDate (str)
	
	'Function to change all date formats to medium date to avoid American storage challenge!
	
	Dim aDay
	Dim aMonth
	Dim aYear
	
		aDay = 	(Left((str),InStr(1,(str),"/")-1))
		aMonth = Mid(str,(InStr(1,(str),"/")+1),2)
	
	If Right(aMonth,1) = "/" Then
		aMonth = Left(aMonth,1)
	End If
	
		aMonth = MonthName(aMonth)
		aYear = Year(str)
	
	If Len(aDay) = 1 Then aDay = "0" & aDay
	
		MediumDate = aDay & "-" & aMonth & "-" & aYear
		
End Function

Set objRS = Nothing
Set objCon = Nothing


%>
