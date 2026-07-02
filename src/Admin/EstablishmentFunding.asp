<%@ Language=VBScript %>
<% Option Explicit %>
<!-- #Include file=../ADOVBS.inc -->

<%

Response.Expires = -1500

If IsEmpty(Session("EstablishmentFundingID")) Then Session("EstablishmentFundingID") = 0


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

Dim intFundID
Dim lngCostCentreID
Dim dblWeighting
Dim strEstablishmentName

intFundID = 0

'2. Capture Querystring variables	
If Not IsEmpty(Request.QueryString("FundID")) Then		
    'Session("BusinessAreaID") = Request.QueryString("BusinessAreaID")
   Session("FundID") = Request.QueryString("FundID")  	
End If   		

If Not IsEmpty(Request.QueryString("EstablishmentID")) Then		
    'Session("CostCentreID") = Request.QueryString("CostCentreID")
    Session("EstablishmentID") = clng(Request.QueryString("EstablishmentID"))	
End If		

If Not IsEmpty(Request.QueryString("EstablishmentFundingID")) Then		
    'Session("CostCentreID") = Request.QueryString("CostCentreID")
    Session("EstablishmentFundingID") = clng(Request.QueryString("EstablishmentFundingID"))	
End If		


'Execute save 	
If Request.QueryString("Action") = "Save" Then
    SaveDetails()
End If

'Execute save 	
If Request.QueryString("Action") = "Delete" Then
    deleteRecord()
End If


    LoadDetails()

    objRS.Open "SELECT * FROM tblEstablishments WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND EstablishmentID = " & Session("EstablishmentID") & "",objCon

        If Not objRS.EOF Then
            strEstablishmentName = objRS("EstablishmentName")
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
   
	if(frm.FundID.value=="0")
	{
       varAlert += "Fund Cannot Be Blank. \n \n";
       document.getElementById('FundID').style.backgroundColor="ff8080";
       varSubmit = false;
    }   
    else document.getElementById('FundID').style.backgroundColor="ffffff";

    if (isPositiveInteger(frm.Weighting.value) == false) {
        varAlert += "Please enter a numeric value for BM1. \n \n";
        document.getElementById('Weighting').style.backgroundColor = "ff8080";
        varSubmit = false;
    }
    else document.getElementById('Weighting').style.backgroundColor = "ffffff";
    
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
	self.location="EstablishmentFunding.asp?Show=True&EstablishmentID=" + frm.EstablishmentID.value + "&FundID=" + frm.FundID.value	
}

function deleteRecord()
    {
 	if(window.confirm('Confirm delete')==true){
	self.location="EstablishmentFunding.asp?Action=Delete"
	}
}
//-->
</script>
</head>
<body>
<h3>Position Funding Administration Screen
<%
	Response.write "for the currently selected Budget : <FONT color=Red>" &  Session("BudgetName") & "</FONT> and Version : <FONT color=Red>" & Session("VersionName") & "</FONT></H3>"
%>
<form action="EstablishmentFunding.asp?Action=Save&?FundID=" & intFundID & "&EstablishmentID=" & Session("EstablishmentID") & "" method="POST" id="frm" name="frm">

<TABLE WIDTH="50%" ALIGN="left" BORDER="1" CELLSPACING="1" CELLPADDING="1">
<tr>
        <th style="text-align:left; height:20px; width:20%;">&nbsp;Position Name</th>
		<td style="text-align:left; height:20px; width:30%;">&nbsp;<%=strEstablishmentName%></td>
	</tr>
<tr>
		<th style="text-align:left; height:20px; width:20%;">&nbsp;Funding Source</th>
		<td style="text-align:left; height:20px; width:30%;">
		    <select Style="Width:100%" tabindex="20" id="FundID" name="FundID" onchange="BAIDSearch()"><OPTION Value="0">Please Select..</OPTION>
	        <%	
		    objRS.Open "SELECT * FROM qryFundAppropriation WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND CostObjectID = " & Session("BusinessAreaID") & "  AND FundTransactionTypeID In (1)",objCon
    		
		    Do until objRS.EOF
			    If clng(objRS("FundAppropriationID")) = clng(Session("FundID")) Then
				    strSelected = " SELECTED "
			    Else
				    strSelected = ""
			    End if
				    Response.Write "<option Value=""" & objRS("FundAppropriationID") & """" & strSelected & ">" & objRS("FundCode") & " - " & objRS("FundName") & " : " & objRS("Comments") & "</OPTION>"
			    objRS.Movenext
		    Loop
    		
		    objRS.Close
    		
	        %></select>
	    </td>
	</tr>		
	
	<tr>
        <th style="text-align:left; height:20px; width:20%;">&nbsp;Weighting %</th>
		<td style="text-align:left; height:20px; width:30%;">&nbsp;<input style="text-align:left" style="width:50%" id="Weighting" name="Weighting" maxlength="9" TABINDEX="1" value="<%=dblWeighting%>"></td>
	</tr>
	<tr><td style="text-align:left;height:20px" colspan="4">&nbsp;</td></tr>
</table>
<br />
<br />
<br />
<br />
<br />
<table Class="CallCentre" WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
	<tr>
	    <td class='locked' Width="100px" style="border-right:0px"><button type="button" tabindex="8" onClick="self.location='Establishments.asp';"><img src="../images/door.png" alt="" /> Close </button></td>    
        <td class='locked' Width="100px" style="border-right:0px"><button type="button" tabindex="9" onclick="SaveData()";><img src="../images/tick.png" alt="" /> Save </button></td>  
        <td class='locked' Width="100px"><button type="button" tabindex="10" onClick="self.location='EstablishmentFunding.asp?FundID=0'" )""><img src="../images/page_white_stack.png" alt="" /> New&nbsp;&nbsp; </button></td>
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
		<th align="center">Funding Source</th>				
		<th align="center">Weighting</th>
        <th align="center">Cost</th>
		<th align="center">Funding</th>	
	</tr>
	
<%
    Dim strSQL 
    Dim dblWTotal
    Dim dblCost
    Dim dblFunding
    Dim dblCostTotal
    Dim dblFundingTotal

    
       
    dblWTotal = 0    
   
	strSQL =  "SELECT * FROM qryEstablishmentFunding WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND EstablishmentID = " & Session("EstablishmentID") & ""
    
    objRS.Open strSQL,objCon
    
		Do until objRS.eof
		    dblWTotal = dblWTotal + objRS("Weighting")
            dblCost = objRS("Cost")
            dblFunding = objRS("Funding")
            dblCostTotal = dblCostTotal + dblCost
            dblFundingTotal = dblFundingTotal + dblFunding

            If IsNull(dblCost) Then dblCost = 0 End If
            If IsNull(dblFunding) Then dblFunding = 0 End If

			Response.Write "<TR><TD><A Target='_self' HREF='EstablishmentFunding.asp?FundID=" & objRS("FundAppropriationID") & "&EstablishmentFundingID=" & objRS("EstablishmentFundingID") & "'><IMG SRC=""../images/edit.jpg""></TD><TD><B>&nbsp;" & objRS("FundCode") & " - " & objRS("FundName") & " : " & objRS("Comments") & "</TD><TD style=""text-align:center"">" & objRS("Weighting") * 100 & "%</TD><TD style=""text-align:center"">" & formatnumber(dblCost,0) & "</TD><TD style=""text-align:center"">" & formatnumber(dblFunding,0) & "</TD></TR>"
			objRS.movenext
		Loop
			
	objRS.Close
	
	    If IsNull(dblWTotal) Then 
	        dblWTotal = 0
	    End If 
	    
	    Response.Write "<TR><TH Style=""Height:20px""Colspan=""3""></TH><TH>" & formatnumber(dblCostTotal,0) & "</TH><TH>" & formatnumber(dblFundingTotal,0) & "</TH></TR>"

%>
</table>
</body>

</html>

<% 

Sub LoadDetails()

'Description:	Loads Caller's details into page if applicable.
		
		objRS.Open "SELECT * FROM qryEstablishmentFunding WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND FundAppropriationID = " & Session("FundID") & " AND EstablishmentID=" & clng(Session("EstablishmentID")) & "",objCon							
		'Response.Write "SELECT * FROM qryEstablishmentFunding WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND FundID = " & Session("FundID") & " AND EstablishmentID=" & clng(Session("EstablishmentID")) & ""
        If Not objRS.EOF Then
		   dblWeighting = objRS("Weighting") * 100  
           strEstablishmentName = objRS("EstablishmentName") & " : " & objRS("PositionName")   
           Session("FundID") = objRS("FundAppropriationID")	
           Session("EstablishmentID") = objRS("EstablishmentID")
         						
		Else		                
		   dblWeighting = 0 
           strEstablishmentName = ""   
           'Session("FundID") = 0
           'Session("EstablishmentID") = 0                   
		End if

		objRS.Close	
End Sub

Sub DeleteRecord()
  
        objCon.Execute "Delete from tblEstablishmentFunding WHERE FundAppropriationID = " & Session("FundID") & " AND EstablishmentID = " & Session("EstablishmentID") & " AND BudgetID = " & Session("BudgetID")
        'Response.Write "Delete from tblEstablishmentFunding WHERE FundAppropriationID = " & Session("FundID") & " AND EstablishmentID = " & Session("EstablishmentID") & " AND BudgetID = " & Session("BudgetID")
        strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"	
        strMessage = "<B>RECORD DELETED.</B>"

End Sub

Sub LoadERPData()
    
    objCon.Execute "spLoadERPCCBARelationship " & Session("BudgetID") & ",'N'," & Session("UserID") & ""
     
End Sub

Sub SaveDetails()	
		
		 With objCmd
                .CommandType = 4
                .CommandText = "spEstablishmentFundingSave"                  
                
                .Parameters.Append objCmd.CreateParameter("EstablishmentFundingID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("BudgetID", adInteger, adParamInput)
               
                .Parameters.Append objCmd.CreateParameter("VersionID", adInteger, adParamInput)   
                 .Parameters.Append objCmd.CreateParameter("EstablishmentID", adInteger, adParamInput)                               
                .Parameters.Append objCmd.CreateParameter("FundID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("Weighting", adDecimal, adParamInput,15)
                
                    With .Parameters.Item("Weighting")
                        .Precision = 10
                        .NumericScale = 7
                    End With

                .Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger)                                                                    
              	
                .Parameters("EstablishmentFundingID") = Session("EstablishmentFundingID")		
				.Parameters("BudgetID") = Session("BudgetID")		
                .Parameters("VersionID") = Session("VersionID") 
                .Parameters("EstablishmentID") = Session("EstablishmentID")            
                .Parameters("FundID") = Request.Form("FundID")
                .Parameters("Weighting") = Request.Form("Weighting")
                .Parameters("UpdatedBy") = Session("UserID")

                'Response.Write Session("EstablishmentFundingID")
                'Response.Write Session("EstablishmentID")    
                'Response.Write Request.Form("FundID") 
                           
                .ActiveConnection = objCon
                
            End With
                
            objCmd.Execute          
               			     				
     		strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"	
            strMessage = "<B>RECORD SAVED.</B>"								
    			
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
