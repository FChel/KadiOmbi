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



objRS.Open "SELECT DISTINCT(FUNDAppropriationID) FROM tblFundAppropriation WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & "",objCon

    Do until objRS.EOF
        objCon.Execute "spApplyEstablishmentFunding " & Session("BudgetID") & "," & Session("VersionID") & "," & objRS(0) & "," & Session("UserID") & ""
        objRS.MoveNext
    Loop

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


</head>
<body>
<h3>Position Funding Summary Screen</h3>

<table WIDTH="100%" BORDER="1" CELLSPACING="1" CELLPADDING="1">
	<tr>
	    <th style="text-align:center;height:20px">Position Name</th>		
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
   
	strSQL =  "SELECT * FROM qryEstablishmentFunding WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " Order By EstablishmentName"
    
    objRS.Open strSQL,objCon
    
		Do until objRS.eof
		    dblWTotal = dblWTotal + objRS("Weighting")
            dblCost = objRS("Cost")
            dblFunding = objRS("Funding")
            dblCostTotal = dblCostTotal + dblCost
            dblFundingTotal = dblFundingTotal + dblFunding

            If IsNull(dblCost) Then dblCost = 0 End If
            If IsNull(dblFunding) Then dblFunding = 0 End If

			Response.Write "<TR><TD><A Target='_self' HREF='EstablishmentFunding.asp?EstablishmentID=" & objRS("EstablishmentID") & "&FundID=" & objRS("FundAppropriationID") & "&EstablishmentFundingID=" & objRS("EstablishmentFundingID") & "'>" & objRS("EstablishmentName") & "</TD><TD><B>&nbsp;" & objRS("FundCode") & " - " & objRS("FundName") & " : " & objRS("Comments") & "</TD><TD style=""text-align:center"">" & objRS("Weighting") * 100 & "%</TD><TD style=""text-align:center"">" & formatnumber(dblCost,0) & "</TD><TD style=""text-align:center"">" & formatnumber(dblFunding,0) & "</TD></TR>"
			objRS.movenext
		Loop
			
	objRS.Close
	
	    If IsNull(dblWTotal) Then 
	        dblWTotal = 0
	    End If 
	    
	    Response.Write "<TR><TH Style=""Height:20px""Colspan=""3""></TH><TH>" & formatnumber(dblCostTotal,0) & "</TH><TH>" & formatnumber(dblFundingTotal,0) & "</TH></TR>"

%>
</table>

<table Class="CallCentre" WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
	<tr>
	    <td class='locked' Width="100px" style="border-right:0px"><button type="button" tabindex="8" onClick="self.location='Establishments.asp';"><img src="../images/door.png" alt="" /> Close </button></td>    
 	</tr>
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
