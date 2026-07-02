<%@ Language=VBScript %>
<% Option Explicit
	
	Response.Expires = -1500

   If IsEmpty(Session("Logon")) Then Response.Redirect("AccessDenied.asp")
   If IsEmpty(Session("BudgetID")) Then Response.Redirect("Timeout.asp")

	Session("CurrentPage") = "Home.asp"
	
Dim objCon
Dim objRS
Dim objRS1
Dim objRS2
Dim strSelected
Dim arrStatusID(5)
Dim arrStatusName(5)
Dim arrStatusImg(5)
Dim arrCompanyID(2)
Dim arrCompanyName(2)
Dim x
Dim lngBudgetStatus
Dim lngBaseBudget
Dim lngBudget
Dim strColour 
Dim lngBaseLineRecurrent
Dim lngBaseLineDevelopment
Dim lngOneOffRecurrent
Dim lngOneOffDevelopment
Dim arrLanguage(2)
Dim strManagerName
Dim intManagerID
Dim strManCStatus
	
arrLanguage(1) = "ENGLISH"
arrLanguage(2) = "SWAHILI"	

arrStatusName(1) = "Open"
arrStatusName(2) = "Completed"
arrStatusName(3) = "Rejected"
arrStatusName(4) = "Approved"
arrStatusName(5) = "Closed"

arrStatusID(1) = 1
arrStatusID(2) = 2
arrStatusID(3) = 3
arrStatusID(4) = 4
arrStatusID(5) = 5

arrStatusImg(1) = "<IMG SRC='images/open.png'"
arrStatusImg(2) = "<IMG SRC='images/ready.gif'"
arrStatusImg(3) = "<IMG SRC='images/cross.png'" 
arrStatusImg(4) = "<IMG SRC='images/tick.png'"	
arrStatusImg(5) = "<IMG SRC='images/Closed.png'"
    			

Set objCon = Server.CreateObject("ADODB.Connection")
Set objRS = Server.CreateObject("ADODB.Recordset")
Set objRS1 = Server.CreateObject("ADODB.Recordset")
Set objRS2 = Server.CreateObject("ADODB.Recordset")

objCon.Open Session("DBConnection")

	If Not IsEmpty(Request.QueryString("CostCentreID")) Then
		Session("CostCentreID") = Request.QueryString("CostCentreID")
	End If
		
	If Not IsEmpty(Request.QueryString("VersionID")) Then
		Session("VersionID") = Request.QueryString("VersionID")
	End If	

	If Not IsEmpty(Request.QueryString("BudgetID")) or not IsEmpty(Session("BudgetID")) Then
		If isNull(Session("BudgetID")) Then Session("BudgetID") = Request.QueryString("BudgetID")	 
	    
	    objRS.Open "SELECT * FROM qrySystemDefault WHERE BudgetID = " & Session("BudgetID") & "",objCon

		If Not objRS.EOF Then
		
		    Session("BudgetID") = objRS("BudgetID")
		    Session("BudgetName") = objRS("BudgetName")
		    Session("VersionID") = objRS("DefaultVersionID")
		    Session("VersionTypeID")= objRS("VersionTypeID")
		    Session("FinacialYearID") = objRS("FinancialYearID")
		    Session("FinancialYear") = objRS("ERPFiscalYear")
		    
	    End If
	    
	    objRS.Close

        objRS.Open "SELECT * FROM qryBudgets WHERE BudgetID = " & Session("BudgetID") & "",objCon
		
		If Not objRS.EOF Then
		
		    Session("BudgetName") = objRS("BudgetName")
			 Session("FinacialYearID") = objRS("FinancialYearID")
		    Session("FinancialYear") = objRS("ERPFiscalYear")
               
	    End If
	    
	    objRS.Close
	    'Open the Version table with the BudgetID (which has just been changed) to set the Session VersionID to a relevant ID for the selected Budget.
	    objRS.Open "SELECT * FROM tblVersion WHERE BudgetID = " & Session("BudgetID") & "",objcon
	    
	        If not objRS.EOF Then
		        Session("VersionID") = objRS("VersionID")
	            Session("VersionName") = objRS("VersionName")
			    Session("VersionTypeID") = objRS("VersionTypeID")
			    Session("ColumnLock") = objRS("ColumnLock")
			    Session("BaseBudgetVersionID") = objRS("BaseBudgetVersionID") 
	        End IF
	    
	    objRS.Close
		
	End If	

	If Not IsEmpty(Request.QueryString("BusinessAreaID")) Then
	
			Session("BusinessAreaID") = Request.QueryString("BusinessAreaID")
			Session("PreviousBAID") = Request.QueryString("BusinessAreaID")
				
			'Get default Cost Centre
			 objRS.Open "SELECT * FROM qryBusinessAreaHome WHERE BusinessAreaID = " & Session("BusinessAreaID") & " AND BudgetID = " & Session("BudgetID") & "",objcon
	 
	            If not objRS.EOF Then
	                Session("CostCentreID") = objRS("DefaultCostCentre")
	                Session("BusinessAreaName") = objRS("BusinessAreaName")
                    Session("BusinessAreaCode") = objRS("BusinessAreaCode")
	                Session("PLReportID") = objRS("PLReportID")
			        Session("BSReportID") = objRS("BSReportID")
			        Session("CFReportID") = objRS("CFReportID")
					strManagerName = objRS("FName") & " " & objRS("LName")
					intManagerID = objRS("ManagerID")
		   	        Session("ProjectID") = "0"
		   	        Session("ActivityID") = "0"
                    Session("CCCeilingsOn") = objRS("CostCentreCeilingsOn")
                    Session("ACCeilingsOn") = objRS("AccountClassCeilingsOn")
					Session("ProjectCeilingsOn") = objRS("ProjectCeilingsOn")
		        Else
	                Session("CostCentreID") = 0
	                Session("BusinessAreaName") = ""
	                Session("BusinessAreaCode") = ""
		   	        Session("ProjectID") = "0"
		   	        Session("ActivityID") = "0"
                    Session("CCCeilingsOn") = "N"
                    Session("ACCeilingsOn") = "N"
					Session("ProjectCeilingsOn") = "N"
		        End IF

	        objRS.Close

	Else

		 objRS.Open "SELECT * FROM qryBusinessAreaHome WHERE BusinessAreaID = " & Session("BusinessAreaID") & " AND BudgetID = " & Session("BudgetID") & "",objcon
		 	If not objRS.EOF Then
			 	strManagerName = objRS("FName") & " " & objRS("LName")
			End If
		 objRS.Close
	    
	End If	
	
	If Not IsEmpty(Request.QueryString("StatusID")) Then
		
		Session("StatusID") = Request.QueryString("StatusID")	
		
		
		'Update the Business Area Status

        If Request.QueryString("StatusID") > 1 Then
            objRS.Open "SELECT * FROM qryCostCentreApprovals WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND BusinessAreaID = " & Session("BusinessAreaID") & " AND StatusID <> 2",objCon

                If Not objRS.EOF Then
                    Response.Write "&nbsp;&nbsp;<img src=""images/warning.gif"" /><B><FONT Color=""Red"">&nbsp;&nbsp;" & Session("BAName") & " STATUS CANNOT BE CHANGED TO UNTIL ALL " & Session("CCName") & "'S HAVE THE SAME STATUS.</FONT></B><BR><BR>" 
                Else
					If Request.QueryString("StatusID") = 4 Then
					
							objRS1.Open "SELECT ManagerID FROM tblBusinessArea WHERE BudgetID = " & Session("BudgetID") & " AND BusinessAreaID = " & Session("BusinessAreaID") & "",objCon
								If Not objRS1.EOF Then
									intManagerID = objRS1(0)
								End If
							objRS1.Close

							If cint(intManagerID) = cint(Session("UserID")) Then
								strManCStatus = Check_Man_Ceiling()
								If strManCStatus = "OK" Then
									objCon.Execute "UPDATE tblBusinessAreaStatus SET StatusID = " & Request.QueryString("StatusID") & " WHERE BudgetID = " & Session("BudgetID") & " AND BusinessAreaID = " & Session("BusinessAreaID") & " AND VersionID = " & Session("VersionID") & ""
								Else
									Response.Write "&nbsp;&nbsp;<img src=""images/warning.gif"" /><B><FONT Color=""Red"">&nbsp;&nbsp;" & strManCStatus & "</FONT></B><BR><BR>" 
								End If
							
							Else
								Response.Write "&nbsp;&nbsp;<img src=""images/warning.gif"" /><B><FONT Color=""Red"">&nbsp;&nbsp;ONLY " & Session("BAName") & " MANAGER CAN APPROVE.</FONT></B><BR><BR>" 
							End If
					
					Else
						objCon.Execute "UPDATE tblBusinessAreaStatus SET StatusID = " & Request.QueryString("StatusID") & " WHERE BudgetID = " & Session("BudgetID") & " AND BusinessAreaID = " & Session("BusinessAreaID") & " AND VersionID = " & Session("VersionID") & ""
					End If
				
				End If
        
            objRS.Close
	    Else
    	    objCon.Execute "UPDATE tblBusinessAreaStatus SET StatusID = " & Request.QueryString("StatusID") & " WHERE BudgetID = " & Session("BudgetID") & " AND BusinessAreaID = " & Session("BusinessAreaID") & " AND VersionID = " & Session("VersionID") & ""
		End if
	End If
	
    'Get Business Area Status
      'Response.Write "SELECT * FROM tblBusinessAreaStatus WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND BusinessAreaID = " & Session("BusinessAreaID") & ""
	    objRS.Open "SELECT * FROM tblBusinessAreaStatus WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND BusinessAreaID = " & Session("BusinessAreaID") & "",objCon
    
        If not objRS.EOF Then
            Session("StatusID") = objRS("StatusID")
		Else
            Session("StatusID") = 0
        End IF
    
    objRS.Close
    
    'Get Business Area Ceiling
       
    objRS.Open "SELECT BMCeiling FROM tblBACeilingLevel1 WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND Level1ID = " & Session("BusinessAreaID") & "",objCon
    
        If Not objRS.EOF Then
            Session("BACeiling") = objRS("BMCeiling")
        Else
            Session("BACeiling") = 0
        End If
        
    objRS.Close  
    
 %>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="BERTStyle.css">
</HEAD>
<BODY>
<FORM action="Home.asp" method=POST id=frm name=frm>
<TABLE WIDTH=100% ALIGN=Center BORDER=1 CELLSPACING=1 CELLPADDING=1>
<TR>
		<TH Height="25px">FBT Year</TH>
		<TD Width="30%"><select style="width:80%" id="BudgetID" name="BudgetID" onchange="self.location='Home.asp?BudgetID=' + frm.BudgetID.value">
		
<%	
	objRS.Open "SELECT * FROM tblBudget",objCon

	Do until objRS.EOF
		If objRS("BudgetID") = clng(Session("BudgetID")) Then
			strSelected = " SELECTED "
			Session("FinancialYearID") = objRS("FinancialYearID")
		Else
			strSelected = ""
		End if
			Response.Write "<option Value=""" & objRS("BudgetID") & """" & strSelected & ">" & objRS("BudgetName") & "</OPTION>"
		objRS.Movenext
	Loop
	
	objRS.Close	
%>
</select>
</TD>
<TH>Language</TH>
<TD Align="Right"><b><FONT Color="Blue">&nbsp;<%=arrLanguage(Session("Language"))%></FONT></B></TD>
	</TR>
	<TR>
		<TH Style="Width:15%; Height:25px">Version</TH>
		<TD Width=35%><select style=width:80%  id="Version" name="Version" onchange="self.location='Home.asp?VersionID=' + frm.Version.value">
<%	

	objRS.Open "SELECT * FROM tblVersion WHERE BudgetID = " & Session("BudgetID") & "",objCon
	
	Do until objRS.EOF
		If objRS("VersionID") = clng(Session("VersionID")) Then
			strSelected = " SELECTED "
			Session("VersionName") = objRS("VersionName")
			Session("VersionTypeID") = objRS("VersionTypeID")
			Session("ColumnLock") = objRS("ColumnLock")
		Else
			strSelected = ""
		End if
			Response.Write "<option Value=""" & objRS("VersionID") & """" & strSelected & ">" & objRS("VersionName") & "</OPTION>"
		objRS.Movenext
	Loop
	
	objRS.Close
	
%>
</select>
</TD>

<TH Width="15%"><%=Session("BAName")%> Balance</TH>
<TD Width="35%" style="text-align:left">&nbsp;<FONT Color="<%=strColour%>"><%=Vote_Balance%></FONT>&nbsp;&nbsp;(Expenditure :&nbsp;<%=Vote_Total %>&nbsp;Ceiling :&nbsp;<%=Vote_Ceiling %>)</TD>
</TR>
	<TR>
		<TH Style="Width:15%; Height:25px"><%=Session("BAName") %></TH>
		<TD Width="30%"><select style="width:80%"  id="BusinessArea" name="BusinessArea" onchange="self.location='Home.asp?BusinessAreaID=' + frm.BusinessArea.value">
<%	

    objRS.Open "SELECT * FROM qryBusinessAreaAccess WHERE BudgetID = " & Session("BudgetID") & " AND UserID = " & Session("UserID") & "",objCon
		'Response.Write "SELECT * FROM qryBusinessAreaAccess WHERE BudgetID = " & Session("BudgetID") & " AND UserID = " & Session("UserID") & ""
	Do until objRS.EOF
		If objRS("BusinessAreaID") = clng(Session("BusinessAreaID")) Then
			strSelected = " SELECTED "			
			Session("Vote") = objRS("BusinessAreaCode")
		Else
			strSelected = ""
		End if
			Response.Write "<option Value=""" & objRS("BusinessAreaID") & """" & strSelected & ">" & objRS("BusinessAreaCode") & " - " & objRS("BusinessAreaName") & "</OPTION>"
		objRS.Movenext
	Loop
	
	objRS.Close
	
%>
</select>
</TD>
<TH >User ID</TH>
<TD Align="Right"><b><FONT Color="Blue">&nbsp;<%=Session("Logon")%></FONT></B></TD>
	</TR>
	<TR>
		<TH Style="Width:15%; Height:25px"><%=Session("BAName")%> Status</TH>
				<TD Width="30%"><select style="width:80%" id="Status" name="Status" onchange="self.location='Home.asp?StatusID=' + frm.Status.value">
		
<%	

    If Session("StatusID") = 5 Then
		For x = 5 to 5
			If arrStatusID(x) = clng(Session("StatusID")) Then
				strSelected = " SELECTED "
			Else
				strSelected = ""
			End If
			Response.Write "<option Value=""" & arrStatusID(x) & """" & strSelected & ">" & arrStatusName(x) & "</OPTION>"
		Next
	Else If Session("StatusID") = 4 Then
        For x = 4 to 4
			If arrStatusID(x) = clng(Session("StatusID")) Then
				strSelected = " SELECTED "
			Else
				strSelected = ""
			End If
			Response.Write "<option Value=""" & arrStatusID(x) & """" & strSelected & ">" & arrStatusName(x) & "</OPTION>"
		Next
   
    Else	
		For x = 1 to 4
			If x <> 5 Then
				If arrStatusID(x) = clng(Session("StatusID")) Then
					strSelected = " SELECTED "
				Else
					strSelected = ""
				End If
				Response.Write "<option Value=""" & arrStatusID(x) & """" & strSelected & ">" & arrStatusName(x) & "</OPTION>"
			End If
		Next
	End If
End If

%></select>&nbsp;&nbsp;
<%=arrStatusImg("" & Session("StatusID") & "")%></TD>
<TH >Manager</TH>
<TD Align="Right"><b><FONT Color="Gray">&nbsp;<%=strManagerName%></FONT></B></TD>
	</TR>
</TABLE>

<HR>
<TABLE Align="Center" WIDTH="100%" BORDER="1" CELLSPACING="1" CELLPADDING="1">	

	<TR>
		<TH Style="Width:40%;Height:25px;">Dashboard</TH><TH Width="60%">Workflow Status</TH>
	</TR>
	<TR>
		<TD Colspan="2">&nbsp;</TD> 
	</TR>
	<tr>
		<TD ><iframe id="Iframe1" name="framecontent" src="<%=Session("HomePage1")%>" Width="100%" frameborder="0" height="400px"></iframe></TD>
		<TD ><iframe id="framecontent" name="framecontent" src="<%=Session("HomePage2")%>" Width="100%" frameborder="0" height="400px"></iframe></TD>
	</TR>
	
   <tr><th Style="Height:25px;"  colspan="2">&nbsp;</th></tr>
</TABLE>

</BODY>
</HTML>
<%


Function Vote_Balance()

Dim dblVoteBalance
Dim dblVoteCeiling
Dim bal
Dim strColour

      ' Get Vote Balance and Ceiling  
     objRS.Open "SELECT SUM(BMTotal) FROM qryCalculatedFieldBudgetDataBABySBC WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND BusinessAreaID = " & Session("BusinessAreaID") & " AND CalculatedField = 'Expenditure'",objCon
      'Response.Write "SELECT SUM(Budget) FROM qryExpenditureByVote WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND BusinessAreaID = " & Session("BusinessAreaID") & " AND TransactionType In ('RGEXP','WAGE') AND FundSource = '" & Session("FundSourceID") & "' Group By FundSource"     
        If Not objRS.EOF Then
            dblVoteBalance = objRS(0)
            If IsNull(dblVoteBalance) Then dblVoteBalance = 0
        Else
            dblVoteBalance = 0
        End If   

        If IsNull(dblVoteBalance) Then dblVoteBalance  = 0 End If
                    
     objRS.Close  
     
        If Session("BusinessAreaID") = 1000 Then
            objRS.Open "SELECT SUM(BMCeiling) FROM tblBACeilingLevel1 WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND BusinessAreaID = 1000 AND Approved = 'on'",objCon
        Else
            objRS.Open "SELECT SUM(BMCeiling) FROM tblBACeilingLevel1 WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND BusinessAreaID = 1000 AND Level1ID = " & Session("BusinessAreaID") & " AND Approved = 'on'",objCon
        End If
          
        If Not objRS.EOF Then
            dblVoteCeiling = objRS(0)
            If IsNull(dblVoteCeiling) Then dblVoteCeiling = 0
        Else
            dblVoteCeiling = 0
        End If

        If IsNull(dblVoteCeiling) Then dblVoteCeiling  = 0 End If
                    
     objRS.Close      
     
    bal = dblVoteCeiling - dblVoteBalance 
    If bal <> 0 Then
        strColour = "Red"
    Else
        strColour = "Green"
    End If
    Vote_Balance = "<B><FONT Color=" & strColour & ">" & formatnumber(bal,0) & "</FONT></B>"

End Function

Function Vote_Total()

Dim dblVoteBalance
Dim dblVoteCeiling
Dim bal
Dim strColour

      ' Get Vote Balance and Ceiling  
     objRS.Open "SELECT SUM(BMTotal) FROM qryCalculatedFieldBudgetDataBABySBC WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND BusinessAreaID = " & Session("BusinessAreaID") & " AND CalculatedField = 'Expenditure'",objCon
      'Response.Write "SELECT SUM(Budget) FROM qryExpenditureByVote WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND BusinessAreaID = " & Session("BusinessAreaID") & " AND TransactionType In ('REXP','DEXP') AND FundSource = '" & Session("FundSourceID") & "' Group By FundSource"     
        If Not objRS.EOF Then
            dblVoteBalance = objRS(0)
            If IsNull(dblVoteBalance) Then dblVoteBalance = 0
        Else
            dblVoteBalance = 0
        End If   

        If IsNull(dblVoteBalance) Then dblVoteBalance  = 0 End If
                    
     objRS.Close    
        
     
    bal = dblVoteBalance 
    If bal <> 0 Then
        strColour = "Red"
    Else
        strColour = "Green"
    End If
    Vote_Total = "<B>" & formatnumber(bal,0) & "</B>"

End Function


Function Vote_Ceiling()

Dim dblVoteBalance
Dim dblVoteCeiling
Dim bal
Dim strColour

    If Session("BusinessAreaID") = 1000 Then
        objRS.Open "SELECT SUM(BMCeiling) FROM tblBACeilingLevel1 WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND BusinessAreaID = 1000 AND Approved = 'on'",objCon
    Else
        objRS.Open "SELECT SUM(BMCeiling) FROM tblBACeilingLevel1 WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND BusinessAreaID = 1000 AND Level1ID = " & Session("BusinessAreaID") & " AND Approved = 'on'",objCon
    End If      
        If Not objRS.EOF Then
            dblVoteCeiling = objRS(0)
            If IsNull(dblVoteCeiling) Then dblVoteCeiling = 0
        Else
            dblVoteCeiling = 0
        End If

        If IsNull(dblVoteCeiling) Then dblVoteCeiling  = 0 End If
                    
     objRS.Close      
     
    bal = dblVoteCeiling 
    If bal <> 0 Then
        strColour = "Red"
    Else
        strColour = "Green"
    End If
    Vote_Ceiling = "<B>" & formatnumber(bal,0) & "</B>"

End Function

Function Check_Man_Ceiling()

Dim dblManBal
Dim dblManCeiling

	objRS2.Open "SELECT * FROM tblBACeilingLevel2 WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND BusinessAreaID = " & Session("BusinessAreaID") & " AND Mandatory = 'Y'",objCon
	
		Do Until objRS2.EOF

			' Get Man Balance and Ceiling  
			objRS1.Open "SELECT SUM(BMTotal) FROM qryCalculatedFieldBudgetDataBA WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND BusinessAreaID = " & Session("BusinessAreaID") & " AND CalculatedField = '" & objRS2("LevelID") & "'",objCon
		
				If Not objRS1.EOF Then
					dblManBal = objRS1(0)
					If IsNull(dblManBal) Then dblManBal = 0
				Else
					dblManBal = 0
				End If   

				If IsNull(dblManBal) Then dblManBal  = 0 End If
						
			objRS1.Close

			

			If dblManBal <> objRS2("BMCeiling") Then
				strManCStatus = objRS2("LevelID") & " MANDATORY CEILING NEEDS TO BE MET."			
			End If		
			
			objRS2.Movenext
		
		Loop

	objRS2.Close

	If strManCStatus <> "" Then
		Check_Man_Ceiling = strManCStatus
	Else
		Check_Man_Ceiling = "OK"
	End If

End Function

%>
