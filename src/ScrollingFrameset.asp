<%@ Language=VBScript %>
<% Option Explicit
	
    If IsEmpty(Session("BudgetID")) Then Response.Redirect("Timeout.asp")
	Response.Expires = -1500
	
	Session("Header") = "ScrollingFrameset.asp"
	Session("PreviousPage") = Request.ServerVariables("HTTP_REFERER")

Dim objCon
Dim objRS
Dim strSelected
Dim x

Set objCon = Server.CreateObject("ADODB.Connection")
Set objRS = Server.CreateObject("ADODB.Recordset")

    objCon.Open Session("DBConnection")

	If Not IsEmpty(Request.QueryString("CurrentPage")) Then
		Session("CurrentPage") = Request.QueryString("CurrentPage")
	End If

	'Set the Business Area ID from the previously selected one for Bids screen.
	If Not IsEmpty(Request.QueryString("PreviousBAID")) Then
		If Not IsNull(Session("PreviousBAID")) AND Session("PreviousBAID") <> "" Then
			Session("BusinessAreaID") = Session("PreviousBAID")

            If Not IsNull(Session("PreviousCCID")) AND Session("PreviousCCID") <> 0 Then Session("CostCentreID") = Session("PreviousCCID")
            
			'Get default Cost Centre
			objRS.Open "SELECT DefaultCostCentre,BusinessAreaName,PLReportID,BSReportID,CFReportID FROM tblBusinessArea WHERE BudgetID = " & Session("BudgetID") & " AND BusinessAreaID = " & Session("BusinessAreaID") & "",objCon
			
			    If objRS.EOF Then
			        'Session("CostCentreID") = 0
				Session("BusinessAreaName") = "Business Area Missing"			
			    Else
			        'Session("CostCentreID") = objRS("DefaultCostCentre")
				    Session("BusinessAreaName") = objRS("BusinessAreaName")	
                    Session("BusinessAreaCode") = objRS("BusinessAreaCode")
				    Session("PLReportID") = objRS("PLReportID")
		   	        Session("BSReportID") = objRS("BSReportID")
		   	        Session("CFReportID") = objRS("CFReportID")	
                    Session("CCCeilingsOn") = objRS("CostCentreCeilingsOn")
                    Session("ACCeilingsOn") = objRS("AccountClassCeilingsOn")
			    End If
			
			objRS.Close
		End If
	End If

	If Not IsEmpty(Request.QueryString("HeaderMenu")) Then
		Session("HeaderMenu") = Request.QueryString("HeaderMenu")
	End If

	If Not IsEmpty(Request.QueryString("CostCentreID")) Then
		Session("CostCentreID") = Request.QueryString("CostCentreID")
		Session("EmployeeID") = 0
	End If

    If Not IsEmpty(Request.QueryString("FundingPurposeID")) Then
		Session("FundingPurposeID") = Request.QueryString("FundingPurposeID")
	End If

     If Not IsEmpty(Request.QueryString("FundingPurposeL2ID")) Then
		Session("FundingPurposeL2ID") = Request.QueryString("FundingPurposeL2ID")
	End If		
		
	If Not IsEmpty(Request.QueryString("VersionID")) Then
		Session("VersionID") = Request.QueryString("VersionID")
	End If	
	
	If Not IsEmpty(Request.QueryString("ModeID")) Then
		Session("ModeID") = Request.QueryString("ModeID")
	End If	
	
	If Not IsEmpty(Request.QueryString("TransactionType")) Then
		Session("TransactionType") = Request.QueryString("TransactionType")
	End If
	
	If Not IsEmpty(Request.QueryString("Level1ID")) Then
		Session("Level1ID") = Request.QueryString("Level1ID")
	End If

	If Not IsEmpty(Request.QueryString("BidID")) Then
		Session("BidID") = Request.QueryString("BidID")
	End If		
	
	If Not IsEmpty(Request.QueryString("DetailedGLCode")) Then
		Session("DetailedGLCode") = Request.QueryString("DetailedGlCode")
	End If	
	
	If Not IsEmpty(Request.QueryString("CalculatedField")) Then
		Session("CalculatedField") = Request.QueryString("CalculatedField")
	End If	
	
	If Not IsEmpty(Request.QueryString("ProjectID")) Then
		Session("ProjectID") = Request.QueryString("ProjectID")
	End If
	
	If Not IsEmpty(Request.QueryString("ActivityID")) Then
		Session("ActivityID") = Request.QueryString("ActivityID")
	End If
	
	If Not IsEmpty(Request.QueryString("SubBudgetClassID")) Then
		Session("SubBudgetClassID") = Request.QueryString("SubBudgetClassID")
	End If	
	
	If Not IsEmpty(Request.QueryString("GFSCodeID")) Then
		Session("GFSCodeID") = Request.QueryString("GFSCodeID")
	End If
	
	If Not IsEmpty(Request.QueryString("DepartmentID")) Then
		Session("DepartmentID") = Request.QueryString("DepartmentID")
	End If
	
	If Not IsEmpty(Request.QueryString("FundTypeID")) Then
		Session("FundTypeID") = Request.QueryString("FundTypeID")
	End If
	
	If Not IsEmpty(Request.QueryString("FundSourceID")) Then
		Session("FundSourceID") = Request.QueryString("FundSourceID")
	End If				
	
	If Not IsEmpty(Request.QueryString("SubProgrammeID")) Then
		Session("SubProgrammeID") = Request.QueryString("SubProgrammeID")
	End If	
    
	If Not IsEmpty(Request.QueryString("CCStatusID")) Then
		Session("CCStatusID") = Request.QueryString("CCStatusID")
	End If	

    If Not IsEmpty(Request.QueryString("Segment3ID")) Then
		Session("Segment3") = Request.QueryString("Segment3ID")
	End If
	
    If Not IsEmpty(Request.QueryString("TransferDocumentTypeID")) Then
		Session("TransferDocumentTypeID") = Request.QueryString("TransferDocumentTypeID")
	End If	
		
	If Not IsEmpty(Request.QueryString("BusinessAreaID")) Then
		
			Session("BusinessAreaID") = Request.QueryString("BusinessAreaID")
			Session("PreviousBAID") = Request.QueryString("BusinessAreaID")
			'Get default Cost Centre
			objRS.Open "SELECT DefaultCostCentre,BusinessAreaName,PLReportID,BSReportID,CFReportID  FROM tblBusinessArea WHERE BudgetID = " & Session("BudgetID") & " AND BusinessAreaID = " & Session("BusinessAreaID") & "",objCon
			
			    If objRS.EOF Then
			    
			        Session("CostCentreID") = 0
				    Session("BusinessAreaName") = "Business Area Missing"	
				    Session("ProjectID") = "0"
		   	        Session("ActivityID") = "0"	
		   	        Session("SubBudgetClassID") = "0"
		   	        Session("GFSCodeID") = 0		   	   
                    Session("CCCeilingsOn") = objRS("CostCentreCeilingsOn")
                    Session("ACCeilingsOn") = objRS("AccountClassCeilingsOn")
		   	        Session("ProjectCeilingsOn") = objRS("ProjectCeilingsOn")	
			    Else
			    
			        Session("CostCentreID") = objRS("DefaultCostCentre")
				    Session("BusinessAreaName") = objRS("BusinessAreaName")	
				    Session("PLReportID") = objRS("PLReportID")
		   	        Session("BSReportID") = objRS("BSReportID")
		   	        Session("CFReportID") = objRS("CFReportID")	
		   	        Session("ProjectID") = "0"
		   	        Session("ActivityID") = "0"
		   	        Session("SubBudgetClassID") = "0"
		   	        Session("GFSCodeID") = 0
		   	        Session("CCCeilingsOn") = "N"
                    Session("ACCeilingsOn") = "N"
					Session("ProjectCeilingsOn") = "N"
		   	        		
			    End If
			
			objRS.Close
		
	End If

        objRS.Open "SELECT StatusID FROM tblBusinessAreaStatus WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND BusinessAreaID = " & Session("BusinessAreaID") & "",objCon
        
                If objRS.EOF Then
			        Session("StatusID") = 0
				Else
			        Session("StatusID")	= objRS("StatusID")
			    End If
			
		objRS.Close	 
		
		 objRS.Open "SELECT * FROM tblVersion WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & "",objCon
        
                If Not objRS.EOF Then
			        Session("ColumnLock")= objRS("ColumnLock")
			    End If
			
		objRS.Close	 

	If Not IsEmpty(Request.QueryString("Header")) Then
		Session("Header") = Request.QueryString("Header")
	End If
	
	
    
 %>
<html>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<LINK rel="stylesheet" type="text/css" href="../BERTStyle.css">
<title>Isidore</title>
</HEAD>
<frameset FrameBorder="0" COLS="*%" ROWS="5%,95%,*%">
    	
	<frame NAME="Header" SCROLLING="No" Border="0" SRC=<%=Session("HeaderMenu")%>></frame>
	<frame NAME="Body" SCROLLING="Yes" Border="0" SRC="<%=Session("CurrentPage")%>"></frame>
</frameset>
<noframes>
<body >
<b>Isidore IT<p>
Sorry your browser doesn't support frames.
</b></body>
</noframes>
</html>
