<%@ Language=VBScript %>
<% Option Explicit
	
	Response.Expires = -1500
	
	Session("Header") = "BERTFrameset.asp"

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
	
	If Not IsEmpty(Request.QueryString("PLReportID")) Then
		Session("PLReportID") = Request.QueryString("PLReportID")
	End If

	If Not IsEmpty(Request.QueryString("BSReportID")) Then
		Session("BSReportID") = Request.QueryString("BSReportID")
	End If

	If Not IsEmpty(Request.QueryString("CFReportID")) Then
		Session("CFReportID") = Request.QueryString("CFReportID")
	End If

	If Not IsEmpty(Request.QueryString("HeaderMenu")) Then
		Session("HeaderMenu") = Request.QueryString("HeaderMenu")
	End If
	
	If Not IsEmpty(Request.QueryString("Currency")) Then
		Session("Currency") = Request.QueryString("Currency")
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

	If Not IsEmpty(Request.QueryString("BidType")) Then
		Session("BidType") = Request.QueryString("BidType")
	End If
	
	If Not IsEmpty(Request.QueryString("Level1ID")) Then
		Session("Level1ID") = Request.QueryString("Level1ID")
		Session("Level1Name") = Request.QueryString("Level1Name")
	End If	
	
	If Not IsEmpty(Request.QueryString("GLCode")) Then
		Session("GLCode") = Request.QueryString("GLCode")
	End If	
	
    If Not IsEmpty(Request.QueryString("FundingPurposeID")) Then
		Session("FundingPurposeID") = Request.QueryString("FundingPurposeID")
	End If	

	
	If Not IsEmpty(Request.QueryString("SubProgrammeID")) Then
		Session("SubProgrammeID") = Request.QueryString("SubProgrammeID")
	End If	
	
	 If Not IsEmpty(Request.QueryString("FundingPurposeID")) Then
		Session("FundingPurposeID") = Request.QueryString("FundingPurposeID")
	End If

     If Not IsEmpty(Request.QueryString("FundingPurposeL2ID")) Then
		Session("FundingPurposeL2ID") = Request.QueryString("FundingPurposeL2ID")
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
	
    If Not IsEmpty(Request.QueryString("RecordID")) Then
		Session("RecordID") = Request.QueryString("RecordID")
    Else
        Session("RecordID") = ""
	End If
    
     If Not IsEmpty(Request.QueryString("Action")) Then
		Session("Action") = Request.QueryString("Action")
    Else
        Session("Action") = ""
	End If					
	
	If Not IsEmpty(Request.QueryString("DetailedGLCode")) Then
		Session("DetailedGLCode") = Request.QueryString("DetailedGlCode")
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
			objRS.Open "SELECT * FROM tblBusinessArea WHERE BudgetID = " & Session("BudgetID") & " AND BusinessAreaID = " & Session("BusinessAreaID") & "",objCon
			
			    If objRS.EOF Then
			        
			        Session("CostCentreID") = 0
				    Session("BusinessAreaName") = "Business Area Missing"	
				    Session("SubProgrammeID") = 0
		   	        Session("ServiceID") = "0"
		   	        Session("ProjectID") = "0"
		   	        Session("ActivityID") = "0"	
		   	        Session("SubBudgetClassID") = "0"
		   	        Session("GFSCodeID") = 0
		   	        Session("DepartmentID") = "0"
		   	        Session("FundTypeID") = 0
		   	        Session("FundSourceID") = 0
                    Session("Vote") = ""
                    Session("CCCeilingsOn") = objRS("CostCentreCeilingsOn")
                    Session("ACCeilingsOn") = objRS("AccountClassCeilingsOn")
                    Session("BusinessAreaCode") = ""
		   	        
			    Else
			        
			        Session("CostCentreID") = objRS("DefaultCostCentre")
				    Session("BusinessAreaName") = objRS("BusinessAreaName")	
                    Session("BusinessAreaCode") = objRS("BusinessAreaCode")
				    Session("PLReportID") = objRS("PLReportID")
		   	        Session("BSReportID") = objRS("BSReportID")
		   	        Session("CFReportID") = objRS("CFReportID")	
		   	        Session("SubProgrammeID") = 0
		   	        Session("ServiceID") = "0"
		   	        Session("ProjectID") = "0"
		   	        Session("ActivityID") = "0"
		   	        Session("SubBudgetClassID") = "0"
		   	        Session("GFSCodeID") = 0
		   	        Session("DepartmentID") = "0"
		   	        Session("FundTypeID") = 0
		   	        Session("FundSourceID") = 0
                    Session("Vote") = objRS("BusinessAreaCode")
                    Session("CCCeilingsOn") = "N"
                    Session("ACCeilingsOn") = "N"
		   	        
			    End If
			
			objRS.Close	
			
			objRS.Open "SELECT Top 1 ProjectID FROM qryProjectAccess WHERE BudgetID = " & Session("BudgetID") & " AND BusinessAreaID = " & Session("BusinessAreaID") & "",objCon
			
			    If Not objRS.EOF Then
			        Session("ProjectID") = objRS("ProjectID")
			    Else
			        Session("ProjectID") = ""
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

	If Not IsEmpty(Request.QueryString("CostCentreID")) Then
	
		Session("CostCentreID") = Request.QueryString("CostCentreID")
		Session("SubProgrammeID") = 0
		objRS.Open "SELECT LocalCurrency FROM tblCostCentres WHERE BudgetID = " & Session("BudgetID") & " AND CostCentreID = " & Session("CostCentreID") & "",objCon
		
		    If Not objRS.EOF Then
		        Session("Currency") = objRS("LocalCurrency")
		    End If
		
		objRS.Close
		
	End If

	If Session("CurrentPage") = "ProfitLoss/DataEntry7.asp" OR Left(Session("CurrentPage"),15) = "BudgetTransfers" Then
	'If IsEmpty(Session("Scroll")) Then
	
		Session("SCroll") = "Yes"
	Else
		Session("SCroll") = "No"
	End If
	
	'Response.Write Session("CostCentreID")
	
 %>
<html>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<LINK rel="stylesheet" type="text/css" href="../BERTStyle.css">
<title>Isidore</title>
</HEAD>
<frameset FrameBorder="0" COLS="*%" ROWS="5%,95%,*%">
	<!--<frame NAME="Header1" SCROLLING="No" Border="0" SRC="HeaderFrameset1.asp">
	<frame NAME="Body1" SCROLLING="Yes" Border="0" SRC="<%=Session("CurrentPage")%>">-->
       
	<frame NAME="Header1" SCROLLING="No" Border="0" SRC="Header2.asp">
	<frame NAME="Body1" SCROLLING="<%=Session("Scroll")%>" Border="0" SRC="<%=Session("CurrentPage")%>">
</frameset>
<noframes>
<body>
<b>Isidore IT<p>
Sorry your browser doesn't support frames.
</b></body>
</noframes>
</html>
