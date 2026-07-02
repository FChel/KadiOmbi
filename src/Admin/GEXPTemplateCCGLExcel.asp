<%@ Language=VBScript %>
<%
Option Explicit

	Response.Expires = -1500
	
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment; filename=BBGEXP.xls" 
 
	
Dim objCon
Dim objRS
Dim objRS1
Dim arrHeadings(5)
Dim intFinYearPart1
Dim intFinYearPart2
Dim strBackColour
Dim dblDisplay
Dim dblActual
Dim dblBudget
Dim dblVariance
Dim dblVariancePercentage
Dim strForeColour
Dim intMode
Dim dblTotal
Dim dblTotal1
Dim dblOriginal
Dim dblVarianceTotal
Dim dblVariancePercentageTotal
Dim dblOriginalTotal
Dim dblVarianceTotal1
Dim dblVariancePercentageTotal1
Dim x
Dim strMessage
Dim strLevelName
Dim strSQL
Dim dblVar
Dim strCostCentreName
Dim strVersionName
Dim arrMonthName(12)
Dim strComment

 'Set Headings
    For x = 0 to 5
	
	    intFinYearPart1 = cint(Session("FinancialYear")) + (x - 1)
	    intFinYearPart1 = Right(intFinYearPart1,2)
	    intFinYearPart2 = cint(Session("FinancialYear")) + x
	    intFinYearPart2 = Right(intFinYearPart2,2)

	    'If the Financial Year starts in January then do not display the split/multiple years.
        If Session("FirstMonth") = "JAN" Then
            arrHeadings(x) = "20" & cstr(intFinYearPart1)
        Else
	        arrHeadings(x) = cstr(intFinYearPart1) & "/" & cstr(intFinYearPart2)
        End If

    Next

Set objCon = Server.CreateObject("ADODB.Connection")
Set objRS = Server.CreateObject("ADODB.Recordset")
Set objRS1 = Server.CreateObject("ADODB.Recordset")

objCon.Open Session("DBConnection")	

'Get Level1ID

If Not IsEmpty(Request.QueryString("Level1ID")) Then
	Session("Level1ID") = Request.QueryString("Level1ID")
End If

If Not IsEmpty(Request.QueryString("TransactionType")) Then
	Session("TransactionType") = Request.QueryString("TransactionType")
End If

'Get Level1 Name
objRS.Open "SELECT * FROM tblReportLayoutLevel1 WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND Level1ID = " & Session("Level1ID") & " AND ReportID = 1",objCon

   If Not objRS.EOF Then
        strLevelName = objRS("Level1Name")
   Else
        strLevelName = "Missing"
   End IF

objRS.Close

'Execute Action
If Request.QueryString("Action") = "Save" Then
   If Session("StatusID") = 1 Then
        Call Save_Data()
   Else
        strMessage = "Budget is closed, no changes can be made!"
   End If
End If 

    objRS.Open "Select CostCentreID,CostCentreName from tblCostcentre where CostCentreID = " & clng(Session("CostCentreID")) & " AND BudgetID = " & clng(Session("BudgetID")),objCon

        If Not objRS.EOF Then
            strCostCentreName = objRS("CostCentreID") & " (" & objRS("CostCentreName") & ")" 
        Else
            strCostCentreName = ""
        End If

    objRS.Close

    objRS.Open "SELECT * FROM tblVersion WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & "",objCon

        If not objRS.EOF Then
            strVersionName = objRS("VersionName")
        End If

    objRS.Close

    'call the procedure to create the Month Names
    Call GetMonthNames()
%>

<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</head>
<body>

<table WIDTH="1560px" BORDER="1" CELLSPACING="1" CELLPADDING="1">
<thead>
	<%Response.Write "<th>CostCentreID</th>" & _
	        "<th>GLCode</th>" & _
	        "<th Width=""80px"">CY</th>" & _
                "<th Width=""80px"">" & arrMonthName(1) & "</th><th Width=""80px"">" & arrMonthName(2) & "</th>" & _
                "<th Width=""80px"">" & arrMonthName(3) & "</th><th Width=""80px"">" & arrMonthName(4) & "</th>" & _
                "<th Width=""80px"">" & arrMonthName(5) & "</th><th Width=""80px"">" & arrMonthName(6) & "</th>" & _
                "<th Width=""80px"">" & arrMonthName(7) & "</th><th Width=""80px"">" & arrMonthName(8) & "</th>" & _
                "<th Width=""80px"">" & arrMonthName(9) & "</th><th Width=""80px"">" & arrMonthName(10) & "</th>" & _
                "<th Width=""80px"">" & arrMonthName(11) & "</th><th Width=""80px"">" & arrMonthName(12) & "</th>" & _
                "<TH Width=""80px"">OY1</TH><TH Width=""80px"">OY2</TH><TH Width=""80px"">OY3</TH>" & _
	        "<th bgcolor=""CCCCCC"">GLCodeName</th>" & _
	        "<th bgcolor=""CCCCCC"">CostCentreName</th>" & _
	        "</thead>"
	
	'Call ther procedure to fill in the data for this client
	Call DisplayData() 
	 
	 %>
	 
	
</table>
</body>
</html>
<%

Public Sub DisplayData()
'Procedure to write all GLCodes by CostCentre for the current BudgetID
Dim intCostCentreID

objRS.Open "SELECT * FROM qryCostCentreGLCodes WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID"),objCon
'To add in the Budget data for months use the query below
'objRS.Open "SELECT * FROM qryCostCentreGLCodesBudgetData WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID"),objCon

        If not objRS.EOF Then
            intCostCentreID = objRS("CostCentreID")
        End If
        
		Do until objRS.EOF

            Response.Write "<TR><TD>" & objRS("CostCentreID") & "</TD><TD>" & objRS("GLCode") & "</TD><TD>0</TD><TD>0</TD><TD>0</TD><TD>0</TD><TD>0</TD><TD>0</TD><TD>0</TD>" & _
                            "<TD>0</TD><TD>0</TD><TD>0</TD><TD>0</TD><TD>0</TD><TD>0</TD><TD>0</TD><TD>0</TD><TD>0</TD><TD>" & objRS("GLCodeName") & "</TD><TD>" & objRS("CostCentreName") & "</TD></TR>"
				
			objRS.Movenext

		Loop

	objRS.Close

End Sub

Public Sub GetMonthNames()
'This is a procedure to get the order of Month names to be used as titles for Month Columns
Dim intFirstMonth

    'set the First Month name to an integer
    intFirstMonth = Month("21-" & Session("FirstMonth") & "-2012")
    'intFirstMonth = intFirstMonth -1
    arrMonthName(0) = intFirstMonth
    For x = 1 to 12
    
        arrMonthName(x) = Left(MonthName(intFirstMonth + x - 1),3)'intFirstMonth + x
  '      arrMonthName(x) = intFirstMonth'MonthName(intFirstMonth)
        
        'Once the count goes over 12 then go back to 1 to fill the remaining months
        If intFirstMonth + x - 1 > 11 Then 
            If intFirstMonth > 6 Then
                intFirstMonth = 2 - intFirstMonth
            Else
                intFirstMonth = (x - 1) * - 1
            End If
        End If
        
  '      intFirstMonth = x
        
    Next
    
End Sub
%>
