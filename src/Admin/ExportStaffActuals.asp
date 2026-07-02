<%@ Language=VBScript %>
<!-- #Include file=../ADOVBS.inc -->
<%
'Description:	Staff Data Actuals of All Staff Actuals Data to Excel
'Author:		MG
'Date:			April 2016

	Response.Expires = -1500
	
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment; filename=BBStaffActuals.xls" 
	
    'If the session has expired then send the user back to the Default/login page
	If IsEmpty(Session("DBConnection")) Then
		Response.Redirect "../Default.asp?State=Expired"
	End If
	
'Instantiate Common Page Variables.
Dim objCon
Dim objRS
Dim arrMonthName(12)
Dim strClient
 
'Open Database Connection
Set objCon = Server.CreateObject("ADODB.Connection")
Set objRS = Server.CreateObject("ADODB.Recordset")

objCon.Open Session("DBConnection")

	'Call the procedure to create the Month Names
    Call GetMonthNames()
    Call GetClientName()
 %>

<html>
<head>

</head>
<body>

<% 
    'If there is a Version New selecetd then check to see if there is any existing data that will be deleted if the rollover continues.
    If IsNull(Session("VersionID")) or Session("VersionID") = "" Then
    
    Else

        objRS.Open "SELECT * FROM qryStaffDataVersion WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND [Deleted] = 'N' Order By CostCentreID",objCon
		    If objRS.eof Then
		        Response.Write"<table WIDTH=""100%"" BORDER=""1"" CELLSPACING=""1"" CELLPADDING=""1"">" & _
                "<tr><th colspan=7 style=""text-align:left"">" & strClient & " - There is no data in the Version <B>" & strVersionName & "</B></th></tr>" 
		    Else
		    
		         Response.Write"<table WIDTH=""100%"" BORDER=""1"" CELLSPACING=""1"" CELLPADDING=""1"">" & _
                "<tr><th colspan=""18"" style=""text-align:left"">" & strClient & " - BizBudg - All existing Staff Data for Budget - <B>" &  Session("BudgetName") & "</B> and Version - <B>" & Session("VersionName") & "</b></th></tr>" & _
                "<tr><td colspan=""18"" style=""text-align:left; color:red;font-size:20px""><B>To Upload back into BizBudg after changes DELETE this entire row (and the one above) so that the headers (below row) are the first row. DO NOT Remove any columns!</B></td></tr><tr>" & _
		        "<th>Budget</th>" & _	
		        "<th>Version</th>" & _
	 	        "<th>Cost Centre</th>" & _	
	 	        "<th>Employee/Level</th>" & _
		        "<th Width=""80px"">" & arrMonthName(1) & "</th><th Width=""80px"">" & arrMonthName(2) & "</th>" & _
                "<th Width=""80px"">" & arrMonthName(3) & "</th><th Width=""80px"">" & arrMonthName(4) & "</th>" & _
                "<th Width=""80px"">" & arrMonthName(5) & "</th><th Width=""80px"">" & arrMonthName(6) & "</th>" & _
                "<th Width=""80px"">" & arrMonthName(7) & "</th><th Width=""80px"">" & arrMonthName(8) & "</th>" & _
                "<th Width=""80px"">" & arrMonthName(9) & "</th><th Width=""80px"">" & arrMonthName(10) & "</th>" & _
                "<th Width=""80px"">" & arrMonthName(11) & "</th><th Width=""80px"">" & arrMonthName(12) & "</th>" & _
		        "<th style=""background-color:#EDECEC;"">CostCentreID</th>" & _
		        "<th style=""background-color:#EDECEC;"">EmployeeID</th></tr>"
		        
		    End If
		    
		    Do until objRS.eof
		        'If objRS("GLCode") <> 0 Then
			    Response.Write "<TR><TD>&nbsp;" & objRS("BudgetName") & "</A></B></TD>" & _
			                    "<TD style=""text-align:center"">" & objRS("VersionName") & "</TD><TD style=""text-align:center"">" & objRS("CostCentreName") & "</TD><TD style=""text-align:center"">" & objRS("StaffClassificationDesc") & "</TD>" & _
			                    "<TD style=""text-align:center"">" & objRS("AM1") & "</TD><TD style=""text-align:center"">" & objRS("AM2") & "</TD><TD style=""text-align:center"">" & objRS("AM3") & "</TD>" & _
			                    "<TD style=""text-align:center"">" & objRS("AM4") & "</TD><TD style=""text-align:center"">" & objRS("AM5") & "</TD><TD style=""text-align:center"">" & objRS("AM6") & "</TD>" & _
			                    "<TD style=""text-align:center"">" & objRS("AM7") & "</TD><TD style=""text-align:center"">" & objRS("AM8") & "</TD><TD style=""text-align:center"">" & objRS("AM9") & "</TD>" & _
			                    "<TD style=""text-align:center"">" & objRS("AM10") & "</TD><TD style=""text-align:center"">" & objRS("AM11") & "</TD><TD style=""text-align:center"">" & objRS("AM12") & "</TD>" & _
			                    "<TD>" & objRS("CostCentreID") & "</TD><TD>" & objRS("EmployeeID") & "</TD></TR>"
    			'End If
    			
			    objRS.movenext
		    Loop
    			
	    objRS.Close

        Response.Write "</table>"
    End If


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


Public Sub GetClientName()

'Get Client Details
    objRS.Open "SELECT * FROM tblClient WHERE ClientID = " & Session("ClientID") & "",objCon
    
        If not objRS.EOF Then
            strClient = objRS("ClientName")
            Session("FirstMonth") = objRS("FirstMonth")
        Else
            strClient = ""
            Session("FirstMonth") = "JAN"
        End IF
    
    objRS.Close
    
End Sub

Set objRS = Nothing
Set objCon = Nothing

 %>
 </body>
</html>

