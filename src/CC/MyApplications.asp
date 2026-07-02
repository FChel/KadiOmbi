<%@ Language=VBScript %>
<% Option Explicit %>
<!-- #Include file=../ADOVBS.inc -->

<%
 
'Description:	Student Screen
'Author:		Andrew Bull - Isidore Limited (www.isidore.com - 07969 589413)
'Date:			August 2004

'Declare default variables

Dim objCon
Dim objCmd
Dim objRS
Dim objRS1
Dim strScreen
Dim strSelected
Dim x 
Dim strMessage
Dim strColour

'Set Database objects

Set objCon = Server.CreateObject("ADODB.Connection")
Set objCmd = Server.CreateObject("ADODB.Command")
Set objRS = Server.CreateObject("ADODB.Recordset")
Set objRS1 = Server.CreateObject("ADODB.Recordset")

'Open database connection

objCon.Open Session("DBConnection")



'1. Declare screen specific variables naming convention = variable type prefix + database field name

Dim lngBudgetID
Dim lngCostCentreID
Dim lngStatusID
Dim strBudgetName
Dim intFinancialYearID
Dim intDefaultVersionID
Dim strBalanceSheet
Dim lngCashFlowGLCode
Dim lngBadDebtGLCode
Dim lngPrePaymentGLCode
Dim lngInvestmentGLCode
Dim lngLoanGLCode
Dim lngAPGLCode
Dim lngARGLCode
Dim intVariancePercentage
Dim strComments
Dim strActive

Dim dblRevenue
Dim dblExpenditure
Dim dblNetTotal

Dim dblApprovedRev
Dim dblApprovedExp
Dim dblApprovedNet

Dim dblTotalRev
Dim dblTotalExp
Dim dblTotalNet

Dim dblOpexRev
Dim dblOpexExp
Dim dblOpexNet

Dim dblCapxRev
Dim dblCapxExp
Dim dblCapxNet
Dim dblEstimate
Dim dblEstimateTotal

'Declare and set default arrays

Dim arrYesNo(2)
	
	arrYesNo(1) = "Y"
	arrYesNo(2) = "N"
	
	'3. Capture Querystring variables
	
	If Not IsEmpty(Request.QueryString("CostCentreID")) Then
		
		lngCostCentreID = Request.QueryString("CostCentreID")
						
	End If
	
	If Not IsEmpty(Request.QueryString("StatusID")) Then
		
		lngStatusID = Request.QueryString("StatusID")
					
	End If	
	
    
Dim arrStatus(7)

arrStatus(0) = "<IMG SRC='../images/delete.png'"
arrStatus(1) = "<IMG SRC='../images/open.png'"
arrStatus(2) = "<IMG SRC='../images/ready.gif'"
arrStatus(3) = "<IMG SRC='../images/cross.png'" 
arrStatus(4) = "<IMG SRC='../images/tick.png'"	
arrStatus(5) = "<IMG SRC='../images/Closed.png'"
arrStatus(6) = "<IMG SRC='../images/bug.png'"   
arrStatus(7) = "<IMG SRC='../images/wrench.png'"
 			
	'Session("EmployeeID") = Session("Logon")
			
%>

<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="../CAPSStyle.css">

 <!-- Bootstrap Core CSS -->
    <!--<link href="../css/bootstrap.min.css" rel="stylesheet">-->
	<link href="../vendor/font-awesome/css/font-awesome.min.css" rel="stylesheet" type="text/css">
	
	 <!-- jQuery -->
    <script src="../js/jquery.js"></script>

    <!-- Bootstrap Core JavaScript -->
    <script src="../js/bootstrap.min.js"></script>



</head>

<body>

<!-- Modal -->
<div class="modal hide fade" id="myModal">
  <div class="modal-header">
    <a class="close" data-dismiss="modal">×</a>
    <h3>Modal header</h3>
  </div>
  <div class="modal-body">
    <p>One fine body…</p>
  </div>
  <div class="modal-footer">
    <a href="#" class="btn">Close</a>
    <a href="#" class="btn btn-primary">Save changes</a>
  </div>
</div>

<form action="CostCentreStatus.asp?Action=Save" method="POST" id="frm" name="frm">

<div class="card mb-3">
       
        <div class="card-body">
          <div class="table-responsive">

<table class="table table-bordered table-hover" id="dataTable" BORDER="1" CELLSPACING="1" CELLPADDING="1">
 <THEAD>
	<tr><th colspan="8" Style="Height:25px;font-size:14px;font-weight:bold; text-align:center;"><%=Session("EmployeeID") & " - " & Session("UserName") %></td></tr>
	<tr>
		<th Style="Height:25px;font-size:11px;font-weight:bold;background-color:#2394F2;color:white;text-align:center;" Width="10%">Card Type</th>
		<th Style="Height:25px;font-size:11px;font-weight:bold;background-color:#2394F2;color:white;text-align:center;" Width="5%">ID</th>
		<th Style="Height:25px;font-size:11px;font-weight:bold;background-color:#2394F2;color:white;text-align:center;" Width="20%">Applicant</th>
		<th Style="Height:25px;font-size:11px;font-weight:bold;background-color:#2394F2;color:white;text-align:center;" Width="15%">Received Date</th>
        <th Style="Height:25px;font-size:11px;font-weight:bold;background-color:#2394F2;color:white;text-align:center;" Width="10%">Status</th>
		<th Style="Height:25px;font-size:11px;font-weight:bold;background-color:#2394F2;color:white;text-align:center;" Width="10%">Approved Date</th>
		<th Style="Height:25px;font-size:11px;font-weight:bold;background-color:#2394F2;color:white;text-align:center;" Width="10%">View Progress</th>
		<th Style="Height:25px;font-size:11px;font-weight:bold;background-color:#2394F2;color:white;text-align:center;" Width="20%">Actions</th>
	</tr>
	</THEAD>
<%

Dim y

Dim strCardType, strApplicant, strStatus, dteReceived, dteExiryDate
Dim dblTotal
Dim strStatusHold
Dim strReceivedFormat
Dim strLink
Dim strCardType2, strNameOnCard, strCardNo, strActions, strCreditLimit
Dim strCreditAction, strAction, strReview, lngID
Dim strClass

'Open the master recordset
'objRS.Open "SELECT * FROM qryCardsCardType WHERE EmployeeID = " & Session("EmployeeID") & "",objCon

'objRS.Open "SELECT [CardType] As CardType1 FROM qryCardsCardType left outer join (SELECT * FROM qryCardsCardType WHERE EmployeeID = " & Session("EmployeeID") & ") AS A On qryCardsCardType.CardTypeID = A.CardTypeID ",objCon

'response.write "select * from qryApplicationCardType left outer join (select * from qryApplicationCardType WHERE EmployeeID = " & Session("EmployeeID") & ") AS A On qryApplicationCardType.CardTypeID = A.CardTypeID"
'objRS.Open "select * from qryApplicationCardType left outer join (select * from qryApplicationCardType WHERE EmployeeID = " & Session("EmployeeID") & ") AS A On qryApplicationCardType.CardTypeID = A.CardTypeID",objCon

objRS.Open "select * from tblCardType left outer join (select * from qryApplicationCardType WHERE EmployeeID = '" & Session("EmployeeID") & "') AS A On tblCardType.CardTypeID = A.CardTypeID",objCon,0,1


	strStatusHold = ""
	
	Do Until objRS.EOF
		'Insert Row Heading
		'Response.Write "<TR><TD >&nbsp;" & objRS("CardType") & "</TD>"	

		If strStatusHold = objRS(1) Then
		
		Else
			Response.Write "<TR><TH colspan=""8"" style=""text-align:center; font-size:12px;font-weight:bold;background-color:#86c5f9;color:white;"">" & objRS(1) & "</TH></TD>"
		End If
			
		If objRS.EOF Then
			'Response.Write "<TD style=""text-align:right;"">0</TD><TD style=""text-align:right;"">0</TD><TD style=""text-align:right;"">0</TD>"
		Else
			
			If IsNull(objRS("ApplicationID")) Then
				lngID = 0
			Else
				lngID = objRS("ApplicationID")
				
			End If
			
			If IsNull(objRS("CardType1")) Then
				strCardType = ""
			Else
				strCardType = objRS("CardType1")
				strCardType = objRS(1)
			End If
			strCardType2 = objRS(1)
			
			If trim(strCardType2) = "Diners DTC" Then
				strCardType = "<img src=""img/Diners2.png"" height=""40px"" width=""50px""> "'& strCardType
			ElseIf strCardType2 = "ANZ DPC" Then
				strCardType = "<img src=""img/ANZ2.png"" height=""40px"" width=""70px""> " '& strCardType
			Else
				strCardType = "<img src=""img/Mastercard2.png"" height=""40px"" width=""50px""> " '& strCardType
			End If
			
			If IsNull(objRS("DateReceived")) Then
				dteExiryDate = ""
			Else
				dteExiryDate = FormatDateTime(objRS("DateReceived"),vbShortDate)
				If dteExiryDate < now() -10 then
					strReceivedFormat = "color:red; font-weight:bold;"
				Else
					strReceivedFormat = "color:black;"
				End If
				
			End If
			
			If IsNull(objRS("FirstName")) Then
				strNameOnCard = "<span style=""color:#95a5a6;""><i>No Applications</i></span>"
			Else
				strNameOnCard = objRS("FirstName") & " " & objRS("Surname")
				
			'	strLink = "<A target=""_parent"" href=""Cards.asp?CardID=" & objRS("CardID") & "&EmployeeID=" & objRS("EmployeeID") & " "" Style=""" & strReceivedFormat & "font-size:12px;"">"
			End If
			
			
			If isnull(objRS("Status")) Then
				strStatus = ""
			Else
				strStatus = objRS("Status")
			End If
			
			'If isnull(objRS("CreditLimit")) Then
				strCreditLimit = ""
			'Else
		'		strCreditLimit = objRS("CreditLimit")
			'	strCreditAction = "<button type=""button"" class=""btn btn-secondary"" onclick=""self.location='Cards.asp?CardID=" & objrs("CardID") & "'"";><i class=""fa fa-key""></i> Change Limit</button>"
			'End If
			
			If isnull(objRS("Status")) Then
				strReview = ""
			Else
				strReview = "<button type=""button"" class=""btn btn-info btn-xs"" onclick=""parent.location='HomeCC2.asp?ApplicationID=" & objrs("ApplicationID") & "&HomeCC=Application'""><i class=""fa fa-cogs""></i> Check Progress</button>"
			End If
			
			Select Case objRS("Status")
		
				Case  "Active"
					strAction = "<button type=""button"" class=""btn btn-danger btn-xs"" onclick=""self.location='MyCards.asp?Action=Release&ApplicationID=" & objrs("CardID") & "'"";><i class=""fa fa-key""></i> Cancel</button>"
					strAction = strAction & " <button type=""button"" class=""btn btn-primary btn-xs"" onclick=""self.location='MyCards.asp?Action=Reject&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-cogs""></i> Reject</button>"
				
				Case "Cancelled"

					strAction = "<button type=""button"" class=""btn btn-secondary btn-xs"" onclick=""parent.location='ApplicationsEmployee.asp?ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-key""></i> View Application</button>"
					strClass ="class=""list-group-item-danger"""
				
				Case "Submitted"
				
				Case Else
					strAction = "<button type=""button"" class=""btn btn-primary btn-xs"" onclick=""parent.location='ApplicationsEmployee.asp?EmployeeID=" & session("Logon") & "'""><i class=""fa fa-credit-card""></i> Apply</button>"
				
			End Select

			Response.Write "<TR><TD style=""text-align:center;font-size:14px; "">" & strLink & "" & strCardType & "</A></TD><TD style=""text-align:center; font-size:14px;"">" & strLink & "" & lngID & "</A></TD><TD style=""text-align:center;font-size:14px; "">" & strLink & "" & strNameOnCard & "</A></TD>" & _
				"<TD style=""text-align:center;font-size:14px; " & strReceivedFormat & """>" & strLink & " " & dteExiryDate & "</A></TD><TD style=""text-align:center;font-size:14px; "">" & strLink & "" & strStatus & "</A></TD>" & _ 
				"<TD style=""text-align:center;font-size:14px;"">" & strLink & "" & strCreditLimit & " " & strCreditAction & "</A></TD><TD style=""text-align:center;font-size:14px;"">" & strReview & "</A></TD>" & _
				"<TD style=""text-align:center;font-size:14px; " & strReceivedFormat & """>" & strAction & "</A></TD></TR>"
			
		
		End If
		
		strStatusHold = objRS("Status")
		
		objRS.Movenext
	Loop
	objRS.Close
	
	
%>
</table>
</DIV>
</DIV>
</DIV>
</DIV>

<div class="alert alert-success alert-dismissible">
    <button type="button" class="close" data-dismiss="alert">×</button>
    <strong>Your Applications are above!</strong> Click on the APPLY button next to the card you would like to apply for.
  </div>
  
</body>

</html>
<%

Set objRS = Nothing
Set objCon = Nothing


%>
