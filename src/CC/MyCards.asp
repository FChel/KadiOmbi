<%@ Language=VBScript %>
<% Option Explicit %>
<!-- #Include file=../ADOVBS.inc -->

<%
 
'Description:	User's Cards Screen
'Author:		Michael Giacomin
'Date:			January 2020

'Declare default variables

If IsEmpty(Session("UserID")) Then Response.Redirect("../Default.asp?State=Expired")

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
  

</head>

<body>
<form action="CostCentreStatus.asp?Action=Save" method="POST" id="frm" name="frm">

<!-- Modal -->
<div class="modal fade" id="exampleModalCenter" tabindex="-1" role="dialog" aria-labelledby="exampleModalCenterTitle" aria-hidden="true">
  <div class="modal-dialog modal-dialog-centered" role="document">
    <div class="modal-content">
      <div class="modal-header">
        <h5 class="modal-title" id="exampleModalLongTitle">Credit Application Declaration Form</h5>
        <button type="button" class="close" data-dismiss="modal" aria-label="Close">
          <span aria-hidden="true">&times;</span>
        </button>
      </div>
      <div class="modal-body">
        <label>I DECLARE that all the application details are correct.</label><br>
            <label for="name">First Name:</label>
            <input type="text" name="FirstName" id="FirstName" class="form-control input-md">
			<label for="name">Last Name:</label>
            <input type="text" name="LastName" id="LastName" class="form-control input-md">
			<label for="email">Email:</label>
            <input type="email" name="email" id="email" class="form-control input-md">
            <label for="phone">Phone:</label>
            <input type="text" name="phone" id="phone" class="form-control input-md">
			<label for="phone">Group:</label>
            
			<select class="form-control" id="Group">
			  <option>Please select your group...</option>
			  <option>Army</option>
			  <option>Navy</option>
			  <option>Air Force</option>
			  <option>CIOG</option>
			  <option>CFOG</option>
			  <option>DSTO</option>
			  <option>JOC</option>
			</select>
      </div>
      <div class="modal-footer">
        <button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
        <button type="button" class="btn btn-primary">Save changes</button>
      </div>
    </div>
  </div>
</div>


<div class="card mb-3">
       
        <div class="card-body">
          <div class="table-responsive">

<table class="table table-bordered table-hover" id="dataTable" BORDER="1" CELLSPACING="1" CELLPADDING="1">
 <THEAD>
	<tr><th colspan="7" Style="Height:25px;font-size:14px;font-weight:bold; text-align:center;"><%=Session("EmployeeID") & " - " & Session("UserName") %></td></tr>
	<tr>
		<th Style="Height:25px;font-size:11px;font-weight:bold;background-color:#2394F2;color:white;text-align:center;" Width="10%">Card Type</th>
		<th Style="Height:25px;font-size:11px;font-weight:bold;background-color:#2394F2;color:white;text-align:center;" Width="20%">Name On Card</th>
		<th Style="Height:25px;font-size:11px;font-weight:bold;background-color:#2394F2;color:white;text-align:center;" Width="20%">Card No</th>
        <th Style="Height:25px;font-size:11px;font-weight:bold;background-color:#2394F2;color:white;text-align:center;" Width="10%">Status</th>
		<th Style="Height:25px;font-size:11px;font-weight:bold;background-color:#2394F2;color:white;text-align:center;" Width="10%">Credit Limit</th>
		<th Style="Height:25px;font-size:11px;font-weight:bold;background-color:#2394F2;color:white;text-align:center;" Width="10%">Expiry Date</th>
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
Dim strCreditAction, strAction

'Open the master recordset
'objRS.Open "SELECT * FROM qryCardsCardType WHERE EmployeeID = " & Session("EmployeeID") & "",objCon

'objRS.Open "SELECT [CardType] As CardType1 FROM qryCardsCardType left outer join (SELECT * FROM qryCardsCardType WHERE EmployeeID = " & Session("EmployeeID") & ") AS A On qryCardsCardType.CardTypeID = A.CardTypeID ",objCon


'objRS.Open "select * from qryCardsCardType left outer join (select * from qryCardsCardType WHERE EmployeeID = " & Session("EmployeeID") & ") AS A On qryCardsCardType.CardTypeID = A.CardTypeID",objCon

objRS.Open "SELECT * FROM tblCardType left outer join (select * from qryCardsCardType WHERE EmployeeID = '" & Session("EmployeeID") & "') AS A On tblCardType.CardTypeID = A.CardTypeID",objCon,0,1

	strStatusHold = ""
	
	Do Until objRS.EOF
		'Insert Row Heading
		'Response.Write "<TR><TD >&nbsp;" & objRS("CardType") & "</TD>"	

		If strStatusHold = objRS(1) Then
		
		Else
			Response.Write "<TR><TH colspan=""7"" style=""text-align:center; font-size:12px;font-weight:bold;background-color:#86c5f9;color:white;"">" & objRS(1) & "</TH></TD>"
		End If
			
		If objRS.EOF Then
			'Response.Write "<TD style=""text-align:right;"">0</TD><TD style=""text-align:right;"">0</TD><TD style=""text-align:right;"">0</TD>"
		Else
			
			If IsNull(objRS("CardType1")) Then
				strCardType = ""
			Else
				strCardType = objRS("CardType1")
				strCardType = objRS(1)
			End If
			strCardType2 = objRS(1)
			
			If trim(strCardType2) = "Diners DTC" Then
				strCardType = "<img src=""img/Diners2.png"" height=""40px"" width=""50px"" title=""" & strCardType & """> "
			ElseIf strCardType2 = "ANZ DPC" Then
				strCardType = "<img src=""img/ANZ2.png"" height=""40px"" width=""70px""> " '& strCardType
			Else
				strCardType = "<img src=""img/Mastercard2.png"" height=""40px"" width=""50px""> " '& strCardType
			End If
			
			If IsNull(objRS("Expiry")) Then
				dteExiryDate = ""
			Else
				dteExiryDate = objRS("Expiry")
				If dteExiryDate < now() -10 then
					strReceivedFormat = "color:red; font-weight:bold;"
				Else
					strReceivedFormat = "color:black;"
				End If
				
			End If
			
			If IsNull(objRS("NameOnCard")) Then
				strNameOnCard = "<span style=""color:#95a5a6;""><i>No Card</i></span>"
			Else
				strNameOnCard = objRS("NameOnCard")
				
				strLink = "<A target=""_parent"" href=""Cards3.asp?CardID=" & objRS("CardID") & "&EmployeeID=" & objRS("EmployeeID") & " "" Style=""" & strReceivedFormat & "font-size:12px;"">"
			End If
			
			If IsNull(objRS("CardNumber")) Then
				strCardNo = ""
			Else
				strCardNo = objRS("CardNumber")
				If len(strCardNo)>8 Then strCardNo = mid(strCardNo,6,2) & "****" & right(strCardNo,4)
			End If
			
			If isnull(objRS("Status")) Then
				strStatus = ""
			Else
				strStatus = objRS("Status")
				If strStatus = "00" Then
					strStatus = "Active"
					strStatus = "<span class=""label label-success"" Title=""Active""><i class=""fa fa-check""></i></span>"
				Else
					strStatus = "Cancelled"
					strStatus = "<span class=""label label-danger"" Title=""Cancelled""><i class=""fa fa-times""></i></span>"
				End If
			End If
			
			If isnull(objRS("CreditLimit")) OR objRS("CreditLimit") = "" Then
				strCreditLimit = ""
				strCreditAction = ""
			Else
				If IsNumeric(objRS("CreditLimit")) Then
					strCreditLimit = FormatCurrency(objRS("CreditLimit")/100,0)
				Else
					strCreditLimit = ""
				End If
				
				strCreditAction = strCreditLimit & "<br><button type=""button"" class=""btn btn-secondary btn-xs"" onclick=""parent.location='Cards3.asp?CardID=" & objrs("CardID") & "'"";><i class=""fa fa-money""></i> Change Limit</button>"
			
			End If
			
			Select Case objRS("Status")
		
				Case  "00"
					strAction = "<button type=""button"" class=""btn btn-danger btn-xs"" onclick=""parent.location='MyCards.asp?Action=Release&CardID=" & objrs("CardID") & "'"";><i class=""fa fa-minus-circle""></i> Cancel</button>"
					'strAction = strAction & " <button type=""button"" class=""btn btn-danger"" onclick=""self.location='MyCards.asp?Action=Reject&CardID=" & objrs("CardID") & "'"";><i class=""fa fa-cogs""></i> Reject</button>"
					strAction = strAction & "<br><button type=""button"" class=""btn btn-secondary btn-xs"" onclick=""parent.location='Cards3.asp?CardID=" & objrs("CardID") & "'"";><i class=""fa fa-key""></i> View Card</button>"
				Case "01"

					strAction = "<button type=""button"" class=""btn btn-secondary btn-xs"" onclick=""parent.location='Cards3.asp?CardID=" & objrs("CardID") & "'"";><i class=""fa fa-key""></i> View Card</button>"
				
				Case Else
					strAction = "<button type=""button"" class=""btn btn-primary btn-xs"" onclick=""parent.location='ApplicationsEmployee.asp?EmployeeID=" & session("Logon") & "'""><i class=""fa fa-credit-card""></i> Apply</button>"
				
			End Select
		
		
			
			
				
			Response.Write "<TR><TD style=""text-align:center; "">" & strLink & "" & strCardType & "</A></TD><TD style=""text-align:center; "">" & strLink & "" & strNameOnCard & "</A></TD>" & _
				"<TD style=""text-align:center; "">" & strLink & "" & strCardNo & "</A></TD><TD style=""text-align:center; "">" & strLink & "" & strStatus & "</A></TD>" & _ 
				"<TD style=""text-align:center;"">" & strLink & "" & strCreditAction & "</A></TD><TD style=""text-align:center; " & strReceivedFormat & """>" & strLink & "" & dteExiryDate & "</A></TD>" & _
				"<TD style=""text-align:center; " & strReceivedFormat & """>" & strAction & "</A></TD></TR>"
			
		
		End If
		
		strStatusHold = objRS(1)
		
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
    <strong>Your Cards are above!</strong> Click on the APPLY button next to the card you would like to apply for.
  </div>
  
</body>

</html>
<%

Set objRS = Nothing
Set objCon = Nothing


%>
