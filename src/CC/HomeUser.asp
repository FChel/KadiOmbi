
<!-- #Include file=CAPSHeader.asp -->
<!-- #Include file=ADOVBS.inc -->
<%

Response.ContentType = "text/html" 
Response.CodePage = 65001
Response.CharSet = "UTF-8"

If IsEmpty(Session("UserID")) Then Response.Redirect("../Timeout.asp?State=Expired")


'Description:	Create and view applications
'Author:		MG
'Date:			January 2020

	Response.Expires = -1500	

Dim objCon
Dim objCmd
Dim objRS
Dim objRS1
Dim strSelected
Dim x 
Dim strMessage
Dim strColour

'Set Database objects

Set objCon = Server.CreateObject("ADODB.Connection")
Set objCmd = Server.CreateObject("ADODB.Command")
Set objRS = Server.CreateObject("ADODB.Recordset")

'Open database connection
objCon.Open Session("DBConnection")
%>
<script>
function OpenSs(cb) {

	//var id = cb.getAttribute('CardTypeSelect');
	//var id = document.getElementByName("CardTypeSelect").value;
	//alert(id);
	//var e = document.getElementById("CardTypeSelect");
	//var result = e.options[e.selectedIndex].value;
	
	//document.getElementById('CardType').value=result;
	alert('asa');
}

</script>
<!-- Modal -->
<div class="modal fade" id="ModApp" tabindex="-1" role="dialog" aria-labelledby="exampleModalCenterTitle" aria-hidden="true">
  <div class="modal-dialog modal-dialog-centered" role="document">
    <div class="modal-content">
      <div class="modal-header">
        <h5 class="modal-title" id="exampleModalLongTitle">CAPS Contact</h5>
        <button type="button" class="close" data-dismiss="modal" aria-label="Close">
          <span aria-hidden="true">&times;</span>
        </button>
      </div>
      <div class="modal-body">
        <label>Contact Card Providers</label><br>

			<h6 class="modal-title" id="exampleModalLongTitle">Diners - 1800 123 123</h6>
			<h6 class="modal-title" id="exampleModalLongTitle">ANZ - 132313</h6>
      </div>
      <div class="modal-footer">
        <button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
      </div>
    </div>
  </div>
</div>


    <main class="main py-5">
      <div class="container">
        <div class="row">
          <div class="col-md-9">
            <div class="jumbotron caps-intro mb-5">
              <div class="content">
                <h1>Welcome to CAPS</h1>
                <p>
                  CAPS is the Department of Defence Credit Card Application Processing and Management System.  Apply for a card, view or edit existing cards.
                </p>
                <a href="#" class="btn btn-primary">Get started</a>
              </div>
            </div>
            <div class="my-information">
              <ul class="nav nav-tabs" id="myFiTab" role="tablist">
                <li class="nav-item" role="presentation">
                  <a class="nav-link active" id="my-cards-tab" data-toggle="tab" href="#my-cards" role="tab" aria-controls="my-cards" aria-selected="true">My Cards</a>
                </li>
                <li class="nav-item" role="presentation">
                  <a class="nav-link" id="my-applications-tab" data-toggle="tab" href="#my-applications" role="tab" aria-controls="my-applications" aria-selected="false">My Applications</a>
                </li>
              </ul>
              <div class="tab-content" id="myFiTabContent">
                <div class="tab-pane fade show active" id="my-cards" role="tabpanel" aria-labelledby="my-cards-tab">
				<%
				Call LoadCards()
				
				%>
                  
                </div>
                <div class="tab-pane fade" id="my-applications" role="tabpanel" aria-labelledby="my-applications-tab">
                 <%
				Call LoadApplications()
				
				%>
				 
                </div>
              </div>
            </div>
          </div>
          <div class="col-md-3 sidebar">
            <div class="panel panel-shadow mb-3">
              <div class="panel-header">
                <h4>Getting started</h4>
                <span class="panel-subheader">Getting started with CAPS</span>
              </div>
              <div class="panel-content">
                <a href="img/ApplyForACard.pdf" target="_new" class="block-link">
                  <i class="fa fa-comments"></i><span class="content">Frequently Asked Questions</span>
                </a>
                <a href="HelpFile.html" target="_new" class="block-link">
                  <i class="fa fa-paper-plane"></i><span class="content">How to guide</span>
                </a>
              </div>
            </div>
            <div class="panel panel-shadow mb-3">
              <div class="panel-header">
                <h4>Shortcuts</h4>
                <span class="panel-subheader">Shortcuts to common providers</span>
              </div>
              <div class="panel-content">
                <a href="#" class="block-link" data-toggle="modal" data-target="#ModApp" >
                  <i class="fa fa-credit-card"></i>
                  <span class="content">Contact Diners</span>
                </a>
                <a href="#" class="block-link" data-toggle="modal" data-target="#ModApp">
                  <i class="fa fa-credit-card"></i>
                  <span class="content">Contact ANZ</span>
                </a>
              </div>
            </div>
          </div>
        </div>
      </div>
    </main>

	<script src="js/jquery.js"></script>
    <script src="bootstrap/js/bootstrap.bundle.min.js"></script>
    <script>
      jQuery(function ($) {
        $('[data-toggle="popover"]').popover();
      });
    </script>
	
	
<!-- #Include file=CAPSFooter.asp -->
  </body>
</html>

<%

Public Sub LoadCards()

Dim y

Dim strCardType, strApplicant, strStatus, dteReceived, dteExiryDate
Dim dblTotal
Dim strStatusHold
Dim strReceivedFormat
Dim strLink
Dim strCardType2, strNameOnCard, strCardNo, strActions, strCreditLimit
Dim strCreditAction, strAction
Dim strSQL

'Open the master recordset
objRS.Open "SELECT * FROM tblCAPSCardType WITH(NOLOCK) left outer join (select * from qryCardsCardType WITH(NOLOCK) WHERE EmployeeID = '" & Session("EmployeeID") & "') AS A On tblCAPSCardType.CardTypeID = A.CardTypeID",objCon,0,1

	strStatusHold = ""
	
	Do Until objRS.EOF
		'Insert Row Heading
		'Response.Write "<TR><TD >&nbsp;" & objRS("CardType") & "</TD>"	

		If strStatusHold = objRS(1) Then
		
		Else
			'Response.Write "<TR><TH colspan=""7"" style=""text-align:center; font-size:12px;font-weight:bold;background-color:#86c5f9;color:white;"">" & objRS(1) & "</TH></TD>"
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
				strCardType = "<img src=""../images/diners2.png"" height=""40px"" width=""50px"" title=""" & strCardType & """> "
			ElseIf strCardType2 = "ANZ DPC" Then
				strCardType = "<img src=""../images/ANZ.png"" height=""30px"" width=""80px""> " '& strCardType
			ElseIf strCardType2 = "Diners MasterCard" Then
				strCardType = "<img src=""../images/mc.png"" height=""40px"" width=""50px""> " '& strCardType
			Else
				strCardType = "<img src=""../images/high-limit.png"" height=""40px"" width=""50px""> "
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
				
				'strLink = "<A target=""_parent"" href=""Cards3.asp?CardID=" & objRS("CardID") & "&EmployeeID=" & objRS("EmployeeID") & " "" Style=""" & strReceivedFormat & "font-size:12px;"">"
				strLink = "<a target=""_parent"" href=""Cards3.asp?CardID=" & objRS("CardID") & "&EmployeeID=" & objRS("EmployeeID") & " "" class=""btn btn-link"">"
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
					'strAction = "<button type=""button"" class=""btn btn-danger btn-xs"" onclick=""parent.location='MyCards.asp?Action=Release&CardID=" & objrs("CardID") & "'"";><i class=""fa fa-minus-circle""></i> Cancel</button>"
					'strAction = strAction & " <button type=""button"" class=""btn btn-danger"" onclick=""self.location='MyCards.asp?Action=Reject&CardID=" & objrs("CardID") & "'"";><i class=""fa fa-cogs""></i> Reject</button>"
					'strAction = strAction & "<br><button type=""button"" class=""btn btn-secondary btn-xs"" onclick=""parent.location='Cards3.asp?CardID=" & objrs("CardID") & "'"";><i class=""fa fa-key""></i> View Card</button>"
					strAction = "<span class=""badge badge-pill badge-success"">Active</span>"
				Case "01"

					'strAction = "<button type=""button"" class=""btn btn-secondary btn-xs"" onclick=""parent.location='Cards3.asp?CardID=" & objrs("CardID") & "'"";><i class=""fa fa-key""></i> View Card</button>"
					strAction = "<span class=""badge badge-pill badge-danger"">Cancelled</span>"
				Case Else
					'strAction = "<button type=""button"" class=""btn btn-primary btn-xs"" onclick=""parent.location='ApplicationsEmployee.asp?EmployeeID=" & session("Logon") & "'""><i class=""fa fa-credit-card""></i> Apply</button>"
					strAction = ""
			End Select
		
		
			If strAction = "" Then
				Response.Write "<div class=""panel panel-shadow mb-3""><div class=""panel-content row"">" & _
					"<div class=""col-md-1 text-center my-auto"">" & strLink & "" & strCardType & "</div>" & _
					"<div class=""col-md-4""><h6>DPC / ANZ</h6><p>You don't have DPC / ANZ card</p></div>" & _
					"<div class=""col-md-7 text-right my-auto""><a href=""ApplicationsSubmit.asp"" class=""btn btn-primary"">Apply <i class=""fa fa-arrow-right""></i></a></div></div></div>"
			Else	  
				Response.Write "<div class=""panel panel-shadow mb-3""><div class=""panel-content row"">" & _
					"<div class=""col-md-1 text-center my-auto"">" & strLink & "" & strCardType & "</div>" & _
					"<div class=""col-md-4""><h6>" & strCardType2 & " " & strAction & _
					"</h6><p>Number: <strong>" & strCardNo & "</strong></p></div>" & _
					"<div class=""col-md-4""><p>Limit: <strong>" & strCreditLimit & "</strong></p><p>Expiry: <strong>" & dteExiryDate & "</strong></p></div>" & _
					"<div class=""col-md-3 text-right my-auto"">" & strLink & "View Details <i class=""fa fa-arrow-right""></i></a></div></div></div>"
			End If	  
				  
			'Response.Write "<TR><TD style=""text-align:center; "">" & strLink & "" & strCardType & "</A></TD><TD style=""text-align:center; "">" & strLink & "" & strNameOnCard & "</A></TD>" & _
			'	"<TD style=""text-align:center; "">" & strLink & "" & strCardNo & "</A></TD><TD style=""text-align:center; "">" & strLink & "" & strStatus & "</A></TD>" & _ 
			'	"<TD style=""text-align:center;"">" & strLink & "" & strCreditAction & "</A></TD><TD style=""text-align:center; " & strReceivedFormat & """>" & strLink & "" & dteExiryDate & "</A></TD>" & _
			'	"<TD style=""text-align:center; " & strReceivedFormat & """>" & strAction & "</A></TD></TR>"
			
		
		End If
		
		strStatusHold = objRS(1)
		
		objRS.Movenext
	Loop
	objRS.Close
	

End Sub


Public Sub LoadApplications()
Dim y

Dim dteDateSubmitted
Dim strCardType
Dim strCardType2
Dim strNameOnCard
Dim strCreditLimit
Dim strLink
Dim strSQL
Dim strReceivedFormat
Dim strAction
Dim strCardNo

If Session("EmployeeID") = "" OR ISNull(Session("EmployeeID")) Then
	strSQL = "SELECT * FROM qryCAPSApplications WITH(NOLOCK) WHERE EmployeeID = '" & Session("EmployeeID") & "'"
Else
	strSQL = "SELECT * FROM qryCAPSApplications WITH(NOLOCK) WHERE EmployeeID = '" & Session("EmployeeID") & "'"
End If

objRS.Open strSQL,objCon

    y = 0
	
	'Write a message in the list if there are no applications
	If objRS.EOF Then
		Response.Write "<div class=""panel panel-light""><div class=""panel-content text-center"">" & _
				"<h4 class=""my-3"">No pending Applications</h4>" & _ 
                "<a href=""ApplicationsSubmit.asp"" class=""btn btn-primary"">New Application <i class=""fa fa-arrow-right""></i></a></div></div>"
	End If
    	
    Do until objRS.EOF 
		
		If IsNull(objRS("CardTypeSub")) Then
			strCardType = ""
		Else
			strCardType = objRS("CardTypeSub")
			strCardType = objRS(1)
		End If
		strCardType2 = objRS(1)
		
		If trim(strCardType2) = "Diners" Then
			strCardType = "<img src=""../images/diners2.png"" height=""40px"" width=""50px"" title=""" & strCardType & """> "
		ElseIf strCardType2 = "ANZ" Then
			strCardType = "<img src=""../images/ANZ.png"" height=""30px"" width=""80px""> " '& strCardType
		ElseIf strCardType2 = "MasterCard" Then
			strCardType = "<img src=""../images/mc.png"" height=""40px"" width=""50px""> " '& strCardType
		Else
			strCardType = "<img src=""../images/high-limit.png"" height=""40px"" width=""50px""> "
		End If
		
		If IsNull(objRS("DateSubmitted")) Then
			dteDateSubmitted = ""
		Else
			dteDateSubmitted = objRS("DateSubmitted")
			If dteDateSubmitted < now() -10 then
				strReceivedFormat = "color:red; font-weight:bold;"
			Else
				strReceivedFormat = "color:black;"
			End If
			
		End If
		
		If IsNull(objRS("NameOnCard")) Then
			strNameOnCard = "<span style=""color:#95a5a6;""><i>No Name on Application</i></span>"
		Else
			strNameOnCard = objRS("NameOnCard")
			
			'strLink = "<a target=""_parent"" href=""Cards3.asp?CardID=" & objRS("CardID") & "&EmployeeID=" & objRS("EmployeeID") & " "" class=""btn btn-link"">"
		End If
		
		If IsNull(objRS("CreditLimit")) Then
			strCreditLimit = ""
		Else
			If IsNumeric(objRS("CreditLimit")) Then
				strCreditLimit = FormatCurrency(objRS("CreditLimit"),0)
			Else
				strCreditLimit = ""
			End If
		End If
		
		strLink = "<a href=""ApplicationsEmployeeHF.asp?ApplicationID=" & objRS("ApplicationID") & "&EmployeeID=" & objRS("EmployeeID") & " "">"
		
		Response.Write "<div class=""panel panel-shadow mb-3""><div class=""panel-content row"">" & _
					"<div class=""col-md-1 text-center my-auto"">" & strLink & "" & strCardType & "</div>" & _
					"<div class=""col-md-4""><h6>" & strCardType2 & " " & strAction & _
					"</h6><p>Number: <strong>" & strCardNo & "</strong></p></div>" & _
					"<div class=""col-md-4""><p>Limit: <strong>" & strCreditLimit & "</strong></p><p>Date Submitted: <strong>" & dteDateSubmitted & "</strong></p></div>" & _
					"<div class=""col-md-3 text-right my-auto"">" & strLink & "View Details <i class=""fa fa-arrow-right""></i></a></div></div></div>"
			
			y = y + 1
			
		objRS.movenext
	Loop
				
objRS.Close

End Sub


Set objRS = Nothing
Set objCon = Nothing

%>