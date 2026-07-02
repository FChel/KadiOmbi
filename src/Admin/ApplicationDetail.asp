
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
Dim strViewButton

'Set Database objects

Set objCon = Server.CreateObject("ADODB.Connection")
Set objCmd = Server.CreateObject("ADODB.Command")
Set objRS = Server.CreateObject("ADODB.Recordset")

'Open database connection
objCon.Open Session("DBConnection")

'Collect the action results to save details
If Not IsEmpty(Request.QueryString("Task")) Then
	'Response.Write Request.QueryString("TaskID") & "," & Request.QueryString("Task") & "," & Request.QueryString("TaskLogID")
	Call SaveTaskLog(Request.QueryString("TaskID"),Request.QueryString("Task"),Request.QueryString("TaskLogID"))
End If

If Not IsEmpty(Request.QueryString("ViewButton")) Then
	strViewButton = Request.QueryString("ViewButton")
End If

If Not IsEmpty(Request.QueryString("ApplicationID")) Then
	Session("ApplicationID") = Request.QueryString("ApplicationID")
End If

If Not IsEmpty(Request.QueryString("Action")) Then
	If Request.QueryString("Action") = "Save" Then
		Call SaveMessage
	End If	
End If
	
%>
<script>
function OpenSs(cb) {

	//alert("asas");
	//var e = document.getElementById(this.cb);
	//var result = e.options[e.selectedIndex].value;
	
	//document.getElementById('ContinueMod').value=result;
	document.getElementById('ModApp').showModal();
}




</script>
<!-- Modal -->
<div class="modal fade" id="ModApp" tabindex="-1" role="dialog" aria-labelledby="exampleModalCenterTitle" aria-hidden="true">
  <div class="modal-dialog modal-dialog-centered" role="document">
    <div class="modal-content">
      <div class="modal-header">
        <h5 class="modal-title" id="exampleModalLongTitle">CAPS Admin Tasks</h5>
        <button type="button" class="close" data-dismiss="modal" aria-label="Close">
          <span aria-hidden="true">&times;</span>
        </button>
      </div>
      <div class="modal-body">
        <label>Continue With Process</label><br>
            <input type="text" name="ContinueMod" id="ContinueMod" class="form-control input-md" value="">
			<input type="text" name="ContinueModTask" id="ContinueModTask" class="form-control input-md" value="">
      </div>
      <div class="modal-footer">
        <button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
      </div>
    </div>
  </div>
</div>


    <main class="main py-3">
      <div class="container">
		<form action="ApplicationDetail.asp?Action=Save" method="POST" id="frm" name="frm">
        <div class="row">
          <div class="col-md-6">
  
				<div class="panel-header">
					<h4>DTC Application</h4>
				 </div>
				 
				  
					 
						<%
						Call DisplayTableDetails()
						
						%>
			
				
            </div>
			
			 <div class="col-md-6 sidebar">
		  
            <div class="panel panel-shadow mb-3">
              <div class="panel panel-shadow panel-validation mb-3">
				  <div class="panel-header">
					<h4>Application Message</h4>
					<span class="panel-subheader">Recent Messages</span>
				  </div>
				  <div class="panel-content mb-8">
					<% Call LoadMessages() %>
					
					<textarea rows="4" id="MessageS" name="MessageS" class="form-control input-md" value="" placeholder="Type a message"></textarea>
					<div class="col-md-12 text-right my-auto">
					<button type="button" class="btn btn-primary" onClick="frm.submit();">Send Message</button>
					</div>
				</div>
			  </div>
            </div>
			
            <div class="panel panel-shadow mb-3">
              <div class="panel-header">
                <h4>Application Summary</h4>
                <span class="panel-subheader">Application Progress</span>
              </div>
			  <div class="panel-content row">
              <div class="col-md-3 text-left my-auto">
                  <span class="content">Submitted <i class="fa fa-arrow-right"></i> <p style="font-size:12px; font-weight:bold; color:black;"><%=now()%></p></span>
			</div>
				<div class="col-md-3 text-right my-auto">
                  <span class="content"><span style="color:green; font-weight:bold;">DOD Admin</span> <i class="fa fa-arrow-right"></i><p style="font-size:14px; font-weight:bold; color:black;">2 Days</p></span>
                </div>
			<div class="col-md-3 text-right my-auto">
                  <span class="content">Bank <i class="fa fa-arrow-right"></i><p style="font-size:14px; font-weight:normal; color:grey;">Not Sent</p></span>
                </div>
			<div class="col-md-3 text-right my-auto">
                  <span class="content">Mailed <i class="fa fa-credit-card"></i><p style="font-size:14px; font-weight:normal; color:grey;">Not Sent</p></span>
                </div>
			</div>
              </div>
            </div>
          </div>
		  </form>
        </div>
    </main>

	
    <script>
     $(".ActionButton1").on("show.bs.dropdown", function(event){
	  var x = $(event.relatedTarget).text(); // Get the text of the element
	  alert(x);
	});
    </script>
	
	
<!-- #Include file=CAPSFooter.asp -->
  </body>
</html>

<%


Public Sub DisplayTableDetails()

Dim strAction
Dim strStatus
Dim strAddress
Dim dteDateSubmitted
Dim dteDateReviewed
Dim strCreditLimit

'If Session("EmployeeID") = "" OR ISNull(Session("EmployeeID")) Then
'	strSQL = "SELECT * FROM qryCAPSApplications WHERE ApplicationID = '" & Session("ApplicationID") & "'"
'Else
	strSQL = "SELECT * FROM qryCAPSApplications WHERE ApplicationID = '" & Session("ApplicationID") & "'"
'End If

objRS.Open strSQL,objCon

	
    If Not objRS.EOF Then
		If isNull(objRS(9)) Then
			dblEmpCont = 0
		Else
			dblEmpCont = objRS(9)
		End If

		Select Case objRS("Status")
		
			Case  "Awaiting review"
				strAction = "<button type=""button"" class=""btn btn-primary btn-sm"" onclick=""self.location='ApplicationDetail.asp?Action=Release&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-check""></i> Release</button>"
				strAction = strAction & " <button type=""button"" class=""btn btn-outline-danger btn-sm"" onclick=""self.location='ApplicationDetail.asp?Action=Reject&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-times""></i> Reject</button>"
			Case "Added To CS"
				strAction = "<button type=""button"" class=""btn btn-secondary btn-sm"" onclick=""self.location='CSToDiners.asp?ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-file-o""></i> View CS</button>"
			
			Case "Submitted"
				strAction = "<button type=""button"" class=""btn btn-primary btn-sm"" onclick=""self.location='ApplicationDetail.asp?Action=Release&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-check""></i> Release</button>"
				strAction = strAction & " <button type=""button"" class=""btn btn-danger btn-sm"" onclick=""self.location='ApplicationDetail.asp?Action=Reject&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-times""></i> Reject</button>"
			
			Case "On Hold"
				strAction = "<button type=""button"" class=""btn btn-outline-secondary btn-sm"" onclick=""self.location='CSToDiners.asp?ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-file-o""></i> View CS</button>"
			
			Case "Cancelled"
				strAction = "Cancelled - " & FormatDateTime(objRS("DateUpdated"),vbShortDate)'<button type=""button"" class=""btn btn-danger"" onclick=""self.location='ApplicationDetail.asp?Action=Cancel&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-minus-circle""></i> Cancel</button>"
				strStatus  = "<button type=""button"" class=""btn btn-outline-danger btn-sm"" onclick=""self.location='ApplicationDetail.asp?ApplicationID=" & objrs("ApplicationID") & "'"";>Cancelled</button>"
			
			Case "Awaiting issue"
				strAction = "<button type=""button"" title=""Approved by GCFO"" class=""btn btn-outline-success btn-sm"" onclick=""self.location='ApplicationDetail.asp?ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-check""></i> Approved</button>"
				strStatus  = "<button type=""button"" class=""btn btn-outline-secondary btn-sm"" onclick=""self.location='ApplicationDetail.asp?ApplicationID=" & objrs("ApplicationID") & "'"";>Awaiting Issue</button>"
			
			Case Else
				strAction = "<button type=""button"" class=""btn btn-outline-danger btn-sm"" onclick=""self.location='ApplicationDetail.asp?Action=Cancel&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-minus-circle""></i> Cancel</button>"
				'strAction = "Rejected"
				strStatus  = "<button type=""button"" class=""btn btn-success btn-sm"" onclick=""self.location='ApplicationDetail.asp?ApplicationID=" & objrs("ApplicationID") & "'"";>Submitted</button>"
		End Select

		strAddress = Trim(objRS("Address1")) & " " & Trim(objRS("Address2")) & " " & Trim(objRS("Suburb")) & " " & Trim(objRS("State")) & " " & Trim(objRS("PostCode"))
		
		If len(strAddress) > 15 Then strAddress = left(strAddress,15) & "..."
		
		If IsNull(objRS("DateSubmitted")) Then
			dteDateSubmitted = ""
		Else
			dteDateSubmitted = FormatDateTime(objRS("DateSubmitted"),vbShortDate)
		End If
		
		If IsNull(objRS("DateReviewed")) Then
			dteDateReviewed = ""
		Else
			dteDateReviewed = FormatDateTime(objRS("DateReviewed"),vbShortDate)
		End If
		
		If IsNull(objRS("CreditLimit")) Then
			strCreditLimit = ""
		Else
			If IsNumeric(objRS("CreditLimit")) Then
				If objRS("CreditLimit")  > 0 Then
					strCreditLimit = FormatCurrency(objRS("CreditLimit")/100,0)
				Else
					strCreditLimit = FormatCurrency(objRS("CreditLimit"),0)
				End If
			Else
				strCreditLimit = objRS("CreditLimit")
			End If
			
		End If
		
		Response.Write "<div class=""panel-content row"" >" & _
			"<div class=""form-row col-md-4""><label>Application ID</label></div><div class=""form-row col-md-8""><input style=""border: 0px; font-weight:bold;"" type=""text"" id=""ApplicationID"" name=""ApplicationID"" class=""form-control input-md"" value=""" & objRS("ApplicationID") & """></div></div>" & _
			"<div class=""panel-content row"" >" & _
			"<div class=""form-row col-md-4""><label>Status</label></div><div class=""form-row col-md-8""><input style=""border: 0px; font-weight:bold;"" type=""text"" id=""Status"" name=""Status"" class=""form-control input-md"" value=""" & objRS("Status") & """></div></div>" & _
			"<div class=""panel-content row"" >" & _
			"<div class=""form-row col-md-4""><label>Credit Limit</label></div><div class=""form-row col-md-8""><input type=""text"" id=""CreditLimit"" name=""CreditLimit"" class=""form-control input-md"" value=""" & strCreditLimit & """></div></div>" & _
			"<div class=""panel-content row"" >" & _
			"<div class=""form-row col-md-4""><label>Date Submitted</label></div><div class=""form-row col-md-8""><input style=""border: 0px; font-weight:bold;"" type=""text"" id=""DateSubmitted"" name=""DateSubmitted"" class=""form-control input-md"" value=""" & objRS("DateSubmitted") & """></div></div>" & _
			"<div class=""panel-content row"" >" & _
			"<div class=""form-row col-md-4""><label>Group</label></div><div class=""form-row col-md-8""><input style=""border: 0px; font-weight:bold;"" type=""text"" id=""ReportGroup"" name=""ReportGroup"" class=""form-control input-md"" value=""" & objRS("ReportGroup") & """></div></div>" & _
			"<div class=""panel-content row"" >" & _
			"<div class=""form-row col-md-4""><label>Name On Card</label></div><div class=""form-row col-md-8""><input style=""border: 0px; font-weight:bold;"" type=""text"" id=""NameOnCard"" name=""NameOnCard"" class=""form-control input-md"" value=""" & objRS("NameOnCard") & """></div></div>" & _
			"<div class=""panel-content row"" >" & _
			"<div class=""form-row col-md-4""><label>Address 1</label></div><div class=""form-row col-md-8""><input style=""border: 0px; font-weight:bold;"" type=""text"" id=""Address1"" name=""Address1"" class=""form-control input-md"" value=""" & objRS("Address1") & """></div></div>" & _
			"<div class=""panel-content row"" >" & _
			"<div class=""form-row col-md-4""><label>Address 2</label></div><div class=""form-row col-md-8""><input style=""border: 0px; font-weight:bold;"" type=""text"" id=""Address2"" name=""Address2"" class=""form-control input-md"" value=""" & objRS("Address2") & """></div></div>" & _
			"<div class=""panel-content row"" >" & _
			"<div class=""form-row col-md-4""><label>Suburb State City</label></div><div class=""form-row col-md-8""><input style=""border: 0px; font-weight:bold;"" type=""text"" id=""Suburb"" name=""Suburb"" class=""form-control input-md"" value=""" & objRS("Suburb") & " " & objRS("State") & " " & objRS("PostCode") & """></div></div>"
		
	End If
	
	Response.Write strAction & " " & strStatus
	
objRS.Close

End Sub


Public Sub LoadMessages()
'Procedure to load any messages relating to the application
Dim strSQL
Dim strPerson

	strSQL = "SELECT * FROM qryCAPSMessage WITH(NOLOCK) WHERE [Object] = 'Application' AND [ObjectID] = '" & Session("ApplicationID") & "'"

	objRS.Open strSQL,objCon
	
    Do Until objRS.EOF
		
		If IsNull(objRS("MessageFrom") ) or objRS("MessageFrom")  = "" Then
			strPerson = ""
		Else
			If objRS("MessageFrom") = Session("UserID") Then
				strPerson = "(You)"
			Else
				If IsNull(objRS("MessageFrom") ) or objRS("MessageFrom")  = "" Then
					If objRS("MessageFrom") = 0 Then
						strPerson = "(Admin)"
					End If
				End If
			End If
		End If
		
		Response.write "<div class=""panel panel-light col-9""><div class=""panel-header"">" & _
			"<h6>" & objRS("UserFrom") & " " & strPerson & "</h6><span class=""panel-subheader"">" & objRS("MessageDetail") & "</span></div></div>"

		objRS.Movenext
	Loop
				
objRS.Close

End Sub

Public Sub SaveMessage()

Dim lngMessageID
Dim lngAdminID
Dim strMessage
Dim intRecord

If Session("MessageID") = "" or IsNull(Session("MessageID")) Then
	lngMessageID = 0
Else
	lngMessageID = Session("MessageID")
End If

If Session("AdminID") = "" or IsNull(Session("AdminID")) Then
	lngAdminID = 0
Else
	lngAdminID = Session("AdminID")
End If

If Request.Form("MessageS") = "" or IsNull(Request.Form("MessageS")) Then
	strMessage = ""
Else
	strMessage = Request.Form("MessageS")
End If

'Set the message title to Application
strMessageTitle = "Application"

	'Makes sure that there is content in the message or do not save
	If strMessage = "" Then
		Response.Write "<div class=""alert alert-danger"" role=""alert"">No Message detail to save...NOT SAVED!</div>"
	Else
		With objCmd

			.CommandType = 4
			.CommandText = "spCAPSMessageSave"

			.Parameters.Append objCmd.CreateParameter("MessageID", adInteger)
			.Parameters.Append objCmd.CreateParameter("MessageFrom", adInteger)
			.Parameters.Append objCmd.CreateParameter("MessageTo", adInteger)
			.Parameters.Append objCmd.CreateParameter("MessageTitle", adVarChar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("MessageDetail", adVarChar, adParamInput, 500)
			.Parameters.Append objCmd.CreateParameter("MessageDate", adInteger)
			.Parameters.Append objCmd.CreateParameter("MessageStatus", adVarChar, adParamInput, 20)
			.Parameters.Append objCmd.CreateParameter("MessageRead", adChar, adParamInput, 1)
			.Parameters.Append objCmd.CreateParameter("MessageThreadID", adInteger)
			.Parameters.Append objCmd.CreateParameter("Object", adVarChar, adParamInput, 50)
			.Parameters.Append objCmd.CreateParameter("ObjectID", adInteger)
			.Parameters.Append objCmd.CreateParameter("Active", adChar, adParamInput, 1)
			.Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger)
			.Parameters.Append objCmd.CreateParameter("MessageIDOutput", adInteger, adParamOutput)
			
			.Parameters("MessageID") = lngMessageID
			.Parameters("MessageFrom") = Session("UserID")
			.Parameters("MessageTo") = lngAdminID
			.Parameters("MessageTitle") = strMessageTitle
			.Parameters("MessageDetail") = strMessage
			.Parameters("MessageDate") = now()
			.Parameters("MessageStatus") = "Created"'Session("MessageStatus")
			.Parameters("MessageRead") = "N"'Session("MessageRead")
			.Parameters("MessageThreadID") = 0
			.Parameters("Object") = "Application"
			.Parameters("ObjectID") = Session("ApplicationID")
			.Parameters("Active") = "Y"
			.Parameters("UpdatedBy") = Session("UserID")
			.ActiveConnection = objCon
			 
		End With
	   
		objCmd.Execute        
	  
		'Return the result of the Save Function.
		intRecord = objCmd.Parameters.Item("MessageIDOutput") 
	 
		Response.Write "<div class=""alert alert-success"" role=""alert"">Message " & intRecord & " Saved!</div>"
	End If

End Sub



Set objRS = Nothing
Set objCon = Nothing

%>