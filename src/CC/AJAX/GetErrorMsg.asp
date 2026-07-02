<!DOCTYPE html>
<%

Dim objCon
Dim objRS
Dim x
Dim lngEmailErrorMsgID
Dim strEmailErrorMsg
Dim strEmailErrorMsgNo
Dim strEmailErrorMsgFriendly
Dim strCardType
Dim strCardTypeSub

Dim arrCardType(2)
Dim arrCardTypeSub(4)
Dim strSelected

arrCardType(1) = "DTC"
arrCardType(2) = "DPC"

arrCardTypeSub(1) = "Diners"
arrCardTypeSub(2) = "CTS"
arrCardTypeSub(3) = "Mastercard"
arrCardTypeSub(4) = "ANZ"

'Open Database Connection
Set objRS = Server.CreateObject("ADODB.Recordset")
Set objCon = Server.CreateObject("ADODB.Connection")

objCon.Open Session("DBConnection")

If Not IsEmpty(Request.QueryString("EmailErrorMsgID")) Then
	lngEmailErrorMsgID = Request.QueryString("EmailErrorMsgID")
End If

Call Get_Record(lngEmailErrorMsgID)

%>
<html>
  <head>
    <meta http-equiv="X-UA-Compatible" content="IE=edge" />
    <meta charset="utf-8" />
    <title>MyFi - Bootstrap Components</title>
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <link href="fontawesome/css/all.css" rel="stylesheet" />
    <link rel="stylesheet" href="css/myfi-bootstrap.css" />
  </head>
 
  <body>
    <form class="needs-validation" novalidate>
	<div class="container">
      <div class="form-row">
        <div class="col-md-4 mb-3">
          <label for="EmailTemplateName">xx Error Msg No</label>
		  <input type="text" class="form-control" name="EmailErrorMsgNo" id="EmailErrorMsgNo" value="<%=strEmailErrorMsgNo%>" maxlength="50" required>
		  <input type="hidden" class="form-control" name="EmailErrorMsgID" id="EmailErrorMsgID" value="<%=lngEmailErrorMsgID%>" >
          <div class="valid-feedback">
            Valid
		  </div>
		  <div class="invalid-feedback">
			Message No must be entered.
		  </div>
		</div>

        <div class="col-md-8 mb-3">
          <label for="EmailSubject">Error Message</label>
          <input type="text" class="form-control" name="EmailErrorMsg" placeholder="EmailErrorMsg" value="<%=strEmailErrorMsg%>" maxlength="200" required>
          <div class="valid-feedback">
            Valid
		  </div>
		  <div class="invalid-feedback">
				Error Message must be entered.
		  </div>
		</div>
	</div> 

	<div class="form-row">
		<div class="col-md-12 mb-3">
			<label for="EmailHeader">Error Message Friendly</label>
			<textarea class="form-control" name="EmailErrorMsgFriendly" id="EmailErrorMsgFriendly" rows="15" cols="1" wrap="hard" required><%=strEmailErrorMsgFriendly%></textarea>
		  	<div class="valid-feedback">
		  		Valid
			</div>
		</div>
	</div>	

	<div class="form-row">
		<div class="col-md-6 mb-3">
			<label for="CardType">Card Type</label>
			<SELECT class="form-control" name="CardType" id="CardType" >
			<%
				For x = 1 to 2
					If arrCardType(x) = cstr(strCardType) Then
						strSelected = " SELECTED "
					Else
						strSelected = ""
					End If
						Response.Write "<option Value=""" & arrCardType(x) & """" & strSelected & ">" & arrCardType(x) & "</OPTION>"
				Next
			%>
			</Select>
		</div>

		<div class="col-md-6 mb-3">
			<label for="CardTypeSub">Card Sub Type</label>
			<SELECT class="form-control" name="CardTypeSub" id="CardTypeSub" >
				<%
				For x = 1 to 4
					If arrCardTypeSub(x) = cstr(strCardTypeSub) Then
						strSelected = " SELECTED "
					Else
						strSelected = ""
					End If
						Response.Write "<option Value=""" & arrCardTypeSub(x) & """" & strSelected & ">" & arrCardTypeSub(x) & "</OPTION>"
				Next
			%>
			</Select>
		</div>
	</div>
</div>

<div class="row">
	<div class="col-md-12 mb-3">
		<button class="btn btn-primary btn-sm" type="submit">Save</button>
		<button type="button" class="btn btn-secondary" data-dismiss="modal"><i class="fa fa-close"></i> Close</button>
	</div>
</div>

	<div class="row">
	  <div class="col-md-12 mb-3">
		
		<div class="alert alert-success" role="alert" id="AlertSuccess" style="display:none">
		  	Record has been successfully saved.
		</div>

		<div class="alert alert-danger" role="alert" id="AlertDanger" style="display:none">
		  	Invalid entries exist.  Please correct and resubmit.
		</div>
		
	  </div>
	</div>
</div>

</form>
</body>
</html>
<%
Public Sub Get_Record(lngRecordID)

 'Description:	Loads SELECT * FROM tblCAPSEmailErrorMsg

	objRS.Open "SELECT * FROM tblCAPSEmailErrorMsg WHERE EmailErrorMsgID = " & lngRecordID & "",objCon

		If Not objRS.EOF Then
			
			strEmailErrorMsgNo = objRS("EmailErrorMsgNo")
			strEmailErrorMsg = objRS("EmailErrorMsg")
			strEmailErrorMsgFriendly = objRS("EmailErrorMsgFriendly")
			strCardType = objRS("CardType")
			strCardTypeSub = objRS("CardTypeSub")

		End If		

	objRS.Close
	
End Sub
%>