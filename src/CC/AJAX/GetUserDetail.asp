<!DOCTYPE html>
<%

Dim objCon
Dim objRS
Dim x
Dim lngEmployeeID
Dim lngUserID
Dim lngEmployeeNo
Dim strFirstName
Dim strLastName
Dim strUserLogon
Dim intUserType
Dim strActive
Dim strComments
Dim strEmailAddress

Dim arrUserType(4)
Dim arrActive(2)
Dim strSelected

arrUserType(1) = "General"
arrUserType(2) = "Manager"
arrUserType(3) = "ReadOnly"
arrUserType(4) = "Admin"

arrActive(1) = "Y"
arrActive(2) = "N"

'Open Database Connection
Set objRS = Server.CreateObject("ADODB.Recordset")
Set objCon = Server.CreateObject("ADODB.Connection")

objCon.Open Session("DBConnection")

If Not IsEmpty(Request.QueryString("UserID")) Then
	lngUserID = Request.QueryString("UserID")
End If

Call Get_Record(lngUserID)

%>

    <form class="needs-validation" novalidate>
	<div class="container">
      	<div class="form-row">
        	<div class="col-md-3 mb-3">
          		<label for="EmailTemplateName">Employee ID</label>
				  <input type="text" class="form-control" name="EmployeeID" id="EmployeeID" value="<%=lngEmployeeID%>" maxlength="50" required>
				  <input type="hidden" class="form-control" name="UserID" id="UserID" value="<%=lngUserID%>" maxlength="50">
		  		<div class="valid-feedback">Valid</div>
		  		<div class="invalid-feedback">Employee ID must be entered.</div>
			</div>
		
			<div class="col-md-5 mb-3">
				<label for="EmailTemplateName">Last Name</label>
				<input type="text" class="form-control" name="LastName" id="LastName" value="<%=strLastName%>" maxlength="50" required>
				<div class="valid-feedback">Valid</div>
				<div class="invalid-feedback">Last Name must be entered.</div>
			</div>

			<div class="col-md-4 mb-3">
				<label for="EmailTemplateName">First Name</label>
				<input type="text" class="form-control" name="FirstName" id="FirstName" value="<%=strFirstName%>" maxlength="50" required>
				<div class="valid-feedback">Valid</div>
				<div class="invalid-feedback">First Name must be entered.</div>
			</div>
		</div>

		<div class="form-row">
			<div class="col-md-12 mb-3">
				<label for="EmailTemplateName">Email Address</label>
				<input type="text" class="form-control" name="EmailAddress" id="EmailAddress" value="<%=strEmailAddress%>" maxlength="50" required>
				<div class="valid-feedback">Valid</div>
				<div class="invalid-feedback">Email Address must be entered.</div>
			</div>
		</div>

		<div class="form-row">
			<div class="col-md-12 mb-3">
				<label for="EmailTemplateName">User Logon</label>
				<input type="text" class="form-control" name="UserLogon" id="UserLogon" value="<%=strUserLogon%>" maxlength="50" required>
				<div class="valid-feedback">Valid</div>
				<div class="invalid-feedback">User Logon must be entered.</div>
			</div>
		</div> 

	<div class="form-row">
		<div class="col-md-12 mb-3">
			<label for="EmailHeader">Comments</label>
			<textarea class="form-control" name="Comments" id="Comments" rows="4" cols="1" wrap="hard" ><%=strComments%></textarea>
		  	<div class="valid-feedback">
		  		Valid
			</div>
		</div>
	</div>	

	<div class="form-row">
		<div class="col-md-6 mb-3">
			<label for="CardType">User Type</label>
			<SELECT class="form-control" name="UserType" id="UserType" >
			<%
				objRS.Open "SELECT * FROM tblUserTypes WITH(NOLOCK) WHERE Active = 'Y'",objCon
		
				Do until objRS.EOF
					If objRS("UserTypeID") = intUserType Then
						strSelected = " SELECTED "
					Else
						strSelected = ""
					End if
						Response.Write "<option Value=""" & objRS("UserTypeID") & """" & strSelected & ">" & objRS("UserTypeName") & "</OPTION>"
					objRS.Movenext
				Loop
				
				objRS.Close
		
			%>
			</Select>
		</div>

		<div class="col-md-6 mb-3">
			<label for="CardTypeSub">Active</label>
			<SELECT class="form-control" name="Active" id="Active" >
				<%
				For x = 1 to 2
					If arrActive(x) = cstr(strActive) Then
						strSelected = " SELECTED "
					Else
						strSelected = ""
					End If
						Response.Write "<option Value=""" & arrActive(x) & """" & strSelected & ">" & arrActive(x) & "</OPTION>"
				Next
			%>
			</Select>
		</div>
	</div>


<div class="row">
	<div class="col-md-12 mb-3" style="text-align:right;">
		<button class="btn btn-primary btn-sm" type="submit"><i class="fa fa-check"></i> Save</button>
		<button type="button" class="btn btn-secondary btn-sm" data-dismiss="modal"><i class="fa fa-times"></i> Close</button>
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

<%
Public Sub Get_Record(lngUserID)

 'Description:	Loads Users

	objRS.Open "SELECT * FROM tblUsers WHERE UserID = " & lngUserID & "",objCon

		If Not objRS.EOF Then
			
			lngUserID = objRS("UserID")
			lngEmployeeID = objRS("EmployeeID")
			strLastName = objRS("LName")
			strFirstName = objRS("FName")
			strUserLogon = objRS("UserLogon")
			strEmailAddress = objRS("EmailAddress")
			intUserType = objRS("UserTypeID")
			strActive = objRS("Active")
			strComments = objRS("Comments")

		End If		

	objRS.Close
	
End Sub

%>