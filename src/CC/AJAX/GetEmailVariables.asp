<!DOCTYPE html>
<%

Dim objCon
Dim objRS
Dim x
Dim str_string_pattern
Dim strSQLSearchString
Dim strSelected

'Open Database Connection
Set objRS = Server.CreateObject("ADODB.Recordset")
Set objCon = Server.CreateObject("ADODB.Connection")

objCon.Open Session("DBConnection")

If Not IsEmpty(Request.QueryString("string_pattern")) Then
	str_string_pattern = Request.QueryString("string_pattern")
	Session("string_pattern") = str_string_pattern
End If

Public Sub Get_Record()

 'Description:	Loads tblCAPSEmailDetail
	
	objRS.Open "SELECT * FROM tblStringReplace WHERE string_pattern = '" & Session("string_pattern") & "'",objCon
	
		If Not objRS.EOF Then

				str_string_pattern = objRS("string_pattern")
			  	strSQLSearchString = objRS("SQLSearchString")
			  
		End If		

	objRS.Close
	
End Sub

Call Get_Record()

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
        	<div class="col-md-12 mb-3">
				<label for="VariableName">Variable Name</label>
				<input type="text" class="form-control" name="VariableName" id="VariableName" value="<%=str_string_pattern%>" maxlength="50" required>
				<div class="valid-feedback">Valid </div>
		  		<div class="invalid-feedback">Variable name must be entered.</div>
			</div>
		</div>

		<div class="form-row">
        	<div class="col-md-12 mb-3">
          		<label for="SQLSearchString">SQL Seacrh String</label>
          		<input type="text" class="form-control" name="SQLSearchString" id="SQLSearchString" placeholder="SQL Search String" value="<%=strSQLSearchString%>" maxlength="200" required>
          		<div class="valid-feedback">Valid</div>
		  		<div class="invalid-feedback">SQL Search String must be entered.</div>
			</div>
		</div> 

		<div class="modal-footer">
			<button type="button" class="btn btn-link" data-dismiss="modal">Cancel</button>
			<button type="submit" class="btn btn-primary">Save changes</button>
  		</div>

		<div class="row">
	  		<div class="col-md-12 mb-3">		
				<div class="alert alert-success" role="alert" id="AlertSuccess" style="display:none">Record has been successfully saved.</div>
				<div class="alert alert-danger" role="alert" id="AlertDanger" style="display:none">Invalid entries exist.  Please correct and resubmit.</div>
		</div>
		
	</div>

</form>
</body>
</html>

