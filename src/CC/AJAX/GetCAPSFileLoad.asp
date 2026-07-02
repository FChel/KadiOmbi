<!DOCTYPE html>
<%

Dim objCon
Dim objRS
Dim x
Dim lngFileLoadID

Dim arrStatus(4)
Dim arrDeleted(2)
Dim strSelected
Dim strDeleted
Dim strStatus

arrStatus(1) = "Processed"
arrStatus(2) = "Imported"
arrStatus(3) = "Exported"
arrStatus(4) = ""

arrDeleted(1) = "Y"
arrDeleted(2) = "N"

'Open Database Connection
Set objRS = Server.CreateObject("ADODB.Recordset")
Set objCon = Server.CreateObject("ADODB.Connection")

objCon.Open Session("DBConnection")

If Not IsEmpty(Request.QueryString("FileLoadID")) Then
	lngFileLoadID = Request.QueryString("FileLoadID")
	
End If

Call Get_Record(lngFileLoadID)

%>

    <form class="needs-validation" novalidate>
	<div class="container">
      	
	<div class="form-row">
		<div class="col-md-6 mb-3">
			<label for="CardType">Status</label>
			<SELECT class="form-control" name="Status" id="Status" >
			<%
				For x = 1 to 4
					If cstr(arrStatus(x)) = cstr(strStatus) Then
						strSelected = " SELECTED "
					Else
						strSelected = ""
					End If
						Response.Write "<option Value=""" & arrStatus(x) & """" & strSelected & ">" & arrStatus(x) & "</OPTION>"
				Next
			%>
			</Select>
		</div>

		<div class="col-md-6 mb-3">
			<label for="CardTypeSub">Deleted</label>
			<SELECT class="form-control" name="Deleted" id="Deleted" >
			<%
				For x = 1 to 2
					If arrDeleted(x) = cstr(strDeleted) Then
						strSelected = " SELECTED "
					Else
						strSelected = ""
					End If
						Response.Write "<option Value=""" & arrDeleted(x) & """" & strSelected & ">" & arrDeleted(x) & "</OPTION>"
				Next
			%>
			</Select>
		</div>
		 <input type="hidden" class="form-control" name="FileLoadID" id="FileLoadID" value="<%=lngFileLoadID%>" maxlength="50" required>
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
Public Sub Get_Record(lngFileLoadID)

 'Description:	Loads Users

	objRS.Open "SELECT * FROM tblCAPSFileLoad WHERE FileLoadID = " & lngFileLoadID & "",objCon

		If Not objRS.EOF Then
			
			strStatus = objRS("Status")
			strDeleted = objRS("Deleted")
			
		End If		

	objRS.Close
	
End Sub

%>