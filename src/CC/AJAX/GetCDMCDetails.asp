<!-- #Include file=../CAPSFunctions.asp -->
<html lang="en">
<%

Dim objCon
Dim objRS
Dim x
Dim strEmployeeID
Dim strStatus
Dim strName
Dim strSubStatPos
Dim strNameCard
Dim strSubStatPosCard
Dim strPostalMessage

'Open Database Connection
Set objCon = Server.CreateObject("ADODB.Connection")
    
    objCon.Open Session("DBConnection")

Set objRS = Server.CreateObject("ADODB.Recordset")

If Not IsEmpty(Request.QueryString("EmployeeID")) Then
	strEmployeeID = " WHERE EmployeeID = '" & Request.QueryString("EmployeeID") & "'"
End If


Public Sub ShowDetails()

	'Description:	Loads CDMC Details onto the page called from
	objRS.Open "SELECT * FROM qryCAPSCDMCHistoryActive WITH(NOLOCK) " & strEmployeeID,objCon

			  
		If Not objRS.EOF Then
			
			'Write the Header			
			Response.write "<div class=""row""><div class=""col-md-12"">" & _
				"<table class=""table table-compact"">"
				
			'Make the CS Title Given names and Surname into one field/variable (for space savings)
			If IsNull(objRS("Surname")) Then
				strName = ""
			Else
				strName = Trim(Trim(objRS("Title")) & " " & Trim(objRS("FirstName")) & " " & Trim(objRS("Surname")))
			End If
			
			'Make the CDMC Suburb, State, PostCode into one field/variable (for space savings)
			If IsNull(objRS("OutSuburb")) Then
				strSubStatPosCard = ""
			Else
				strSubStatPosCard = Trim(Trim(objRS("OutSuburb")) & " " & Trim(objRS("OutState")) & " " & Trim(objRS("OutPostCode")))
			End If	
			
			
			If IsNull(Trim(objRS("IsValidPostal"))) Then
				strPostalMessage = "<span class=""badge badge-pill badge-danger"">No</span>"
			Else
				If Trim(objRS("IsValidPostal")) = "Y" Then
					strPostalMessage = "<span class=""badge badge-pill badge-success"">Yes</span>"
				Else
					strPostalMessage = "<span class=""badge badge-pill badge-danger"">No</span>"
				End If
			End If
			
			Response.Write "<tr><td style=""font-weight:bold;"">CDMC ID</td><td>" & objRS("CDMCID") & "</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">EID No</td><td>" & objRS("EmployeeID") & "</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">Cardholder</td><td>" & strName & "</td></tr>" & _
				"<tr class=""updated""><td colspan=""2"">CAPS Formatted Address</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">Formatted Address 1</td><td>" & Trim(objRS("OutDinersAddress1")) & "</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">Formatted Address 2</td><td>" & Trim(objRS("OutDinersAddress2")) & "</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">Formatted Suburb/ State/ PostCode</td><td>" & strSubStatPosCard & "</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">Formatted Work Phone</td><td>" & objRS("OutDinersWorkPhone") & "</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">Formatted Mobile Phone</td><td>" & objRS("OutDinersMobilePhone") & "</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">Email</td><td>" & Trim(objRS("Email_Address")) & "</td></tr>" & _
				"<tr class=""updated""><td colspan=""2"">Corporate Directory Details Added by Employee</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">Postal Address Unit &nbsp;&nbsp;<span style=""font-weight:italic; font-size:10px;"">" & Len(Trim(objRS("PostalAddress_Unit"))) & " chars</span></td><td>" & Trim(objRS("PostalAddress_Unit")) & "</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">Postal Address Client Location &nbsp;&nbsp;<span style=""font-weight:italic; font-size:10px;"">" & Len(Trim(objRS("PostalAddress_ClientLocation"))) & " chars</span></td><td>" & Trim(objRS("PostalAddress_ClientLocation")) & "</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">Postal Address Delivery Location &nbsp;&nbsp;<span style=""font-weight:italic; font-size:10px;"">" & Len(Trim(objRS("PostalAddress_DeliveryLocation"))) & " chars</span></td><td>" & Trim(objRS("PostalAddress_DeliveryLocation")) & "</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">Postal Address City</td><td>" & Trim(objRS("PostalAddress_City")) & "</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">Postal Address State</td><td>" & Trim(objRS("PostalAddress_State")) & "</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">Postal Address PostCode</td><td>" & Trim(objRS("PostalAddress_PostCode")) & "</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">Telephone Number</td><td>" & Trim(objRS("TelephoneNumber")) & "</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">Mobile Number</td><td>" & Trim(objRS("MobileNumber")) & "</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">Date Of Birth</td><td>" & Trim(objRS("DateOfBirth")) & "</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">Postal Address OK</td><td>" & strPostalMessage & "</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">Postal Message</td><td>" & Trim(objRS("PostalMessage")) & "</td></tr>"
			
			'Response.Write "</table></div>"

			'Response.Write "</table></div><div class=""col-md-6""><div class=""row mb-3""><div class=""col-md-6 my-auto""><button type=""button"" class=""btn btn-outline-secondary"" onClick=""loadDoc(" & objRS(0) & ")"">Card Details</button></div>" & _
			'	"<div class=""col-md-6""><select class=""form-control""><option>Record 1</option><option>Record 2</option><option>Record 3</option></select>" & _
			'	"</div></div><table class=""table table-compact"">"
			
		Else
			Response.write "No CDMC Record for " & Request.QueryString("EmployeeID")
	   End If

	objRS.Close
	
	Response.write "</div></table>"
	
	
End Sub
%>
<head>
  
</head>

<body>
  <%
  Call ShowDetails()
  
  %>
</body>

</html>
<%

Set objRS = Nothing
Set objCon = Nothing
  
  %>