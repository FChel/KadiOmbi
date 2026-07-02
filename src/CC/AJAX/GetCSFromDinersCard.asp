<!-- #Include file=../CAPSFunctions.asp -->
<html lang="en">
<%

Dim objCon
Dim objRS
Dim x
Dim strCSFromDinersID
Dim strStatus
Dim strName
Dim strSubStatPos
Dim strNameCard
Dim strSubStatPosCard

'Open Database Connection
Set objCon = Server.CreateObject("ADODB.Connection")
    
    objCon.Open Session("DBConnection")

Set objRS = Server.CreateObject("ADODB.Recordset")

If Not IsEmpty(Request.QueryString("CSFromDinersID")) Then
	strCSFromDinersID = " WHERE CSFromDinersID = " & Request.QueryString("CSFromDinersID") & ""
End If


Public Sub ShowDetails()

	'Description:	Loads Position details into page if applicable.
	objRS.Open "SELECT * FROM qryCAPSCSFromDinersCard " & strCSFromDinersID,objCon

			  
		If Not objRS.EOF Then
		
			
			'Write the Header
			'Response.write "<div class=""panel-content row""><div class=""form-row col-md-6"" style=""font-weight:bold;"">CS From Diners</div><div class=""form-row col-md-6"" style=""font-weight:bold;""><A HREF=""#"" onClick=""loadDoc(" & objRS(0) & ")"">Card Details</a> </div></div>"
			  
			'Response.write "<div class=""row""><div class=""col-md-6""><h6 class=""mb-3"">CS From Diners</h6><table class=""table table-compact"">"
			
			Response.write "<div class=""row""><div class=""col-md-6""><div class=""row mb-3""><div class=""col-md-6""><h6 class=""mb-3"">CS From Diners</h6></div>" & _
				"<div class=""col-md-6""><select class=""form-control""><option>Record 1</option><option>Record 2</option><option>Record 3</option></select></div></div><table class=""table table-compact"">"
				
			'Create the Status list badge based on the status field
			If IsNull(objRS("Status")) Then
				strStatus = ""
			Else
				If objRS("Status") = "Processed" Then
					strStatus = "<span class=""badge badge-pill badge-success"">Processed</span>"
				ElseIf objRS("Status") = "Imported" Then
					strStatus = "<span class=""badge badge-pill badge-warning"">Imported</span>"
				Else
					strStatus = objRS("Status")
				End If
			End If
			'Make the CS Title Given names and Surname into one field/variable (for space savings)
			If IsNull(objRS("Surname")) Then
				strName = ""
			Else
				strName = Trim(Trim(objRS("Title")) & " " & Trim(objRS("GivenNames")) & " " & Trim(objRS("Surname")))
			End If
			
			'Make the CARD Title Given names and Surname into one field/variable (for space savings)
			If IsNull(objRS("CardSurname")) Then
				strNameCard = ""
			Else
				strNameCard = Trim(Trim(objRS("CardTitle")) & " " & Trim(objRS("FirstName")) & " " & Trim(objRS("CardSurname")))
			End If
			'Make the CS Suburb, State, PostCode into one field/variable (for space savings)
			If IsNull(objRS("Suburb")) Then
				strSubStatPos = ""
			Else
				strSubStatPos = Trim(Trim(objRS("Suburb")) & " " & Trim(objRS("State")) & " " & Trim(objRS("PostCode")))
			End If
			'Make the CARD Suburb, State, PostCode into one field/variable (for space savings)
			If IsNull(objRS("CardSuburb")) Then
				strSubStatPosCard = ""
			Else
				strSubStatPosCard = Trim(Trim(objRS("CardSuburb")) & " " & Trim(objRS("CardState")) & " " & Trim(objRS("CardPostCode")))
			End If
			
			'Response.write "<tr onClick=""SelectFileSeqNum(String('" & objRS("CSFromDinersID") & "'));""><td>" & objRS("CSFromDinersID") & "</td><td>" & objRS("FileDateTime") & "</td><td>" & objRS("Status") & "</td>" & _
			'			"<td>" & objRS("EIDNo") & "</td><td>" & objRS("Title") & " " & objRS("GivenNames") & " " & objRS("Surname") & "</td></tr>"
			
			
			Response.Write "<tr><td>CS ID</td><td>" & objRS("CSFromDinersID") & "</td></tr>" & _
				"<tr><td>Status</td><td>" & strStatus & "</td></tr>" & _
				"<tr><td>File Seq Num</td><td>" & objRS("FileSeqNum") & "</td></tr>" & _
				"<tr><td>EID No</td><td>" & objRS("EIDNo") & "</td></tr>" & _
				"<tr><td>Cardholder</td><td>" & strName & "</td></tr>" & _
				"<tr><td>Card Number</td><td>" & MaskCard(objRS("CardNo")) & "</td></tr>" & _
				"<tr><td>CardUpdateInd</td><td>" & objRS("CardUpdateInd") & "</td></tr>" & _
				"<tr><td>Card Expiry Date</td><td>" & objRS("CardExpiryDate") & "</td></tr>" & _
				"<tr><td>Name On Card</td><td>" & Trim(objRS("NameOnCard")) & "</td></tr>" & _
				"<tr><td>Address 1</td><td>" & Trim(objRS("Address1")) & "</td></tr>" & _
				"<tr><td>Address 2</td><td>" & Trim(objRS("Address2")) & "</td></tr>" & _
				"<tr><td>Address 3</td><td>" & Trim(objRS("Address3")) & "</td></tr>" & _
				"<tr><td>Suburb/ State/ PostCode</td><td>" & strSubStatPos & "</td></tr>" & _
				"<tr><td>Home Phone</td><td>" & objRS("HomePhone") & "</td></tr>" & _
				"<tr><td>Work Phone</td><td>" & objRS("WorkPhone") & "</td></tr>" & _
				"<tr><td>Mobile Phone</td><td>" & objRS("MobilePhone") & "</td></tr>" & _
				"<tr><td>Email</td><td>" & Trim(objRS("Email")) & "</td></tr>" & _
				"<tr><td>Report Group</td><td>" & objRS("ReportGroup") & "</td></tr>" & _
				"<tr><td>Credit Limit</td><td>" & objRS("CreditLimit") & "</td></tr>" & _
				"<tr><td>Relationship</td><td>" & objRS("Relationship") & "</td></tr>" & _
				"<tr><td>Cat2</td><td>" & objRS("Cat2") & "</td></tr>" & _
				"<tr><td>Account Number</td><td>" & objRS("AccountNumber") & "</td></tr>" & _
				"<tr><td>ActivationFlag</td><td>" & objRS("ActivationFlag") & "</td></tr>" & _
				"<tr><td>PlasticID</td><td>" & objRS("PlasticID") & "</td></tr>" & _
				"<tr><td>Companion</td><td>" & MaskCard(objRS("Companion")) & "</td></tr>" & _
				"<tr><td>Status</td><td>" & objRS("Status") & "</td></tr>" 
			
			Response.Write "</table></div><div class=""col-md-6""><div class=""row mb-3""><div class=""col-md-6 my-auto""><button type=""button"" class=""btn btn-outline-secondary"" onClick=""loadDoc(" & objRS(0) & ")"">Card Details</button></div>" & _
				"<div class=""col-md-6""><select class=""form-control""><option>Record 1</option><option>Record 2</option><option>Record 3</option></select>" & _
				"</div></div><table class=""table table-compact"">"

			Response.Write "<tr><td>Card ID</td><td>" & objRS("CardID") & "</td></tr>" & _
				"<tr><td>Card Status</td><td>" & objRS("CardStatus2") & "</td></tr>" & _
				"<tr><td>Card Type</td><td>" & objRS("CardType") & "</td></tr>" & _
				"<tr><td>Card Type Sub</td><td>" & objRS("CardTypeSub") & "</td></tr>" & _
				"<tr><td>Employee ID</td><td>" & objRS("EmployeeID") & "</td></tr>" & _
				"<tr><td>Card Name</td><td>" & strNameCard & "</td></tr>" & _
				"<tr><td>Card Number</td><td>" & MaskCard(objRS("CardNumber")) & "</td></tr>" & _
				"<tr><td>CardCard Update Ind</td><td>" & objRS("CardCardUpdateInd") & "</td></tr>" & _
				"<tr><td>Expiry</td><td>" & objRS("Expiry") & "</td></tr>" & _
				"<tr><td>Name On Card</td><td>" & objRS("CardNameOnCard") & "</td></tr>" & _
				"<tr><td>Address 1</td><td>" & objRS("CardAddress1") & "</td></tr>" & _
				"<tr><td>Address 2</td><td>" & objRS("CardAddress2") & "</td></tr>" & _
				"<tr><td>Address 3</td><td>" & objRS("CardAddress3") & "</td></tr>" & _
				"<tr><td>Suburb/ State/ PostCode</td><td>" & strSubStatPosCard & "</td></tr>" & _
				"<tr><td>Home Phone</td><td>" & objRS("CardHomePhone") & "</td></tr>" & _
				"<tr><td>Work Phone</td><td>" & objRS("CardWorkPhone") & "</td></tr>" & _
				"<tr><td>Mobile Phone</td><td>" & objRS("CardMobilePhone") & "</td></tr>" & _
				"<tr><td>Email</td><td>" & objRS("CardEmail") & "</td></tr>" & _
				"<tr><td>Report Group</td><td>" & objRS("CardReportGroup") & "</td></tr>" & _
				"<tr><td>Credit Limit</td><td>" & objRS("CardCreditLimit") & "</td></tr>" & _
				"<tr><td>Relationship</td><td>" & objRS("CardRelationship") & "</td></tr>" & _
				"<tr><td>Cat 2</td><td>" & objRS("CardCat2") & "</td></tr>" & _
				"<tr><td>Account Number</td><td>" & objRS("CardAccountNumber") & "</td></tr>" & _
				"<tr><td>Activation Flag</td><td>" & objRS("CardActivationFlag") & "</td></tr>" & _
				"<tr><td>Plastic ID</td><td>" & objRS("CardPlasticID") & "</td></tr>"  &_
				"<tr><td>Companion</td><td>" &  MaskCard(objRS("CardCompanion")) & "</td></tr>" 
				
		
			'Response.Write "<div class=""panel-content row"" >" & _
			'	"<div class=""form-row col-md-2""><label class=""ModTextLabel"">CS ID</label></div><div class=""form-row col-md-4""><input class=""ModText"" type=""text"" id=""CSID"" name=""CSID"" class=""form-control input-md"" value=""" & objRS("CSFromDinersID") & """></div>" & _
			'	"<div class=""form-row col-md-2""><label class=""ModTextLabel"">Card ID</label></div><div class=""form-row col-md-4""><input class=""ModTextAudit"" type=""text"" id=""CardID"" name=""CardID"" class=""form-control input-md"" value=""" & objRS("CardID") & """></div></div>" & _
			'	"<div class=""panel-content row"" >" & _
			'	"<div class=""form-row col-md-2""><label class=""ModTextLabel"">Status</label></div><div class=""form-row col-md-2""><input class=""ModText"" type=""text"" id=""Status"" name=""Status"" class=""form-control input-md"" value=""" & objRS("Status") & """></div><div class=""form-row col-md-2"">" & strStatus & "</div>" & _
			'	"<div class=""form-row col-md-2""><label class=""ModTextLabel"">Card Status</label></div><div class=""form-row col-md-4""><input class=""ModTextAudit"" type=""text"" id=""CardStatus"" name=""CardStatus"" class=""form-control input-md"" value=""" & objRS("CardStatus2") & """></div></div>" & _
			'	"<div class=""panel-content row"" >" & _
			'	"<div class=""form-row col-md-2""><label class=""ModTextLabel"">File Date Time</label></div><div class=""form-row col-md-4""><input class=""ModText"" type=""text"" id=""FileDateTime"" name=""FileDateTime"" class=""form-control input-md"" value=""" & objRS("FileDateTime") & """></div>" & _
			'	"<div class=""form-row col-md-2""><label class=""ModTextLabel"">Type</label></div><div class=""form-row col-md-4""><input class=""ModTextAudit"" type=""text"" id=""Type"" name=""Type"" class=""form-control input-md"" value=""" & objRS("CardType") & """></div></div>" & _
			'	"<div class=""panel-content row"" >" & _
			'	"<div class=""form-row col-md-2""><label class=""ModTextLabel"">File Seq Num</label></div><div class=""form-row col-md-4""><input class=""ModText"" type=""text"" id=""FileSeqNum"" name=""FileSeqNum"" class=""form-control input-md"" value=""" & objRS("FileSeqNum") & """></div>" & _
			'	"<div class=""form-row col-md-2""><label class=""ModTextLabel"">CardTypeSub</label></div><div class=""form-row col-md-4""><input class=""ModTextAudit"" type=""text"" id=""CardTypeSub"" name=""CardTypeSub"" class=""form-control input-md"" value=""" & objRS("CardTypeSub") & """></div></div>" & _
			'	"<div class=""panel-content row"" >" & _
			'	"<div class=""form-row col-md-2""><label class=""ModTextLabel"">EID No</label></div><div class=""form-row col-md-4""><input class=""ModText"" type=""text"" id=""EIDNo"" name=""EIDNo"" class=""form-control input-md"" value=""" & objRS("EIDNo") & """></div>" & _
			'	"<div class=""form-row col-md-2""><label class=""ModTextLabel"">EmployeeID</label></div><div class=""form-row col-md-4""><input class=""ModTextAudit"" type=""text"" id=""EmployeeID"" name=""EmployeeID"" class=""form-control input-md"" value=""" & objRS("EmployeeID") & """></div></div>" & _
			'	"<div class=""panel-content row"" >" & _
			'	"<div class=""form-row col-md-2""><label class=""ModTextLabel"">Cardholder</label></div><div class=""form-row col-md-4""><input class=""ModText"" type=""text"" id=""Cardholder"" name=""Cardholder"" class=""form-control input-md"" value=""" & strName & """></div>" & _
			'	"<div class=""form-row col-md-2""><label class=""ModTextLabel"">Card Name</label></div><div class=""form-row col-md-4""><input class=""ModTextAudit"" type=""text"" id=""CardName"" name=""CardName"" class=""form-control input-md"" value=""" & strNameCard & """></div></div>" & _
			'	"<div class=""panel-content row"" >" & _
			'	"<div class=""form-row col-md-2""><label class=""ModTextLabel"">Card Number</label></div><div class=""form-row col-md-4""><input class=""ModText"" type=""text"" id=""CardNo"" name=""CardNo"" class=""form-control input-md"" value=""" & MaskCard(objRS("CardNo")) & """></div>" & _
			'	"<div class=""form-row col-md-2""><label class=""ModTextLabel"">Card Number</label></div><div class=""form-row col-md-4""><input class=""ModTextAudit"" type=""text"" id=""CardNumber"" name=""CardNumber"" class=""form-control input-md"" value=""" & MaskCard(objRS("CardNumber")) & """></div></div>" & _
			'	"<div class=""panel-content row"" >" & _
			'	"<div class=""form-row col-md-2""><label class=""ModTextLabel"">CardUpdateInd</label></div><div class=""form-row col-md-4""><input class=""ModText"" type=""text"" id=""CardUpdateInd"" name=""CardUpdateInd"" class=""form-control input-md"" value=""" & objRS("CardUpdateInd") & """></div>" & _
			'	"<div class=""form-row col-md-2""><label class=""ModTextLabel"">CardUpdateInd</label></div><div class=""form-row col-md-4""><input class=""ModTextAudit"" type=""text"" id=""CardCardUpdateInd"" name=""CardCardUpdateInd"" class=""form-control input-md"" value=""" & objRS("CardCardUpdateInd") & """></div></div>" & _
			'	"<div class=""panel-content row"" >" & _
			'	"<div class=""form-row col-md-2""><label class=""ModTextLabel"">Card Expiry Date</label></div><div class=""form-row col-md-4""><input class=""ModText"" type=""text"" id=""CardExpiryDate"" name=""CardExpiryDate"" class=""form-control input-md"" value=""" & objRS("CardExpiryDate") & """></div>" & _
			'	"<div class=""form-row col-md-2""><label class=""ModTextLabel"">Expiry</label></div><div class=""form-row col-md-4""><input class=""ModTextAudit"" type=""text"" id=""Expiry"" name=""Expiry"" class=""form-control input-md"" value=""" & objRS("Expiry") & """></div></div>" & _ 
			'	"<div class=""panel-content row"" >" & _
			'	"<div class=""form-row col-md-2""><label class=""ModTextLabel"">Name On Card</label></div><div class=""form-row col-md-4""><input class=""ModText"" type=""text"" id=""NameOnCard"" name=""NameOnCard"" class=""form-control input-md"" value=""" & Trim(objRS("NameOnCard")) & """></div>" & _
			'	"<div class=""form-row col-md-2""><label class=""ModTextLabel"">CardNameOnCard</label></div><div class=""form-row col-md-4""><input class=""ModTextAudit"" type=""text"" id=""CardNameOnCard"" name=""CardNameOnCard"" class=""form-control input-md"" value=""" & objRS("CardNameOnCard") & """></div></div>" & _ 
			'	"<div class=""panel-content row"" >" & _
			'	"<div class=""form-row col-md-2""><label class=""ModTextLabel"">Address 1</label></div><div class=""form-row col-md-4""><input class=""ModText"" type=""text"" id=""Address1"" name=""Address1"" class=""form-control input-md"" value=""" & Trim(objRS("Address1")) & """></div>" & _
			'	"<div class=""form-row col-md-2""><label class=""ModTextLabel"">CardAddress1</label></div><div class=""form-row col-md-4""><input class=""ModTextAudit"" type=""text"" id=""CardAddress1"" name=""CardAddress1"" class=""form-control input-md"" value=""" & objRS("CardAddress1") & """></div></div>" & _
			'	"<div class=""panel-content row"" >" & _
			'	"<div class=""form-row col-md-2""><label class=""ModTextLabel"">Address 2</label></div><div class=""form-row col-md-4""><input class=""ModText"" type=""text"" id=""Address2"" name=""Address2"" class=""form-control input-md"" value=""" & Trim(objRS("Address2")) & """></div>" & _
			'	"<div class=""form-row col-md-2""><label class=""ModTextLabel"">CardAddress2</label></div><div class=""form-row col-md-4""><input class=""ModTextAudit"" type=""text"" id=""CardAddress2"" name=""CardAddress2"" class=""form-control input-md"" value=""" & objRS("CardAddress2") & """></div></div>" & _
			'	"<div class=""panel-content row"" >" & _
			'	"<div class=""form-row col-md-2""><label class=""ModTextLabel"">Address 3</label></div><div class=""form-row col-md-4""><input class=""ModText"" type=""text"" id=""Address3"" name=""Address3"" class=""form-control input-md"" value=""" & objRS("Address3") & """></div>" & _
			'	"<div class=""form-row col-md-2""><label class=""ModTextLabel"">CardAddress3</label></div><div class=""form-row col-md-4""><input class=""ModTextAudit"" type=""text"" id=""CardAddress3"" name=""CardAddress3"" class=""form-control input-md"" value=""" & objRS("CardAddress3") & """></div></div>" & _
			'	"<div class=""panel-content row"" >" & _
			'	"<div class=""form-row col-md-2""><label class=""ModTextLabel"">Suburb/ State/ PostCode</label></div><div class=""form-row col-md-4""><input class=""ModText"" type=""text"" id=""SubStatPos"" name=""SubStatPos"" class=""form-control input-md"" value=""" & strSubStatPos & """></div>" & _
			'	"<div class=""form-row col-md-2""><label class=""ModTextLabel"">Card Suburb/ State/ PostCode</label></div><div class=""form-row col-md-4""><input class=""ModTextAudit"" type=""text"" id=""SubStatPosCard"" name=""SubStatPosCard"" class=""form-control input-md"" value=""" & strSubStatPosCard & """></div></div>" & _
			'	"<div class=""panel-content row"" >" & _
			'	"<div class=""form-row col-md-2""><label class=""ModTextLabel"">Home Phone</label></div><div class=""form-row col-md-4""><input class=""ModText"" type=""text"" id=""HomePhone"" name=""HomePhone"" class=""form-control input-md"" value=""" & objRS("HomePhone") & """></div>" & _
			'	"<div class=""form-row col-md-2""><label class=""ModTextLabel"">CardHomePhone</label></div><div class=""form-row col-md-4""><input class=""ModTextAudit"" type=""text"" id=""CardHomePhone"" name=""CardHomePhone"" class=""form-control input-md"" value=""" & objRS("CardHomePhone") & """></div></div>" & _
			'	"<div class=""panel-content row"" >" & _
			'	"<div class=""form-row col-md-2""><label class=""ModTextLabel"">Work Phone</label></div><div class=""form-row col-md-4""><input class=""ModText"" type=""text"" id=""WorkPhone"" name=""WorkPhone"" class=""form-control input-md"" value=""" & objRS("WorkPhone") & """></div>" & _
			'	"<div class=""form-row col-md-2""><label class=""ModTextLabel"">CardWorkPhone</label></div><div class=""form-row col-md-4""><input class=""ModTextAudit"" type=""text"" id=""CardWorkPhone"" name=""CardWorkPhone"" class=""form-control input-md"" value=""" & objRS("CardWorkPhone") & """></div></div>" & _
			'	"<div class=""panel-content row"" >" & _
			'	"<div class=""form-row col-md-2""><label class=""ModTextLabel"">Mobile Phone</label></div><div class=""form-row col-md-4""><input class=""ModText"" type=""text"" id=""MobilePhone"" name=""MobilePhone"" class=""form-control input-md"" value=""" & objRS("MobilePhone") & """></div>" & _
			'	"<div class=""form-row col-md-2""><label class=""ModTextLabel"">CardMobilePhone</label></div><div class=""form-row col-md-4""><input class=""ModTextAudit"" type=""text"" id=""CardMobilePhone"" name=""CardMobilePhone"" class=""form-control input-md"" value=""" & objRS("CardMobilePhone") & """></div></div>" & _
			'	"<div class=""panel-content row"" >" & _
			'	"<div class=""form-row col-md-2""><label class=""ModTextLabel"">Email</label></div><div class=""form-row col-md-4""><input class=""ModText"" type=""text"" id=""Email"" name=""Email"" class=""form-control input-md"" value=""" & Trim(objRS("Email")) & """></div>" & _
			'	"<div class=""form-row col-md-2""><label class=""ModTextLabel"">CardEmail</label></div><div class=""form-row col-md-4""><input class=""ModTextAudit"" type=""text"" id=""CardEmail"" name=""CardEmail"" class=""form-control input-md"" value=""" & objRS("CardEmail") & """></div></div>" & _
			'	"<div class=""panel-content row"" >" & _
			'	"<div class=""form-row col-md-2""><label class=""ModTextLabel"">Report Group</label></div><div class=""form-row col-md-4""><input class=""ModText"" type=""text"" id=""ReportGroup"" name=""ReportGroup"" class=""form-control input-md"" value=""" & objRS("ReportGroup") & """></div>" & _
			'	"<div class=""form-row col-md-2""><label class=""ModTextLabel"">Report Group</label></div><div class=""form-row col-md-4""><input class=""ModTextAudit"" type=""text"" id=""CardReportGroup"" name=""CardReportGroup"" class=""form-control input-md"" value=""" & objRS("CardReportGroup") & """></div></div>" & _
			'	"<div class=""panel-content row"" >" & _
			'	"<div class=""form-row col-md-2""><label class=""ModTextLabel"">Credit Limit</label></div><div class=""form-row col-md-4""><input class=""ModText"" type=""text"" id=""CreditLimit"" name=""CreditLimit"" class=""form-control input-md"" value=""" & objRS("CreditLimit") & """></div>" & _
			'	"<div class=""form-row col-md-2""><label class=""ModTextLabel"">Credit Limit</label></div><div class=""form-row col-md-4""><input class=""ModTextAudit"" type=""text"" id=""CardCreditLimit"" name=""CardCreditLimit"" class=""form-control input-md"" value=""" & objRS("CardCreditLimit") & """></div></div>" & _
			'	"<div class=""panel-content row"" >" & _
			'	"<div class=""form-row col-md-2""><label class=""ModTextLabel"">Relationship</label></div><div class=""form-row col-md-4""><input class=""ModText"" type=""text"" id=""Relationship"" name=""Relationship"" class=""form-control input-md"" value=""" & objRS("Relationship") & """></div>" & _
			'	"<div class=""form-row col-md-2""><label class=""ModTextLabel"">Card Relationship</label></div><div class=""form-row col-md-4""><input class=""ModTextAudit"" type=""text"" id=""CardRelationship"" name=""CardRelationship"" class=""form-control input-md"" value=""" & objRS("CardRelationship") & """></div></div>" & _
			'	"<div class=""panel-content row"" >" & _
			'	"<div class=""form-row col-md-2""><label class=""ModTextLabel"">Cat2</label></div><div class=""form-row col-md-4""><input class=""ModText"" type=""text"" id=""Cat2"" name=""Cat2"" class=""form-control input-md"" value=""" & objRS("Cat2") & """></div>" & _
			'	"<div class=""form-row col-md-2""><label class=""ModTextLabel"">CardCat2</label></div><div class=""form-row col-md-4""><input class=""ModTextAudit"" type=""text"" id=""CardCat2"" name=""CardCat2"" class=""form-control input-md"" value=""" & objRS("CardCat2") & """></div></div>" & _
			'	"<div class=""panel-content row"" >" & _
			'	"<div class=""form-row col-md-2""><label class=""ModTextLabel"">Account Number</label></div><div class=""form-row col-md-4""><input class=""ModText"" type=""text"" id=""AccountNumber"" name=""AccountNumber"" class=""form-control input-md"" value=""" & objRS("AccountNumber") & """></div>" & _
			'	"<div class=""form-row col-md-2""><label class=""ModTextLabel"">Account Number</label></div><div class=""form-row col-md-4""><input class=""ModTextAudit"" type=""text"" id=""CardAccountNumber"" name=""CardAccountNumber"" class=""form-control input-md"" value=""" & objRS("CardAccountNumber") & """></div></div>" & _
			'					"<div class=""panel-content row"" >" & _
			'	"<div class=""form-row col-md-2""><label class=""ModTextLabel"">ActivationFlag</label></div><div class=""form-row col-md-4""><input class=""ModText"" type=""text"" id=""ActivationFlag"" name=""ActivationFlag"" class=""form-control input-md"" value=""" & objRS("ActivationFlag") & """></div>" & _
			'	"<div class=""form-row col-md-2""><label class=""ModTextLabel"">ActivationFlag</label></div><div class=""form-row col-md-4""><input class=""ModTextAudit"" type=""text"" id=""CardActivationFlag"" name=""CardActivationFlag"" class=""form-control input-md"" value=""" & objRS("CardActivationFlag") & """></div></div>" & _
			'					"<div class=""panel-content row"" >" & _
			'	"<div class=""form-row col-md-2""><label class=""ModTextLabel"">PlasticID</label></div><div class=""form-row col-md-4""><input class=""ModText"" type=""text"" id=""PlasticID"" name=""PlasticID"" class=""form-control input-md"" value=""" & objRS("PlasticID") & """></div>" & _
			'	"<div class=""form-row col-md-2""><label class=""ModTextLabel"">PlasticID</label></div><div class=""form-row col-md-4""><input class=""ModTextAudit"" type=""text"" id=""CardPlasticID"" name=""CardPlasticID"" class=""form-control input-md"" value=""" & objRS("CardPlasticID") & """></div></div>" & _
			'					"<div class=""panel-content row"" >" & _
			'	"<div class=""form-row col-md-2""><label class=""ModTextLabel"">Companion</label></div><div class=""form-row col-md-4""><input class=""ModText"" type=""text"" id=""Companion"" name=""Companion"" class=""form-control input-md"" value=""" & MaskCard(objRS("Companion")) & """></div>" & _
			'	"<div class=""form-row col-md-2""><label class=""ModTextLabel"">Companion</label></div><div class=""form-row col-md-4""><input class=""ModTextAudit"" type=""text"" id=""Companion"" name=""Companion"" class=""form-control input-md"" value=""" & MaskCard(objRS("Companion")) & """></div></div>" & _
			'					"<div class=""panel-content row"" >" & _
			'	"<div class=""form-row col-md-2""><label class=""ModTextLabel"">Status</label></div><div class=""form-row col-md-4""><input class=""ModText"" type=""text"" id=""Status"" name=""Status"" class=""form-control input-md"" value=""" & objRS("Status") & """></div>" & _
			'	"<div class=""form-row col-md-2""><label class=""ModTextLabel""></label></div><div class=""form-row col-md-4""></div></div>"
			
		Else
			Response.write "No Record for " & Request.QueryString("CSFromDinersID")
	   End If

	objRS.Close
	
	Response.write "</table>"
	
	
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