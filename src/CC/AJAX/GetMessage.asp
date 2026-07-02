
<%

'Procedure to load any messages relating to the card
Dim strSQL
Dim strPerson
Dim strMessage
Dim strEdit
Dim lngMessageID
Dim objRS
Dim objCon

Set objCon = Server.CreateObject("ADODB.Connection")
Set objRS = Server.CreateObject("ADODB.Recordset")

'Open database connection
objCon.Open Session("DBConnection")


	If Not IsEmpty(Request.QueryString("MessageID")) Then
		lngMessageID = Request.QueryString("MessageID")
	End If
	
	strSQL = "SELECT * FROM qryCAPSMessage WITH(NOLOCK) WHERE [MessageID] = " & lngMessageID & ""

	objRS.Open strSQL,objCon
	
    Do Until objRS.EOF
		
		If IsNull(objRS("MessageFrom") ) or objRS("MessageFrom")  = "" Then
			strPerson = ""
		Else
			If objRS("MessageFrom") = Session("UserID") Then
				strPerson = "(You)"
				'strEdit = " data-toggle=""modal"" data-target=""#NotesModal"" "
			Else
				If IsNull(objRS("MessageFrom") ) or objRS("MessageFrom")  = "" Then
					If objRS("MessageFrom") = 0 Then
						strPerson = "(Admin)"
					End If
				End If
				'strEdit = " data-toggle=""modal"" data-target=""#NotesModal"" "
			End If
		End If
		
		If IsNull(objRS("MessageDetail")) or objRS("MessageDetail")= "" Then
			strMessage=""
		Else
			strMessage = objRS("MessageDetail")
			'strMessage = Replace(strMessage,chr(13),"</BR>")
		End If
		
		Response.write "<div class=""panel panel-light col-12""><div class=""panel-header"">" & _
			"<h6 " & strEdit & ">" & objRS("UserFrom") & " " & strPerson & "</h6>" & _
			"<textarea rows=""4"" id=""NTNewNotes"" name=""NTNewNotes"" class=""form-control input-md"" >" & strMessage & "</textarea>" & _
			"<input type=""text"" id=""NTNotesID"" name=""NTNotesID"" value=""" & lngMessageID & """></div></div>"

		objRS.Movenext
	Loop
				
objRS.Close

Set objRS = Nothing
Set objCon = Nothing
  
  %>