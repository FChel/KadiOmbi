<!--#include file="adovbs.inc"-->
<%
Function BytesToStr(bytes)
    Dim Stream
    Set Stream = Server.CreateObject("Adodb.Stream")
    With Stream
        .Type = 1 'adTypeBinary
        .Open
        .Write bytes
        .Position = 0
        .Type = 2 'adTypeText
        .Charset = "UTF-8"
        BytesToStr = .ReadText
        Stream.Close
    End With
    Set Stream = Nothing
End Function

Dim isPost: isPost = (UCase(Request.ServerVariables("REQUEST_METHOD") & "") = "POST")
Dim lngBytesCount, xmlString
Dim xdoc, fs, tfile
Dim loaded
Dim objCon
	Dim x
	
'Set Database Connection
Set objCon=Server.CreateObject("ADODB.CONNECTION")

Session("DBConnection") = "File Name=" & Server.MapPath("Database/CAPS.udl") & ";"
		
objCon.Open Session("DBConnection")

'Is it a HTTP POST?
If isPost Then

    If Request.TotalBytes > 0 Then
	
        lngBytesCount = Request.TotalBytes
        xmlString = BytesToStr(Request.BinaryRead(lngBytesCount))

        Set xdoc = CreateObject("Msxml2.DOMDocument.6.0")
        xdoc.async = False
        loaded = xdoc.loadXML(xmlString)
		
        If loaded Then 
		
			Set objRoot = xdoc.documentElement

			If IsObject(objRoot) = False Then

				Response.Write "There was an error parsing xml"

			Else
			
					For Each objLevel1 in objRoot.ChildNodes
						nodeName1 = objLevel1.NodeName
						nodeText1 = objLevel1.Text
						child1 = objLevel1.childNodes.length
						
						If child1<2 Then
							'response.Write(nodename1 & " = " & nodeText1 & "<br>")
							strSaveString = "," & strSaveString & nodename1 & " = " & nodeText1

							x = x + 1
							strSaveString2 = "," & strSaveString2 & x & " = " & nodeText1
							
							y = y + 1
							Call SaveXMLTemp (nodename1,nodeText1)
						Else
							'response.Write(nodename1 & "<br>")
							strSaveString = "," & strSaveString & nodename1
							
							'x = x + 1
							'strSaveString2 = "," & strSaveString2 & x
						ENd If
						
						'second level
						If child1>1 Then
							For Each objLevel2 in objLevel1.ChildNodes
								nodeName2 = objLevel2.NodeName
								nodeText2 = objLevel2.Text
								child2 = objLevel2.childNodes.length
								
								If child2<2 Then
									'response.Write("&nbsp;&nbsp;&nbsp;" & nodeName1 & "_" & nodename2 & " = " & nodeText2 & "<br>")
									''''' OLD DELETE 'response.Write("&nbsp;&nbsp;&nbsp;" &  nodename2 & " = " & nodeText2 & "<br>")
									strSaveString = "," & strSaveString & nodeName1 & "_" & nodename2 & " = " & nodeText2
									
									x = x + 1
									strSaveString2 = "," & strSaveString2 & x & " = " & nodeText2
									
									y = y + 1
									Call SaveXMLTemp (nodeName1 & "_" & nodename2,nodeText2)
								Else
									'response.Write("&nbsp;&nbsp;&nbsp;" & nodeName1 & "_" & nodename2 & "<br>")
									''''' OLD DELETE 'response.Write("&nbsp;&nbsp;&nbsp;" &  nodename2 & "<br>")
									strSaveString = "," & strSaveString & nodeName1 & "_" & nodename2
								End If
								
								'third level
								If child2>1 Then
									For Each objLevel3 in objLevel2.ChildNodes
										nodeName3 = objLevel3.NodeName
										nodeText3 = objLevel3.Text
										child3 = objLevel3.childNodes.length
										
										If child3<2 Then
											'response.Write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & nodeName1 & "_" & nodename2 & "_" & nodename3 & " = " & nodeText3 & "<br>")
											''''' OLD DELETE 'response.Write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & nodename3 & " = " & nodeText3 & "<br>")
											strSaveString = "," & strSaveString & nodeName1 & "_" & nodename2 & "_" & nodename3 & " = " & nodeText3
											
											x = x + 1
											strSaveString2 = "," & strSaveString2 & x & " = " & nodeText3
											
											y = y + 1
											Call SaveXMLTemp (nodeName1 & "_" & nodename2 & "_" & nodename3,nodeText3)
									
										Else
											'response.Write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & nodeName1 & "_" & nodename2 & "_" & nodename3 & "<br>")
											''''' OLD DELETE response.Write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & nodename3 & "<br>")
											strSaveString = "," & strSaveString & nodeName1 & "_" & nodename2 & "_" & nodename3
										End If
										
										
											'fourth level
											If child3>1 Then
												For Each objLevel4 in objLevel3.ChildNodes
													nodeName4 = objLevel4.NodeName
													nodeText4 = objLevel4.Text
													child4 = objLevel4.childNodes.length
													
													If child4<2 Then
														'response.Write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" &  nodename4 & " = " & nodeText3 & "<br>")
														strSaveString = "," & strSaveString & nodename4 & " = " & nodeText3
														
														x = x + 1
														strSaveString2 = "," & strSaveString2 & x & " = " & nodeText3
												
														y = y + 1
														Call SaveXMLTemp (nodename4,nodeText3)
													Else
														'response.Write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" &  nodename4 & "<br>")
														strSaveString = "," & strSaveString & nodename4
													End If
												Next
											End If
									Next
								End If
										
							Next
						End If
						
					Next


			End If
	
			Set objXML = Nothing        	
			
            'Let sender know we have received and processing the message.
            Response.Status = "200 OK"
            Response.Write "200 OK"
			
        Else
		
            Response.Status = "400 Bad Request"
            Response.Write xdoc.parseError.errorCode & " - " & xdoc.parseError.Reason
			
        End If
    Else
	
      Response.Status = "400 Bad Request"
      Response.Write "No message was sent"
	  
    End If
	
Else

  'Return method not allowed
  Response.Status = "405 Method Not Allowed"
  Response.Write "Requested method is not supported."
  
End If
Response.End

Sub insertdata(nname,ntext)
	Set objCmd2 = Server.CreateObject("ADODB.Command")
	Set objCmd2.ActiveConnection = objCon
	objCmd2.CommandText = "spCAPSXMLTempSave"
	objCmd2.CommandType = adobjCmd2StoredProc
	objCmd2.Parameters.Append objCmd2.CreateParameter("@NodeName",advarchar,adParamInput,200,nname ) 
	objCmd2.Parameters.Append objCmd2.CreateParameter("@NodeText",advarchar,adParamInput,200,ntext )
	objCmd2.Execute     
End Sub


Public Sub SaveXMLTemp(strNodeName, strNodeText)
'Procedure to run a stored procedure which updates the summary details for a file just loaded, which is used where summary details are displayed

	Set objCmd2 = Server.CreateObject("ADODB.Command")
			Set objCmd2.ActiveConnection = objCon

'If there is no node name (data) then exit the procedure
If IsNull(strNodeName) Then Exit Sub

If IsNull(strNodeText) Then strNodeText = ""

'response.write y & " "
  	With objCmd2
  	
	'If y = 1 then
	
		.CommandType = 4
		.CommandText = "spCAPSXMLTempSave"
		
		.Parameters.Append objCmd2.CreateParameter("NodeName", adVarchar, adParamInput,200)
		.Parameters.Append objCmd2.CreateParameter("NodeText", adVarchar, adParamInput,200)
	'End if
	
		.Parameters("NodeName") = strNodeName
		.Parameters("NodeText") = strNodeText

		.ActiveConnection = objCon
                
    End With
                
	objCmd2.Execute        
	
		
'Set ObjCmd2 = Nothing

End Sub

%>