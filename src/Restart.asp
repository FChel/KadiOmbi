<%@ Language=VBScript %>

<%
Sub Touch(strFolderPath, strFileName, strNewDate)

	Set app = Server.CreateObject("Shell.Application")
	Set strFolder = app.NameSpace(strfolderPath)
	Set strFile = strFolder.ParseName(strfileName)

	strFile.ModifyDate = NewDate

	Set strFile = nothing
	Set strFolder = nothing
	Set app = nothing
	
End Sub

Call Touch(Server.MapPath("/"), "web.config", now)

Response.Write "Restarted"

%>
<html>
<head>
<meta name="robots" content="noindex, nofollow">
</head>

</html>