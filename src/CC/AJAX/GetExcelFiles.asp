<!-- #Include file=../CAPSFunctions.asp -->
<%

Dim objStartFolder
Dim colFiles
Dim strFile
Dim intCount
Dim objFSO
Dim objFolder
Dim objFile
Dim strFileExtension
Dim strServer
Dim strServerPath

Dim strFileDelete

Set objFSO = CreateObject("Scripting.FileSystemObject")

	strServer = GetSystemAdmin("ServerFilePath")
	strServerPath = GetSystemAdmin("ServerPath")
	
	'If this page was called from the Training Report then use that Report URL to Get Excel Files
	If IsNull(Request.QueryString("Training")) Then
		strFileDelete = strServer & "\Admin\CAPSAdmin\Attachments\Reports\"
	Else
		strFileDelete = strServer & "\Admin\CAPSAdmin\Attachments\Training\Reports\"
	End If
	
	'objStartFolder = objFSO.GetAbsolutePathName("D:\Inetpub\WWWRoot\CAPS\ASPNew\Admin\CAPSAdmin\Attachments\Training\Reports\")
	objStartFolder = objFSO.GetAbsolutePathName(strFileDelete)
	
		'strServer = Server.MapPath(GetFilePath()) & "\Attachments\Training\"
		'strServer = GetSystemAdmin("ServerFilePath") & "\Admin\CAPSAdmin\Attachments\Training\"
		'strFileDelete = Server.MapPath(GetFilePath()) & "\Admin\CAPSAdmin\Attachments\Training\"
		strFileDelete = "D:\Inetpub\WWWRoot\CAPS\ASPNew\Admin\CAPSAdmin\Attachments\Training\Reports\"
		'strServer = "http:\\VBMRSN05\CAPS\Admin\CAPSAdmin\Attachments\Training\Reports\"
		
		
		'Add in the local path from the server path name setting
		strServer = strServer & "\Admin\CAPSAdmin\Attachments\Training\Reports\"
		strServerPath = strServerPath & "\Admin\CAPSAdmin\Attachments\Training\Reports\"
		
	Set objFolder = objFSO.GetFolder(objStartFolder)
		
	Set colFiles = objFolder.Files
	
	Response.Write"<table Class=""table table-bordered mb-0 newd"" cellspacing=""0"" cellpadding=""0"">" & _
			"<tr><th style=""text-align:left"" colspan=""2"">Excel Training Files <i class=""fa help-tooltip fa-question-circle"" data-toggle=""tooltip"" title=""Training Excel Files on the CAPS Server waiting to be downloaded. Click to download.""></i></th></tr>"
			
	For Each objFile in colFiles

		intCount = intCount + 1
		
		'If intCount < 6 Then
			If IsNull(objFile.Name) or objFile.Name = "" Then
				strFile = ""
			Else
				strFile = objFile.Name'Left(objFile.Name,10)
			End If
			
			'Display the correct file extension
			'strFileExtension = Right(objFile.Name,3)
			
			Response.Write "<TR><TD><a href=""" & strServerPath & strFile & """>" & strFile & "</a></TD>" & _
						"<td><button type=""button"" class=""btn btn-outline-danger"" onClick=""self.location.href='Training.asp?Link=RP&Action=DeleteExcelFile&File=" & strFile & "';""><i class=""fa fa-times""></i> Delete</button></td></TR>"
			
		'End If
		
	Next
	
	 Response.Write "<tr><th style=""text-align:left"" colspan=""2"">Total: " & intCount & "</th></tr></table>"
	 
Set objFSO = Nothing

  
  %>

