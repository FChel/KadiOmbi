<!-- #Include file=../../CC/CAPSHeader.asp -->
<!-- #Include file=../../ADOVBS.inc -->
<!-- #include file="../../CC/CAPSFunctions.asp" -->
<%
'Description:	Check IIS log files admin screen for viewing errors
'Author:		MG
'Date:			April 2022

    'If the session has expired then send the user back to the Default/login page
	If IsEmpty(Session("UserID")) Then Response.Redirect("../../Timeout.asp")
	
'Instantiate Common Page Variables.
Dim objCon
Dim objRS

Dim strDeleteCheck
Dim dteBatchDate
Dim errors

'Open Database Connection
Set objCon = Server.CreateObject("ADODB.Connection")
Set objRS = Server.CreateObject("ADODB.Recordset")

objCon.Open Session("DBConnection")
'Session("IISLogSelected") = ""
'If a log file is selcted then set the Session value to the one selected
If Not IsEmpty(Request.QueryString("IISLogSelected")) Then

	Session("IISLogSelected") = Request.QueryString("IISLogSelected")
	
	If Session("IISLogSelected") = "0" Then Session("IISLogSelected") = ""
	
End If


 %>

<html>
<head>

	
	
<script language=javascript>

</script>

</head>
<body>

<!-- Modal -->
	<div class="loader" id="ModApp">
        <div class="wrap">
            <div class="spinner"></div>
            <span class="loading-message">Loading...</h6>
        </div>
    </div>
	
<main class="main py-3">
      <div class="container">
<form action="UploadTraining.asp?Action=Save&chkDelete="+frm.chkDelete.value  method="POST" enctype="multipart/form-data" id="frm" name="frm">

<div class="content-wrapper">
    <div class="container-fluid">
	 
	 <div class="row" id="basic-table">
  <div class="col-4">
    <div class="card">
     <div class="card-header">
          <h4 class="card-title"><img src="../../CC/img/DFG_Logo.png" height="40px" width="100px" title="IIS Log files on Server"> IIS Log files</h4>
        </div>
      <div class="card-content">
        <div class="card-body">
		
		<%
        
     DisplayFileSummary()
        
%>	

		</div>
	  </div>
    </div>
   </div>
  

  <div class="col-8">
    <div class="card">
     
      <div class="card-content">
        <div class="card-body">
		
		<%
        
     ReadText()
        
%>	
		</div>
	  </div>
    </div>
   </div>
   
  
    </div>
</div>
</form>
      </div>
    </main>


<!-- #Include file=../../CC/CAPSFooter.asp -->
</body>
</html>

<%

Public Sub DisplayFileSummary()


Dim objStartFolder
Dim colFiles
Dim strFile
Dim intCount
Dim objFSO
Dim objFolder
Dim objFile
Dim strFileSize
Dim strFileAttributes
Dim strLogFolder
Dim strIISFullFilePath
Dim strIISSelected
Dim strStyle
Dim strFileAge

Set objFSO = CreateObject("Scripting.FileSystemObject")

'Get the System Parameter for the folder location of the IISLogFiles
strLogFolder = GetSystemAdmin("IISLogFiles")

'Get the selected file if one is selected
If IsNull(Session("IISLogSelected")) or Session("IISLogSelected") = "" Then
	strIISSelected = ""
Else
	strIISSelected = Session("IISLogSelected")
End If

	'Use the Starting folder from System Parameters and add the Training folder
	objStartFolder = strLogFolder

	Set objFolder = objFSO.GetFolder(objStartFolder)
		
	Set colFiles = objFolder.Files
	
	Response.Write"<table Class=""table table-bordered mb-0 newd"" cellspacing=""0"" cellpadding=""0"">" & _
			"<tr><th style=""text-align:left"">IIS Log Files <i class=""fa help-tooltip fa-question-circle"" data-toggle=""tooltip"" title=""IIS Log files on the server showing the most recent""></i></th></tr>"
	
	If strIISSelected = "" Then
		Response.Write "<TR><TD Title=""Clear Selected file""><a href=""CheckLogs.asp?IISLogSelected=0"">Clear Selected File</a></TD></TR>"
	End If
	
	intCount = 0
	
	For Each objFile in colFiles
	
		If Right(objFile.Name,3) = "log" Then
	
			strFileAge = DateDiff("d",objFile.DateCreated,now())
			If strFileAge <20 Then
				
				intCount = intCount + 1
			
				If intCount < 20 Then
					If IsNull(objFile.Name) or objFile.Name = "" Then
						strFile = ""
						strFileSize = 0
					Else
						strFile = objFile.Name
						strIISFullFilePath = objStartFolder & "\" & strFile
						'strFile = Left(objFile.Name,10) & ".." & Right(objFile.Name,4)
						strFileSize = Round(objFile.Size/1024000,2)
						
						strFileAttributes =  "Created: " & objFile.DateCreated
						strFileAttributes = strFileAttributes & " Last Accessed: " & objFile.DateLastAccessed
						strFileAttributes = strFileAttributes & " Last Modified: " & objFile.DateLastModified  
		
						If strIISSelected = strIISFullFilePath Then
							strStyle = " Style=""background-color:#e6e6e6;"" "
						Else
							strStyle = ""
						End If
						
					End If
					
					Response.Write "<TR><TD " & strStyle & " Title=""" & objFile.Name & " " & strFileAttributes & """><a href=""CheckLogs.asp?IISLogSelected=" & strIISFullFilePath & """>" & strFile & " - (" & strFileAge & " Days old)</a></TD></TR>"
				End If
				
			End If
		End If
	Next
	

	
	 Response.Write "<tr><th style=""text-align:left"">Total: " & intCount & "</th></tr></table>"


End Sub



Sub ReadText()

Const ForReading = 1
Dim strLine
Dim strCardType
Dim strRow
Dim x, y

Dim strFooterCount
Dim lngFileLoadID
Dim objFSO
Dim objStartFolder
Dim objFolder
Dim strFileNameDefault
Dim filePath
Dim colFiles
Dim objFile
Dim objTextFile
Dim strFileDateTime
Dim strFileSeqNum
Dim lngFileID

Set objFSO = CreateObject("Scripting.FileSystemObject")

	'''---Start the New Service Account Login section
	If IsNull(Request.QueryString("IISLogSelected")) OR Request.QueryString("IISLogSelected")="" Then
	
		Response.write "<div class=""content-body""><section id=""basic-alerts""><div class=""alert alert-danger alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
			"<span aria-hidden=""true"">&times;</span></button>" & _
			"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
			"<span>No IIS Log file currently selected.  Click on a Log file in the summary area to select</span></div></div></div>"
		
		Exit Sub
	Else			
		filePath = Request.QueryString("IISLogSelected")
	End If

	If IsNull(filePath) or filePath = "" Then

		Response.write "<div class=""content-body""><section id=""basic-alerts""><div class=""alert alert-danger alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
				"<span aria-hidden=""true"">&times;</span></button>" & _
				"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
				"<span>No IIS Log file exists for filePath: " & filePath & " selected.  See System Admin</span></div></div></div>"
			Exit Sub
	End If


Set objTextFile = objFSO.OpenTextFile (filePath, ForReading)
'Set objTextFile = objFSO.OpenTextFile ("c:\mytextfile.txt", ForReading)

x = 0 
 
	Do Until objTextFile.AtEndOfStream
		
		'Count the rows for use in line counts, summary and for getting header
		x = x + 1
					
		strLine = objTextFile.Readline
		
'Replace any Javascript Script tags so Javascript does not execute
		strLine = Replace(UCASE(strLine),"<SCRIPT>","**--js script tag << removed >> --**")

		'Highlight the errors
		strLine = Replace(strLine,"|800","<span style=""color:red; font-weight:bold;"">|800</span>")

		'Add an error line to identify an error below
		If Instr(1,strLine,"|800")>0 Then Response.Write "<span style=""color:red; font-weight:bold;""><----- ERROR BELOW ------></span></BR>"

		Response.Write strLine & "</BR></BR>"
		
		
	Loop

	
Set objFolder = Nothing
Set colFiles = Nothing
Set objFSO = Nothing


End Sub


Set objRS = Nothing
Set objCon = Nothing

 %>


