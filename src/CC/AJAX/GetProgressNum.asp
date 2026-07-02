
<%

Response.write Session("CDMCLoadProgress")

Dim objCon
Dim objRS
Dim x
Dim strBatchNo

'Open Database Connection
Set objCon = Server.CreateObject("ADODB.Connection")
    
    objCon.Open Session("DBConnection")

Set objRS = Server.CreateObject("ADODB.Recordset")

	'Description:	Writes the Number of records loaded into the CDMC table
	objRS.Open "SELECT COUNT([CDMCID]) AS ProgressNum FROM tblCAPSCDMC WITH(NOLOCK)",objCon
	  
		If Not objRS.EOF Then
		   
			Response.Write objRS("ProgressNum")
			
		End If

	objRS.Close

Set objRS = Nothing
Set objCon = Nothing
  
  %>