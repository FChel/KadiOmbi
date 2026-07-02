<!-- #Include file=../../ADOVBS.inc -->
<!-- #Include file=../CAPSFunctions.asp -->

<html lang="en">
<%
Dim objCon
Dim objRS
Dim x
Dim strWherePM
Dim strSQL
Dim objCmd

'Open Database Connection
Set objCon = Server.CreateObject("ADODB.Connection")
Set objRS = Server.CreateObject("ADODB.Recordset")
Set objCmd = Server.CreateObject("ADODB.Command")

    objCon.Open Session("DBConnection")


If Not IsEmpty(Request.QueryString("Where")) Then
	strWherePM = "" & Request.QueryString("Where") & ""
	Else 
	strWherePM = ""
End If

'Not needed, there is already a WHERE statement in the sproc so we need the AND
'If strWherePM <> "" Then
'	strWherePM = "'" + Replace(strWherePM, "AND","WHERE",1,1) + "'"
'End If

Public Sub AddAllRecords(strWherePM)

Dim intRecord

  		With objCmd
			.CommandType = 4
			.CommandText = "spCAPSAddAllPMReco"

			.Parameters.Append objCmd.CreateParameter("Where", adVarChar, adParamInput,100)
			.Parameters.Append objCmd.CreateParameter("AddAllPMRecoOutput", adInteger, adParamOutput)
			
			.Parameters("Where") = strWherePM
			
			.ActiveConnection = objCon
			 
		End With
	   
		objCmd.Execute        
	  
		'Return the result of the Save Function.
		intRecord = objCmd.Parameters.Item("AddAllPMRecoOutput") 
	 	
		'Response.Write intRecord
		If intRecord = 0 Then
				Response.Write "<div class=""alert alert-warning"" role=""alert"">No cards had their PM Load Status updated. If you believe this is an error, contact System Admin and quote ""PMReconciliationExisting error on strWhere: " & strWherePM & " ""</div>"
		ElseIf intRecord > 0 Then
				Response.Write "<div class=""alert alert-success"" role=""alert"">PM Load Status for " & intRecord & " cards updated successfully.</div>"
		Else
				Response.Write "<div class=""alert alert-danger"" role=""alert"">An error occurred. Please ensure you have selected from the comparison drop-down (Address1, Address2, ect) and try again. If the error persists, contact System Admin and quote ""PMReconciliationExisting error on strWhere: " & strWherePM & " ""</div>"
		End If
		
	
End Sub
%>
<head>
  
</head>

<body>
  <%
  
  	'Call the same procedure for New CS To Diners records
	Call AddAllRecords(strWherePM)

  %>
</body>

</html>
<%

Set objRS = Nothing
Set objCon = Nothing
  
  %>