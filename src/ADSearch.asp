<%@ Language=VBScript %>
<%
'SearchAD.vbs
'On Error Resume Next
' Connect to the LDAP server's root object


' This code prints the attributes of the RootDSE
set objRootDSE = GetObject("LDAP://RootDSE")
objRootDSE.GetInfo
for i = 0 to objRootDSE.PropertyCount - 1 
    set strProp = objRootDSE.Item(i)
    Response.Write strProp.Name & " " & "<BR>"
    for each strPropval in strProp.Values
       Response.Write "  " &  strPropval.CaseIgnoreString & "<BR>"
    next
next


Response.Write "XXX"
Set objRootDSE = GetObject("LDAP://RootDSE")
strDNSDomain = objRootDSE.Get("defaultNamingContext")
strTarget = "LDAP://" & strDNSDomain
Response.Write "Starting search from <BR>" & strTarget & "<BR>"

' Connect to Ad Provider
Set objConnection = CreateObject("ADODB.Connection")
objConnection.Provider = "ADsDSOObject"
objConnection.Properties("User ID") = "dpesit\svc_CAPS_IIS_UAT"
'objConnection.Properties("Password") = "Xw!q^In=,T9.f?Tr)>8.NP8$Z13}63"
'objConnection.Properties("User ID") = "dpe\svc_CAPS_IIS_PROD"
objConnection.Properties("Password") = "P|oS#Z?zGhdgeyaKq%8}fi5@giQL?<"

'objConnection.Properties("User ID") = "dpesit\andrew.bull3_priv"
'objConnection.Properties("Password") = "DJHPJa#22052013?"

objConnection.Open "Active Directory Provider"

Set objCmd =   CreateObject("ADODB.Command")
Set objCmd.ActiveConnection = objConnection 

' Show only computers
'objCmd.CommandText = "SELECT Name, ADsPath FROM '" & strTarget & "' WHERE objectCategory = 'computer'"

'Above Replaced with dynamic removal of the Domain -- March 2022

Dim strLogin

strLogin = Request.ServerVariables("Auth_User")
			If InStr(1,strLogin,"\")> 0 Then
				strLogin = Right(strLogin,Len(strLogin)-InStr(1,strLogin,"\"))
			End If

' or show only users
objCmd.CommandText = "SELECT sAMAccountName,userPrincipalName,EmployeeID, Name, ADsPath,mail FROM '" & strTarget & "' WHERE objectCategory = 'user' AND sAMAccountName = '" & strLogin & "'"

Response.Write "<br>SELECT sAMAccountName,userPrincipalName,EmployeeID, Name, ADsPath,mail FROM '" & strTarget & "' WHERE objectCategory = 'user' AND sAMAccountName = '" & strLogin & "'<br>"

' or show only groups
'objCmd.CommandText = "SELECT Name, ADsPath FROM '" & strTarget & "' WHERE objectCategory = 'group'"

''
Const ADS_SCOPE_SUBTREE = 2
objCmd.Properties("Page Size") = 100
objCmd.Properties("Timeout") = 30
objCmd.Properties("Searchscope") = ADS_SCOPE_SUBTREE
objCmd.Properties("Cache Results") = False

Set objRecordSet = objCmd.Execute

' Iterate through the results
objRecordSet.MoveFirst
Do Until objRecordSet.EOF
 
   Response.Write "EmployeeID = " & objRecordSet.Fields("EmployeeID") & "<BR>"
   Response.Write "ADsPath = " & objRecordSet.Fields("ADsPath") & "<BR>"
    Response.Write "Email = " & objRecordSet.Fields("mail") & "<BR>"
   objRecordSet.MoveNext
Loop
%>