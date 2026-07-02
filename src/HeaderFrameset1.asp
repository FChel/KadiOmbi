<%@ Language=VBScript %>
<%

Dim objCon
Dim objRS
Dim intNavColourID
Dim x

Set objCon = Server.CreateObject("ADODB.Connection")
Set objRS = Server.CreateObject("ADODB.Recordset")

    objCon.Open Session("DBConnection")

    objRS.Open "SELECT NavColourID FROM tblScreens WHERE PageName = '" & Session("CurrentPage") & "'",objCon

        If Not objRS.EOF Then
            intNavColourID = objRS(0)
        Else
            intNavColourID = 0
        End If

    objRS.Close


    For x = 1 to 12

        If intNavColourID = x Then
            Session("NavCol" & x & "") = "Yellow"
        Else
            Session("NavCol" & x & "")  = "White"
        End If


    Next




%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="BERTStyle.css">
<title>Isidore</title>
</head>

<frameset FrameBorder="0" COLS="*%" >
	
	<frame NAME="Body2" SCROLLING="no" Border="0" SRC="<%=Session("HeaderMenu")%>">

</frameset>

<noframes>
<body>
<b>Isidore IT<p>
Sorry your browser doesn't support frames.
</b></body>
</noframes>
</html>
