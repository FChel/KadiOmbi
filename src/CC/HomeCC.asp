<%@ Language=VBScript %>
<% Option Explicit
	
	Response.Expires = -1500

   'If IsEmpty(Session("Logon")) Then Response.Redirect("AccessDenied.asp")
   If IsEmpty(Session("EmployeeID")) Then Response.Redirect("../Timeout.asp")

	Session("CurrentPage") = "CC/HomeCC.asp"
	
Dim objCon
Dim objRS

Set objCon = Server.CreateObject("ADODB.Connection")
Set objRS = Server.CreateObject("ADODB.Recordset")

objCon.Open Session("DBConnection")

     
 %>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="../BERTStyle.css">

 <meta charset="utf-8">
  <meta http-equiv="X-UA-Compatible" content="IE=edge">
  <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
  <meta name="description" content="">
  <meta name="author" content="">
  <title>Cards Home</title>
  <!-- Bootstrap core CSS-->
  <link href="../vendor/bootstrap/css/bootstrap.min.css" rel="stylesheet">
  <!-- Custom fonts for this template-->
  <link href="../vendor/font-awesome/css/font-awesome.min.css" rel="stylesheet" type="text/css">
  <!-- Custom styles for this template-->
  <link href="../css/sb-admin.css" rel="stylesheet">

  
</HEAD>
<BODY>
<FORM action="Home.asp" method="POST" id="frm" name="frm">

<div class="container-fluid">
     
      
      <!-- Icon Cards-->
      <div class="row">
        <div class="col-xl-3 col-sm-6 mb-3">
          <div class="card text-white bg-primary o-hidden h-100px">
            <div class="card-body">
              <div class="card-body-icon">
                <i class="fa fa-fw fa-comments"></i>
              </div>
              <div class="mr-5">26 New Messages!</div>
            </div>
            <a class="card-footer text-white clearfix small z-1" href="#">
              <span class="float-left">View Details</span>
              <span class="float-right">
                <i class="fa fa-angle-right"></i>
              </span>
            </a>
          </div>
        </div>
        <div class="col-xl-3 col-sm-6 mb-3">
          <div class="card text-white bg-secondary o-hidden h-100px">
            <div class="card-body">
              <div class="card-body-icon">
                <i class="fa fa-fw fa-list"></i>
              </div>
              <div class="mr-5">11 New Account Transfers!</div>
            </div>
            <a class="card-footer text-white clearfix small z-1" href="#">
              <span class="float-left">View Details</span>
              <span class="float-right">
                <i class="fa fa-angle-right"></i>
              </span>
            </a>
          </div>
        </div>
        <div class="col-xl-3 col-sm-6 mb-3">
          <%Call LoadApplications()%>
            
            <a class="card-footer text-white clearfix small z-1" href="Applications3.asp">
              <span class="float-left">View Details</span>
              <span class="float-right">
                <i class="fa fa-angle-right"></i>
              </span>
            </a>
          </div>
        </div>
        <div class="col-xl-3 col-sm-6 mb-3">
          <div class="card text-white bg-dark o-hidden h-100px">
            <div class="card-body">
              <div class="card-body-icon">
                <i class="fa fa-fw fa-cogs"></i>
              </div>
              <div class="mr-5">13 New Card Updates!</div>
            </div>
            <a class="card-footer text-white clearfix small z-1" href="Applications3.asp">
              <span class="float-left">View Details</span>
              <span class="float-right">
                <i class="fa fa-angle-right"></i>
              </span>
            </a>
          </div>
        </div>
      </div>
	 </div>



<HR>
<TABLE Align="Center" WIDTH="100%" BORDER="1" CELLSPACING="1" CELLPADDING="1">	

	<TR>
		<TH Style="Width:50%;Height:25px; text-align:center;">Dashboard</TH><TH Width="50%" Style="Width:50%;Height:25px; text-align:center;">Applications</TH>
	</TR>
	<TR>
		<TD Colspan="2">&nbsp;</TD> 
	</TR>
	<tr>
		<TD ><iframe id="Iframe1" name="framecontent" src="<%=Session("HomePage1")%>" Width="100%" frameborder="0" height="500px"></iframe></TD>
		<TD ><iframe id="framecontent" name="framecontent" src="<%=Session("HomePage2")%>" Width="100%" frameborder="0" height="500px"></iframe></TD>
	</TR>
	
   <tr><th Style="Height:25px;"  colspan="2">&nbsp;</th></tr>
</TABLE>

</BODY>
</HTML>

<%
Sub LoadApplications()

Dim strApplications
Dim strAppFormat
Dim strText

   'Description:	Loads summary data at the top of the Home page in the card format
	objRS.Open "SELECT * FROM qryApplicationSummary",objCon
   
		If Not objRS.EOF Then
			Do Until objRS.EOF
				strApplications = strApplications + objRS("Applications")
			
			objRS.MoveNext
			Loop
		Else
			strApplications = 0
		End If
	objRS.Close
	
	If strApplications > 20 Then
		strAppFormat = "secondary"
	ElseIf strApplications > 40 Then
		strAppFormat = "dark"
	ElseIf strApplications > 60 Then
		strAppFormat = "warning"
	ElseIf strApplications > 80 Then
		strAppFormat = "danger"
	Else
		strAppFormat = "success"
	End If
		
	If strApplications = 1 Then
		strText = "Application!"
	Else
		strText = "Applications!"
	End If
		
	Response.Write "<div class=""card text-white bg-" & strAppFormat & " o-hidden h-100px""><div class=""card-body"">" & _
              "<div class=""card-body-icon""><i class=""fa fa-fw fa-address-card""></i>" & _
              "</div><div class=""mr-5"">" & strApplications & " " & strText & "</div></div>"
End Sub

Set objRS = Nothing

%>
