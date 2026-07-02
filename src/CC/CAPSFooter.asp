<% 

Dim strErrorMessage
Dim bolErrors
Dim objError
Set objError = Server.GetLastError()

'Initialise variables
strErrorMessage = ""
bolErrors = False

If Err.Number <> 0 Then

	'TrapError Err.Description & " : " & Err.File & " : " & Err.Line & " : "
	'TrapError objError.Line
	
End If


Sub TrapError(strError)

	bolErrors = True
	strErrorMessage = strErrorMessage & strError & ", "
	ProcessErrors Request.ServerVariables("SCRIPT_NAME"),strErrorMessage
	
End Sub


Sub ProcessErrors(strPageName,strErrorMessage)

	If bolErrors Then
		'Log Error
		objCon.Execute "INSERT INTO tblErrorLog VALUES ('" & strPageName & "','" & strErrorMessage & "',GetDate()," & Session("UserID") & ")"
	End If
	
	Response.write "<div class=""content-body""><section id=""basic-alerts""><div class=""alert alert-danger alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
					"<span aria-hidden=""true"">&times;</span></button>" & _
					"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
					"<span>ERROR """ & strErrorMessage & """</span></div></div></div>"
	
End Sub


 %>

<footer class="footer myfi-footer">
      <div class="container">
        <div class="row">
          <div class="col-md-3">
            <img src="<%=Session("ServerPath")%>images/defence_logo_dark.png" class="defence-logo" alt="Department of Defence"/>
            <a class="btn btn-primary btn-myfi" href="#"><i class="fa fa-home"></i> Back to MyFi Portal</a>
          </div>
          <div class="col-md-3">
            <h4 class="footer-title">Shortcuts</h4>
            <ul class="footer-nav">
              <li class="nav-item">
                <a class="nav-link" href="<%=Session("ServerPath")%>CC/HomeAdmin.asp" >Dashboard</a>
              </li>
              <li class="nav-item">
                <a class="nav-link" href="<%=Session("ServerPath")%>CC/Applications.asp" >Applications</a>
              </li>
              <li class="nav-item">
                <a class="nav-link" href="<%=Session("ServerPath")%>CC/Cards.asp?View=All">Cards</a>
              </li>
              <li class="nav-item">
                <a class="nav-link" href="#">Reporting</a>
              </li>
              <li class="nav-item">
                <a class="nav-link" href="<%=Session("ServerPath")%>CC/HomeAdmin.asp">My details</a>
              </li>
            </ul>
          </div>
          <div class="col-md-3">
            <h4 class="footer-title">Getting started</h4>
            <ul class="footer-nav">
              <li class="nav-item">
                <a class="nav-link" href="<%=Session("ServerPath")%>CC/img/ApplyForACard.pdf" Target="_new">Frequently Asked Questions</a>
              </li>
              <li class="nav-item">
                <a class="nav-link" href="<%=Session("ServerPath")%>CC/HelpFile.html" Target="_new">How to Guide</a>
              </li>
            </ul>
          </div>
          <div class="col-md-3">
            <h4 class="footer-title">Help</h4>
            <ul class="footer-nav">
              <li class="nav-item">
                <a class="nav-link" href="<%=Session("ServerPath")%>CC/img/ApplyForACard.pdf" Target="_new">Frequently Asked Questions</a>
              </li>
              <li class="nav-item">
                <a class="nav-link" href="<%=Session("ServerPath")%>CC/HelpFile.html" Target="_new">How to Guide</a>
              </li>
              <li class="nav-item">
                <a class="nav-link" href="#">Contact</a>
              </li>
            </ul>
          </div>
        </div>
        <div class="row">
          <div class="col-12 text-center">
            <span class="copyright">© Department of Defence - Defence Finance Group 2020</span>
          </div>
        </div>
      </div>
    </footer>

		
	<script src="<%=Session("ServerPath")%>js/custom.js"></script>
  </body>
</html>
