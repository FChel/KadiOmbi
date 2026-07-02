
<!-- #Include file=CAPSHeader.asp -->
<!-- #Include file=ADOVBS.inc -->
<!-- #Include file=CAPSFunctions.asp -->
<%

Response.ContentType = "text/html" 
Response.CodePage = 65001
Response.CharSet = "UTF-8"

If IsEmpty(Session("UserID")) Then Response.Redirect("../Timeout.asp?State=Expired")


'Description:	Create and view applications
'Author:		MG
'Date:			January 2020

	Response.Expires = -1500	

Dim objCon
Dim objCmd
Dim objRS
Dim objRS1
Dim strSelected
Dim x 
Dim strMessage
Dim strColour
Dim dteBatchDate
Dim intTotalTasks
Dim intTasksComplete
Dim intMorningTasks
Dim intAfternoonTasks
Dim intMorningTasksComplete
Dim intAfternoonTasksComplete
Dim strDateButton
Dim arrMonths(12)
Dim strDateButtonTop
Dim strDateButtonDrop
Dim arrYears(10)
Dim strYearButtonTop
Dim strYearButtonDrop
Dim lngApps
Dim dblProcessDays
Dim intProcessDaysDelim
Dim strYearButton
Dim strYear
Dim strMonth
Dim strWhere
Dim strSQL
Dim lngCSRecords
Dim lngCSCancels
Dim lngNewCards
Dim lngANZApps
Dim lngPDFApps
Dim lngSCApps
Dim lngPortalApps
Dim arrAppTypePercent(3)
Dim strBank
Dim strBankButton

'Set Database objects

Set objCon = Server.CreateObject("ADODB.Connection")
Set objCmd = Server.CreateObject("ADODB.Command")
Set objRS = Server.CreateObject("ADODB.Recordset")

'Open database connection
objCon.Open Session("DBConnection")

If IsNull(Session("BatchDate")) Or Session("BatchDate") = "" Then Session("BatchDate") = now()

If Not IsEmpty(Request.QueryString("FileDate")) Then

	Session("BatchDate") = Request.QueryString("FileDate")
End If

If Not IsEmpty(Request.QueryString("ReportMonth")) Then
	Session("ReportMonth") = Request.QueryString("ReportMonth")
End If

If Not IsEmpty(Request.QueryString("ReportYear")) Then
	Session("ReportYear") = Request.QueryString("ReportYear")
End If

If IsNull(Session("ReportYear")) OR Session("ReportYear") = "" Then
	Session("ReportYear") = Year(now())
End If

If IsNull(Session("ReportMonth")) OR Session("ReportMonth") = "" Then
	'Session("ReportMonth") = "Nov"
	Session("ReportMonth") = Left(MonthName(Month(now())),3)
End If

strBank = "NAB"

If Not IsEmpty(Request.QueryString("Bank")) Then
	strBank = Request.QueryString("Bank")
End If


arrMonths(1) = "Jan"
arrMonths(2) = "Feb"
arrMonths(3) = "Mar"
arrMonths(4) = "Apr"
arrMonths(5) = "May"
arrMonths(6) = "Jun"
arrMonths(7) = "Jul"
arrMonths(8) = "Aug"
arrMonths(9) = "Sep"
arrMonths(10) = "Oct"
arrMonths(11) = "Nov"
arrMonths(12) = "Dec"

arrYears(1) = "2020"
arrYears(2) = "2021"
arrYears(3) = "2022"
arrYears(4) = "2023"
arrYears(5) = "2024"
arrYears(6) = "2025"
arrYears(7) = "2026"

'Call the procedure to load the count of tasks and tasks completed for display
'Call LoadTasksCount()

'strDateButton = "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" data-toggle=""modal"" data-target=""#ModApp"" onclick=""self.location='ExportNAFile.asp?Action=Process&FileLoadID=" & strColour & "'""; title=""Click to Process the CS File and update changes to cards in CAPS from the CS File loaded " & strColour & """>Process</button>"

For x = 1 to 12
	If cstr(Session("ReportMonth")) = cstr(arrMonths(x)) Then
		strDateButtonTop = "<div class=""dropdown""><button class=""btn btn-outline-secondary dropdown-toggle"" type=""button"" id=""dropdownMenuButton"" data-toggle=""dropdown"" aria-haspopup=""true"" aria-expanded=""false"">" & arrMonths(x) & "</button>" & _
			"<div class=""dropdown-menu"" aria-labelledby=""dropdownMenuButton"">"
	Else
		strDateButtonDrop = strDateButtonDrop & "<a class=""dropdown-item"" href=""HomeCCAdmin.asp?Action=Link=AD&ReportMonth=" & arrMonths(x) & """;>" & arrMonths(x) & "</a>"
	End If
	
Next

strDateButton = "<div class=""dropdown""><button class=""btn btn-outline-secondary dropdown-toggle"" type=""button"" id=""dropdownMenuButton"" data-toggle=""dropdown"" aria-haspopup=""true"" aria-expanded=""false"">Month</button>" & _
                "<div class=""dropdown-menu"" aria-labelledby=""dropdownMenuButton"">" & _
                "<a class=""dropdown-item"" href=""#"">Action</a>" & _
                "<a class=""dropdown-item"" href=""#"">Another action</a>" & _
                "<a class=""dropdown-item"" href=""#"">Something else here</a>" & _
                "</div></div>"
	
	strDateButton = strDateButtonTop & strDateButtonDrop & "</div></div>"

	'Get the Report Year or use the current year if one has not been selected
	If IsNull(Session("ReportYear")) OR Session("ReportYear") = "" Then
		strYear = Year(Now())
	Else
		strYear = Session("ReportYear")
	End If
	
For x = 1 to 7
	If cstr(Session("ReportYear")) = cstr(arrYears(x)) Then
		strYearButtonTop = "<div class=""dropdown""><button class=""btn btn-outline-secondary dropdown-toggle"" type=""button"" id=""dropdownMenuButton"" data-toggle=""dropdown"" aria-haspopup=""true"" aria-expanded=""false"">" & arrYears(x) & "</button>" & _
			"<div class=""dropdown-menu"" aria-labelledby=""dropdownMenuButton"">"
	Else
		strYearButtonDrop = strYearButtonDrop & "<a class=""dropdown-item"" href=""HomeCCAdmin.asp?Action=Link=AD&ReportYear=" & arrYears(x) & """;>" & arrYears(x) & "</a>"
	End If
	
Next

	strYearButton = strYearButtonTop & strYearButtonDrop & "</div></div>"

	If IsNull(Session("ReportMonth")) OR Session("ReportMonth") = "" Then
		strMonth = "11"
	Else
		Select Case(Session("ReportMonth"))
			Case "Jan"
				strMonth = "01"
			Case "Feb"
				strMonth = "02"
			Case "Mar"
				strMonth = "03"
			Case"Apr"
				strMonth = "04"
			Case "May"
				strMonth = "05"
			Case "Jun"
				strMonth = "06"
			Case "Jul"
				strMonth = "07"
			Case "Aug"
				strMonth = "08"
			Case "Sep"
				strMonth = "09"
			Case "Oct"
				strMonth = "10"
			Case "Nov"
				strMonth = "11"
			Case "Dec"
				strMonth = "12"
			Case Else
				strMonth = "01"
		End Select
		'strMonth = PadDigits(Session("ReportMonth"),2)
	End If

	strWhere = strMonth & strYear

	strSQL = "SELECT COUNT([EmployeeID]) As Apps,CardType,[CardTypeSub],CAST(CAST(RIGHT('00'+CAST(Month(DateSubmitted) as varchar(2)),2) as varchar(2)) + '' + CAST(YEAR(DateSubmitted) as varchar(20)) AS Char(6)) As MonthYear,DATEDIFF(d,[DateSubmitted],[DateReviewed]) as ProcessDays,[DateSubmitted],[DateReviewed],ApplicationType " & _
	"FROM tblCAPSApplication " & _
	"WHERE CAST(CAST(RIGHT('00'+CAST(Month(DateSubmitted) as varchar(2)),2) as varchar(2)) + '' + CAST(YEAR(DateSubmitted) as varchar(20)) AS Char(6)) = '" & strWHERE & "' " & _
	"GROUP BY [CardType],[CardTypeSub],CAST(CAST(RIGHT('00'+CAST(Month(DateSubmitted) as varchar(2)),2) as varchar(2)) + '' + CAST(YEAR(DateSubmitted) as varchar(20)) AS Char(6)),DATEDIFF(d,[DateSubmitted],[DateReviewed]),[DateSubmitted],[DateReviewed],ApplicationType"

	lngSCApps=0
	lngPDFApps=0
	lngPortalApps = 0

   'Description:	Loads Position details into page if applicable.
	objRS.Open strSQL,objCon

	If Not objRS.EOF Then
			
			Do Until objRS.EOF

				lngApps = lngApps + objRS("Apps")

				''''New for the Service Connect Application Types
				If IsNull(objRS("ApplicationType")) Then

				Else
					If objRS("ApplicationType") = "AE602 XML" Then
						lngPDFApps = lngPDFApps + objRS("Apps")
					ElseIf objRS("ApplicationType") = "Portal" Then
						lngPortalApps = lngPortalApps + objRS("Apps")
					Else
						lngSCApps = lngSCApps + objRS("Apps")
					End If
				End If
				
				If IsNull(objRS("ProcessDays")) Then
				Else
					'response.write objRS("ProcessDays") & ", "
					dblProcessDays = dblProcessDays + objRS("ProcessDays")
					intProcessDaysDelim = intProcessDaysDelim + 1
				End If
				
				objRS.Movenext
			Loop
		Else
			lngApps = 0
	   End If
	   
	objRS.Close
	
	If strBank = "NAB" Then
		strBankButton = "<button class=""btn btn-outline-secondary "" onClick=""self.location='HomeCCAdmin.asp?Bank=Diners'"" type=""button"" id=""Bank"">NAB</button>"
	Else
		strBankButton = "<button class=""btn btn-outline-secondary "" onClick=""self.location='HomeCCAdmin.asp?Bank=NAB'"" type=""button"" id=""Bank"">Diners</button>"
	End If
	
	'Get the AppTypePercentages
	If lngApps=0 Then
		arrAppTypePercent(1) = 0
		arrAppTypePercent(2) = 0
		arrAppTypePercent(3) = 0
	Else
		arrAppTypePercent(1) = (lngPDFApps/lngApps)*100
		arrAppTypePercent(2) = (lngSCApps/lngApps)*100
		arrAppTypePercent(3) = (lngPortalApps/lngApps)*100
	End If
	
	arrAppTypePercent(1) = Round(arrAppTypePercent(1),2)
	arrAppTypePercent(2) = Round(arrAppTypePercent(2),2)
	arrAppTypePercent(3) = Round(arrAppTypePercent(3),2)
	
	If IsNull(dblProcessDays) OR dblProcessDays = "" Then
		dblProcessDays = 0
	Else
		dblProcessDays = dblProcessDays/intProcessDaysDelim
		
		dblProcessDays = FormatNumber(dblProcessDays,2)
		'response.write dblProcessDays
		
		'response.write  dblProcessDays & " " & intProcessDaysDelim
	End If
	
	'Call the procedure to get the Summary Stats
	Call GetSummaryStats()
	
%>
<script>
function OpenSs(cb) {

	//var id = cb.getAttribute('CardTypeSelect');
	//var id = document.getElementByName("CardTypeSelect").value;
	//alert(id);
	//var e = document.getElementById("CardTypeSelect");
	//var result = e.options[e.selectedIndex].value;
	
	//document.getElementById('CardType').value=result;
	alert('asa');
}

//function DatePickChange() {
//	self.location="HomeAdmin.asp?FileDate=" + document.getElementById("AdminDate").value;
//}

jQuery(document).ready(function($) {
    $(".clickable-row").click(function() {
        window.location = $(this).data("href");
    });
});

</script>

<!-- Modal -->
<div class="modal fade" id="ModApp" tabindex="-1" role="dialog" aria-labelledby="exampleModalCenterTitle" aria-hidden="true">
  <div class="modal-dialog modal-dialog-centered" role="document">
    <div class="modal-content">
      <div class="modal-header">
        <h5 class="modal-title" id="exampleModalLongTitle">CAPS Contact</h5>
        <button type="button" class="close" data-dismiss="modal" aria-label="Close">
          <span aria-hidden="true">&times;</span>
        </button>
      </div>
      <div class="modal-body">
        <label>Contact Card Providers</label><br>

			<h6 class="modal-title" id="exampleModalLongTitle">Diners - 1800 123 123</h6>
			<h6 class="modal-title" id="exampleModalLongTitle">ANZ - 132313</h6>
      </div>
      <div class="modal-footer">
        <button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
      </div>
    </div>
  </div>
</div>


    <main class="main py-3">
      <div class="container">
		<h2 class="text-left py-3">Credit Card Team Dashboard - <span style="font-size:18px;"><%=Session("UserName")%></span></h2>
	  
        <div class="row">
          <div class="col-md-6">
           		
			<div class="panel panel-shadow mb-3">
              <form class="inline">
			  <div class="panel-header">
			  <div class="panel-content row" style="padding: 0px 0px 0px 0px;">
                <div class="col-md-6"><h4>Card Application Summary</h4></div> <div class="col-md-2"><%=strBankButton%></div><div class="col-md-2 text-right my-auto"><%=strDateButton%></div><div class="col-md-2 text-right my-auto"><%=strYearButton%></div>
              </div>
			  </form>
			   </div>
			  <div class="panel-content row">
				<%
					If strBank = "Diners" Then
						Response.Write "<iframe width=""400px;"" height=""250px;"" src=""../vendor/chart.js/CardAppsBar.asp"" scrolling=""no"" style=""border:none;""></iframe>" & _
								"<iframe width=""200px;"" height=""250px;"" src=""../vendor/chart.js/CardAppsPie.asp"" scrolling=""no"" style=""border:none;""></iframe>"
					Else
						Response.Write "<iframe width=""400px;"" height=""250px;"" src=""../vendor/chart.js/CardAppsBarNAB.asp"" scrolling=""no"" style=""border:none;""></iframe>" & _
								"<iframe width=""200px;"" height=""250px;"" src=""../vendor/chart.js/CardAppsPieNAB.asp"" scrolling=""no"" style=""border:none;""></iframe>"
					End If				
				
				%>
				<!--<iframe width="400px;" height="250px;" src="../vendor/chart.js/CardAppsBarNAB.asp" scrolling="no" style="border:none;"></iframe>
				<iframe width="200px;" height="250px;" src="../vendor/chart.js/CardAppsPieNAB.asp" scrolling="no" style="border:none;"></iframe>-->
				
			 </div>
				
              </div> 
			  
			  <%
					If strBank = "Diners" Then
						Response.Write "<div class=""panel panel-shadow mb-3"">" & _
							"<div class=""panel-header"">" & _
							"<h4><i class=""fa fa-credit-card""></i> Unactivated Cards</h4>" & _
							"</div>"
							
						Call DisplayUnactivated()
							
						Response.Write "</div>"
						
						Response.Write  "<div class=""panel panel-shadow mb-3""><div class=""panel-content row"">" & _
							"<iframe width=""600px;"" height=""350px;"" src=""../vendor/chart.js/CardAppsLine.asp"" scrolling=""no"" style=""border:none;""></iframe>" & _
							"</div></div>"
			 					
					Else
						
						Response.Write  "<div class=""panel panel-shadow mb-3""><div class=""panel-content row"">" & _
							"<iframe width=""600px;"" height=""350px;"" src=""../vendor/chart.js/CardAppsLineNAB.asp"" scrolling=""no"" style=""border:none;""></iframe>" & _
							"</div></div>"
					End If				
				
			%>
			 
			  <div class="panel panel-shadow mb-3">
				  <div class="panel-content row">
					<iframe width="600px;" height="180px;" src="../vendor/chart.js/AppTypesPie.asp" scrolling="no" style="border:none;"></iframe>
				 </div>
			 </div>
			 <div class="panel panel-shadow mb-3">
				  <div class="panel-content row">
					<iframe width="600px;" height="380px;" src="../vendor/chart.js/AppTypesLine.asp" scrolling="no" style="border:none;"></iframe>
				 </div>
			 </div>
			 
          </div>
		  
		  
          <div class="col-md-6 sidebar">
		  

            <div class="panel panel-shadow mb-3">
              <div class="panel-header">
                <h4>Summary</h4>
                <span class="panel-subheader">Admin Summary for</span> <span style="font-weight:bold; font-size=18px;"><%=Session("ReportMonth") & " " & Session("ReportYear")%></span>
              </div>
			   <div class="panel-content row" style="padding: 5px;">
              <div class="col-md-12 text-center my-auto" style="font-weight:bold; height:20px;">Applications
			  </div>
			  </div>
			  
			  <div class="panel-content row">
              <div class="col-md-2 text-left my-auto">
                  <i style="font-size:12px" title="Old style AE602 PDF/XML form completed and emailed to Credit Cards" class="fa fa-file-pdf"></i>
                  <span class="content" style="font-size:14px; text-align:right;">PDF AE602 + <p style="font-size:18px; font-weight:bold; color:black; text-align:right;"><%=lngPDFApps & "</br><small>" & arrAppTypePercent(1) & "%</small>"%></p></span>
				</div>
				
				<div class="col-md-2 text-left my-auto">
                  <i style="font-size:12px" title="Service Connect Applications" class="fa fa-link"></i>
                  <span class="content" style="font-size:14px; text-align:right;">Srv. Connect =<p style="font-size:18px; font-weight:bold; color:black; text-align:right;"><%=lngSCApps & "</br><small>" & arrAppTypePercent(2) & "%</small>"%></p></span>
				</div>
				<div class="col-md-2 text-left my-auto">
                  <i style="font-size:12px" title="DCCP Portal Applications" class="fa fa-globe"></i>
                  <span class="content" style="font-size:14px; text-align:right;">DCCP Portal =<p style="font-size:18px; font-weight:bold; color:black; text-align:right;"><%=lngPortalApps & "</br><small>" & arrAppTypePercent(3) & "%</small>"%></p></span>
				</div>
				<div class="col-md-3 text-left my-auto" style="border-right: 1px gray solid;">
                <a href="Applications.asp?UserView=All&Link=AP" class="block-link" >
                  <i style="font-size:14px" class="fa fa-credit-card"></i>
                  <span class="content" style="font-size:14px; text-align:right;">Total Applications <p style="font-size:18px; font-weight:bold; color:black; text-align:right;"><%=lngApps%></p></span>
				 
                </a>
				</div>
				
				<div class="col-md-3 text-right my-auto">
                <a href="#" class="block-link" data-toggle="modal" data-target="#ModApp">
                  <i style="font-size:14px" class="fa fa-clock"></i>
                  <span class="content" style="font-size:12px">Average Process Time <p style="font-size:18px; font-weight:bold; color:black;"><%=dblProcessDays%> Days</p></span>
                </a></div></div>
				
				<!--<iframe width="600px;" height="200px;" src="../vendor/chart.js/BarChart2.html" scrollbar="no" border="0px;"></iframe>-->
				
              </div> 
			  
			  
			  <div class="panel panel-shadow mb-3">
				<div class="panel-header" style="text-align:center; padding: 0px;">
                <span class="panel-subheader">File Export Summary for</span> <span style="font-weight:bold; font-size=18px;"><%=FormatDateTime(now(),vbShortDate)%></span>
              </div>
			  <div class="panel-content row" style="padding: 0px;">
			  
			  <div class="col-md-3 text-right my-auto" style="padding-top: 25px;">
                <a href="../Admin/CAPSAdmin/ExportNANAB.asp?CardType=DPC&Link=AD" style="font-size:18px; font-weight:bold;">
                  <i class="fa fa-dollar-sign fa-sm" style="padding: 0px; margin:2px;"></i>
                  <span class="content" title="DPC Cards waiting to be exported to Diners in the CS File today" style="padding: 0px; margin:2px;">DPC Apps<p style="font-size:22px; font-weight:bold; color:black;"><%=lngANZApps%></p></span>
                </a></div>
				
              <div class="col-md-3 text-right my-auto" style="padding: 5px;">
                <a href="../Admin/CAPSAdmin/ExportCSNAB.asp?CardType=DTC&Link=AD" class="block-link" >
                  <i class="fa fa-plane fa-xs" style="padding: 0px; margin:2px;"></i>
                  <span class="content" title="Cards waiting to be exported to Diners in the CS File today" style="padding: 0px; margin:2px;">CM To NAB<p style="font-size:22px; font-weight:bold; color:black; "><%=lngCSRecords%></p></span>
				 
                </a>
				</div>
				<div class="col-md-3 text-right my-auto" style="padding: 5px;">
                <a href="../Admin/CSTransactionsToNAB.asp?FileLoadID=0&ViewButton=Cancelled" class="block-link">
                  <i class="fa fa-times fa-xs" style="padding: 0px; margin:2px;"></i>
                  <span class="content" title="Cards Cancellations waiting to be exported to Diners in the CS File today">CM Cancels<p style="font-size:22px; font-weight:bold; color:black;"><%=lngCSCancels%></p></span>
                </a></div>
				<div class="col-md-3 text-right my-auto" style="padding: 5px;">
                <a href="../Admin/CAPSAdmin/ExportNANAB.asp?CardType=DPC&Link=AD" class="block-link">
                  <i class="fa fa-credit-card fa-xs" style="padding: 0px; margin:2px;"></i>
                  <span class="content" title="New Cards waiting to be exported to Diners in the CS File today">NA File<p style="font-size:22px; font-weight:bold; color:black;"><%=lngNewCards%></p></span>
                </a></div>
				
				
				
				</div>
				
				<!--<iframe width="600px;" height="200px;" src="../vendor/chart.js/BarChart2.html" scrollbar="no" border="0px;"></iframe>-->
				
              </div> 
			  
			  
			  
			 
			   <div class="panel panel-shadow mb-3">

					<div class="panel-header">
						<h4><i class="fa fa-credit-card"></i> Cards Due To Expire</h4>
						
					  </div>
			  
					<%Call DisplayExpiring()%>
			
              </div> 
			  
			  <div class="panel panel-shadow mb-3">

					<div class="panel-header">
						<h4><i class="fa fa-file"></i> NA File Today <%=FormatDatetime(Now(),vbShortDate)%></h4>
						
					  </div>
			  
					<%Call DisplayNAFile()%>
			
              </div> 
			  
			 
			  
				

     
		
			
          </div>
        </div>
		
		
		
    </main>

	<!--<script src="js/jquery.js"></script>
    <script src="bootstrap/js/bootstrap.bundle.min.js"></script>-->
   
	 <!-- Chart.js -->
    <script src="../js/plugins/Chart.min.js"></script>
	
<!-- #Include file=CAPSFooter.asp -->
  </body>
</html>

<%



Public Sub DisplayExpiring()
Dim y
Dim strAction
Dim strStatus
Dim dteDateSubmitted
Dim dteDateReviewed
Dim strSearch
Dim strRecordMessage
Dim strCardNo
Dim strDaysColour
Dim strProcessStatus
Dim dteWarningDate
Dim strPages
Dim strSort
Dim strOrderType
Dim strHeader
Dim strCardType
Dim lngTotalRecords

	strSearch = Request.Form("SearchInput")
	
	If IsEmpty(Request.QueryString("SortType")) Then
		'strOrderType = "ASC"
	Else
		If Request.QueryString("SortType") = "ASC" Then
			strOrderType = "DESC"
		Else
			strOrderType = "ASC"
		End If
	End If
	
	'If IsEmpty(Request.QueryString("Sort")) Then
	'	strSort = ""
	'Else		
		strSort = " ORDER BY [DaysToExpiry] ASC"
	'End If

	
	If Session("ViewButton") = "Emailed" Then
		strWhere = " AND [ProcessStatus] = 'Email Expiring' "
	ElseIf Session("ViewButton") = "Removed" Then
		strWhere = " AND [ProcessStatus] = 'Removed Expiring' "
	ElseIf Session("ViewButton") = "AddedToCS" Then
		strWhere = " AND [ProcessStatus] = 'Added To CS' "
	Else
		'This catches ALL
		strWhere = ""
	End If
	
	'Determine the Daye to Expire being displayed
	strWhere = strWhere & " AND ([DaysToExpiry] > 0 AND [DaysToExpiry] < 120)"
	
If strSearch = "" OR ISNull(strSearch) Then
	'If Session("UserView") = "All" Then
		strSQL = "SELECT TOP 5 * FROM qryCAPSCardsExpiry WITH(NOLOCK) WHERE [Status] = ''" & strWhere & strSort
	'Else
	'	strSQL = "SELECT top 100 * FROM qryCAPSCardsExpiry WITH(NOLOCK) WHERE [ActivationFlag] = 'N' AND EmployeeID = '" & Session("EmployeeID") & "'"
	'End If
	
Else
	'If Session("UserView") = "All" Then
		strSQL = "SELECT TOP 5 * FROM qryCAPSCardsExpiry WITH(NOLOCK) WHERE [Status] = '' AND (EmployeeID Like '%" & strSearch & "%' OR FirstName Like '%" & strSearch & "%' OR Surname Like '%" & strSearch & "%')" & strWhere & strSort
	'Else
	'	strSQL = "SELECT top 100 * FROM qryCAPSCardsExpiry WITH(NOLOCK) WHERE [ActivationFlag] = 'N' AND EmployeeID = '" & Session("EmployeeID") & "' AND (EmployeeID Like '%" & strSearch & "%' OR FirstName Like '%" & strSearch & "%' OR Surname Like '%" & strSearch & "%')"
	'End If
End If

'Build the message displayed at the bottom of the screen with the search details
If Session("UserView") = "All" Then
	strRecordMessage = strSearch
Else
	'strRecordMessage = "for " & Session("UserName") 
End If

If strSearch = "" OR ISNull(strSearch) Then
	'strRecordMessage = "<TR><TH colspan=""10"" Style=""text-align:center;"">No Expiring Cards within the next 120 days</TH>" & _
				'"<TH colspan=""3"" style=""text-align:center;""></TH></TR>"
Else
	strRecordMessage = "<TR><TH colspan=""10"" Style=""text-align:center;"">No Expiring Cards for " & strSearch & "</TH>" & _
				"<TH colspan=""3"" style=""text-align:center;""></TH></TR>"
End If

objRS.Open strSQL,objCon,3,1

    y = 0
	
	
	
	'Write a message in the list if there are no Expiring Cards
	If objRS.EOF Then
		Response.Write "<TR><TH colspan=""10"" Style=""text-align:center;"">No Cards Expiring within the next 120 days</TH>" & _
				"<TH colspan=""3"" style=""text-align:center;""></TH></TR>"
	Else
		objRS.Movelast
		objRS.Movefirst
		lngTotalRecords = objRS.Recordcount
		
		
		'Response.Write "<div class=""row""><div class=""col-12"">" & _
		Response.Write	"<table class=""table table-compact text-left""><thead><tr>" & _
			"<th scope=""col""> Card ID </th>" & _
			"<th scope=""col""> EID </th>" & _
			"<th scope=""col""> Name </th>" & _
			"<th scope=""col""> Card  </th>" & _
			"<th scope=""col""> Expiry Date </th>" & _
			"<th scope=""col"" title=""Days until Card Expires (up to today)""> Days </th>" & _
			"</tr></thead><tbody class=""text-left"">"
					
	End If
	
	
	'Write a message in the list if there are no applications
	If objRS.EOF Then
		Response.Write strRecordMessage
	End If
    	
    Do until objRS.EOF 

		y = y + 1
		
			x = x + 1
			
			'Create the actions based on the Process Status of the card
			Select Case objRS("ProcessStatus")
			
			Case  "Removed Expiring"
				strAction = "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" onclick=""self.location='ExpiringCards.asp?Action=UnRemove&CardID=" & objrs("CardID") & "&CardEID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-minus-list""></i> Re-List</button>"
			Case "Added to CS"

				strAction = "<button type=""button"" class=""btn btn-secondary btn-xs"" onclick=""self.location='../Admin/CAPSAdmin/ExportCS.asp?CardID=" & objrs("CardID") & "&CardEID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-key""></i> View CS</button>"
				
			Case "Email Expiring"
				strAction = "<button type=""button"" class=""btn btn-danger btn-xs"" onclick=""self.location='ExpiringCards.asp?Action=Cancel&CardID=" & objrs("CardID") & "&CardEID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-minus-circle""></i> Cancel</button>"
			
			Case Else
				strAction = "<button type=""button"" class=""btn btn-danger btn-xs"" onclick=""self.location='ExpiringCards.asp?Action=Cancel&CardID=" & objrs("CardID") & "&CardEID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-minus-circle""></i> Cancel</button>"
				strAction = strAction & "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" data-toggle=""modal"" data-target=""#EmailModal""><i class=""fa fa-minus-mail""></i> Email</button>"
				strAction = strAction & "<button type=""button"" class=""btn btn-outline-info btn-xs"" onclick=""self.location='ExpiringCards.asp?Action=Remove&CardID=" & objrs("CardID") & "&CardEID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-minus-cross""></i> Remove</button>"

				'strAction = strAction & "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" onclick=""self.location='ExpiringCards.asp?Action=Email&CardID=" & objrs("CardID") & "&CardEID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-minus-mail""></i> Email</button>"
				'data-toggle="modal" data-target="#EmailModal"
				
			End Select
			
			'Create the Status list badge based on the status field
			If IsNull(objRS("Status")) Then
				strStatus = ""
			Else
				If objRS("Status") = "00" Then
					strStatus = "<span class=""badge badge-pill badge-success"">Active</span>"
				ElseIf objRS("Status") = "01" OR objRS("Status") = "02" Then
					strStatus = "<span class=""badge badge-pill badge-danger"">Cancelled</span>"
				Else
					strStatus = ""
				End If
			End If
			
			If IsNull(objRS("Expiry")) Then
				dteDateSubmitted = ""
			Else
				dteDateSubmitted = FormatDateTime(objRS("Expiry"),vbShortDate)
			End If
			
			If IsNull(objRS("Expiry")) Then
				dteDateReviewed = ""
			Else
				dteDateReviewed = DateDiff("d",now(),objRS("Expiry"))
				If dteDateReviewed > 80 Then
					strDaysColour = "Style=""color:red; font-weight:bold;"""
				ElseIf dteDateReviewed > 45 Then
					strDaysColour = "Style=""color:orange; font-weight:bold;"""
				Else
					strDaysColour = "Style=""color:black"""
				End If
			End If
			
			'Format the Card number so it is masked depending on the card type
			If IsNull(objRS("CardNumber")) Then
				strCardNo = ""
			Else
				strCardNo = objRS("CardNumber")
				If mid(strCardNo,5,1)=0 Then 
					strCardNo = mid(strCardNo,6,2) & "****" & right(strCardNo,4)
				Else
					strCardNo = mid(strCardNo,4,2) & "****" & right(strCardNo,4)
				End If
				'If len(strCardNo)>8 Then strCardNo = mid(strCardNo,6,2) & "****" & right(strCardNo,4)
			End If
			
			If IsNull(objRS("ProcessStatus")) or objRS("ProcessStatus") = "" Then
				strProcessStatus = ""
			Else
				If objRS("ProcessStatus") = "Email Expiring" Then
					strProcessStatus = "Emailed"
				Elseif objRS("ProcessStatus") = "Removed Expiring" Then
					strProcessStatus = "Removed"
				Elseif objRS("ProcessStatus") = "Added to CS" Then
					strProcessStatus = "Added To CS"
				End If
			End If
			
			dteWarningDate = "title=" & objRS("Warning") & ""
			
			If isNull(objRS("CardTypeSub")) Then
				strHeader = ""
				strCardType = ""
			Else
				strHeader = objRS("CardTypeSub")
				strCardType  = objRS("CardType") & " " & objRS("CardTypeSub")
			End If
					
			'Determine the image and title based on the card type
			If Trim(strHeader) = "Diners" Then
				strHeader = "<img height=""20px"" width=""20px"" src=""../images/icon_diners.png"" title=""" & strCardType & """> " & strCardType
			ElseIf strHeader = "ANZ" Then
				strHeader = "<img height=""20px"" width=""20px"" src=""../images/logo_ANZ.png"" Title=""" & strCardType & """> " & strCardType
			ElseIf strHeader = "Mastercard" Then
				strHeader = "<img height=""20px"" width=""20px"" src=""../images/logo_mc.png"" Title=""" & strCardType & """> " & strCardType
			Else
				strHeader = "<img height=""20px"" width=""20px"" src=""../images/logo_coa.png"" Title=""" & strCardType & """> " & strCardType
			End If
		
		
			Response.Write "<TR><TD ><a Target=""_self"" HREF=""CardDetail.asp?CardID=" & objRS(0) & """>" & objRS(0) & "</a></TD><TD>" & objRS("EmployeeID") & "</a></TD>" & _
					"<TD style=""font-size:12px;""><a Target=""_self"" HREF=""CardDetail.asp?CardID=" & objRS(0) & """>" & objRS("FirstName") & " " & objRS("Surname") & "</a></TD><TD style=""font-size:12px;""><a Target=""_self"" HREF=""CardDetail.asp?CardID=" & objRS(0) & """>" & strHeader & "</a></TD>" & _
					"<TD >" & dteDateSubmitted & "</TD><TD " & strDaysColour & ">" & dteDateReviewed & "</TD><TD style=""font-size:12px;"" " & dteWarningDate & "></TR>"
					
			'response.write "<TR><TD ><a Target=""_self"" HREF=""CardDetail.asp?CardID=" & objRS(0) & """>" & objRS(0) & "</a></TD><TD>" & objRS("EmployeeID") & "</a></TD>" & _
			'		"<TD style=""font-size:12px;""><a Target=""_self"" HREF=""CardDetail.asp?CardID=" & objRS(0) & """>" & objRS("FirstName") & " " & objRS("Surname") & "</a></TD><TD style=""font-size:12px;""><a Target=""_self"" HREF=""CardDetail.asp?CardID=" & objRS(0) & """>" & objRS("CardType") & " " & objRS("CardTypeSub") & "</a></TD>" & _
			'		"<TD >" & dteDateSubmitted & "</TD><TD " & strDaysColour & ">" & dteDateReviewed & "</TD><TD style=""font-size:12px;"" " & dteWarningDate & "></TR>"
					
		
		objRS.movenext
	Loop
	
	Response.Write "<TR><TH colspan=""4""><a href=""ExpiringCards.asp"">More...</a></TH>" & _
				"<TH colspan=""2"" style=""text-align:center;"">" & x & "</TH></TR></TABLE>"
				
	'	Response.Write "<TR><TH colspan=""4"">Total</TH>" & _
	'			"<TH colspan=""2"" style=""text-align:center;"">" & x & "</TH></TR></TABLE>"

				
objRS.Close

End Sub


Public Sub DisplayUnactivated()
Dim y
Dim strAction
Dim strStatus
Dim dteDateSubmitted
Dim dteDateReviewed
Dim strSearch
Dim strRecordMessage
Dim strCardNo
Dim strDaysColour
Dim strProcessStatus
Dim dteWarningDate
Dim strPages
Dim strSort
Dim strOrderType
Dim strHeader
Dim strCardType
Dim lngTotalRecords

	strSearch = Request.Form("SearchInput")
	
	If IsEmpty(Request.QueryString("SortType")) Then
		'strOrderType = "ASC"
	Else
		If Request.QueryString("SortType") = "ASC" Then
			strOrderType = "DESC"
		Else
			strOrderType = "ASC"
		End If
	End If
	
	If IsEmpty(Request.QueryString("Sort")) Then
		strSort = ""
	Else		
		strSort = " ORDER BY " & Request.QueryString("Sort") & " " & strOrderType
	End If
	
	If Session("ViewButton") = "Emailed" Then
		strWhere = " AND [ProcessStatus] = 'Email Unactivated' "
	ElseIf Session("ViewButton") = "Removed" Then
		strWhere = " AND [ProcessStatus] = 'Removed Unactivated' "
	ElseIf Session("ViewButton") = "AddedToCS" Then
		strWhere = " AND [ProcessStatus] = 'Added To CS' "
	Else
		'This catches ALL
		strWhere = ""
	End If
	
If strSearch = "" OR ISNull(strSearch) Then
	'If Session("UserView") = "All" Then
		strSQL = "SELECT TOP 5 * FROM qryCAPSCards WITH(NOLOCK) WHERE [ActivationFlag] = 'N' AND Status = '00'" & strWhere & strSort
	'Else
	'	strSQL = "SELECT top 100 * FROM qryCAPSCards WITH(NOLOCK) WHERE [ActivationFlag] = 'N' AND EmployeeID = '" & Session("EmployeeID") & "'"
	'End If
	
Else
	'If Session("UserView") = "All" Then
		strSQL = "SELECT TOP 5 * FROM qryCAPSCards WITH(NOLOCK) WHERE [ActivationFlag] = 'N' AND Status = '00' AND (EmployeeID Like '%" & strSearch & "%' OR FirstName Like '%" & strSearch & "%' OR Surname Like '%" & strSearch & "%')" & strWhere & strSort
	'Else
	'	strSQL = "SELECT top 100 * FROM qryCAPSCards WITH(NOLOCK) WHERE [ActivationFlag] = 'N' AND EmployeeID = '" & Session("EmployeeID") & "' AND (EmployeeID Like '%" & strSearch & "%' OR FirstName Like '%" & strSearch & "%' OR Surname Like '%" & strSearch & "%')"
	'End If
End If

'Build the message displayed at the bottom of the screen with the search details
If Session("UserView") = "All" Then
	strRecordMessage = strSearch
Else
	'strRecordMessage = "for " & Session("UserName") 
End If

objRS.Open strSQL,objCon,3,1

    y = 0

	'Write a message in the list if there are no Unactivated Cards
	If objRS.EOF Then
		Response.Write "<TR><TH colspan=""10"" Style=""text-align:center;"">No Unactivated Cards for " & strRecordMessage & "</TH>" & _
				"<TH colspan=""3"" style=""text-align:center;""></TH></TR>"
	Else
		objRS.Movelast
		objRS.Movefirst
		lngTotalRecords = objRS.Recordcount
		
		
		'Response.Write "<div class=""row""><div class=""col-12"">" & _
		Response.Write	"<table class=""table table-compact text-left""><thead><tr>" & _
			"<th scope=""col""><a href=""UnactivatedCards.asp?Sort=ApplicationID&SortType=" & strOrderType & """> Card ID <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col""><a href=""UnactivatedCards.asp?Sort=EmployeeID&SortType=" & strOrderType & """> EID <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col""><a href=""UnactivatedCards.asp?Sort=Surname&SortType=" & strOrderType & """> Name <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col""><a href=""UnactivatedCards.asp?Sort=CardType&SortType=" & strOrderType & """> Card Type <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col""><a href=""UnactivatedCards.asp?Sort=DateIssued&SortType=" & strOrderType & """> Issued <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col"" title=""Days inactive since issued (up to today)""> Days </th>" & _
			"</tr></thead><tbody class=""text-left"">"
					
	End If
	
	
	'Write a message in the list if there are no applications
	If objRS.EOF Then
		Response.Write "<TR><TH colspan=""10"" Style=""text-align:center;"">No Unactivated Cards " & strRecordMessage & "</TH>" & _
				"<TH colspan=""3"" style=""text-align:center;""></TH></TR>"
	End If
    	
    Do until objRS.EOF 

		y = y + 1
		
		
			x = x + 1
			
			'Create the actions based on the Process Status of the card
			Select Case objRS("ProcessStatus")
			
			Case  "Removed Unactivated"
				strAction = "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" onclick=""self.location='UnactivatedCards.asp?Action=UnRemove&CardID=" & objrs("CardID") & "&CardEID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-minus-list""></i> Re-List</button>"
			Case "Added to CS"

				strAction = "<button type=""button"" class=""btn btn-secondary btn-xs"" onclick=""self.location='../Admin/CAPSAdmin/ExportCS.asp?CardID=" & objrs("CardID") & "&CardEID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-key""></i> View CS</button>"
				
			Case "Email Unactivated"
				strAction = "<button type=""button"" class=""btn btn-danger btn-xs"" onclick=""self.location='UnactivatedCards.asp?Action=Cancel&CardID=" & objrs("CardID") & "&CardEID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-minus-circle""></i> Cancel</button>"
			
			Case Else
				strAction = "<button type=""button"" class=""btn btn-danger btn-xs"" onclick=""self.location='UnactivatedCards.asp?Action=Cancel&CardID=" & objrs("CardID") & "&CardEID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-minus-circle""></i> Cancel</button>"
				strAction = strAction & "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" data-toggle=""modal"" data-target=""#EmailModal""><i class=""fa fa-minus-mail""></i> Email</button>"
				strAction = strAction & "<button type=""button"" class=""btn btn-outline-info btn-xs"" onclick=""self.location='UnactivatedCards.asp?Action=Remove&CardID=" & objrs("CardID") & "&CardEID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-minus-cross""></i> Remove</button>"

				'strAction = strAction & "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" onclick=""self.location='UnactivatedCards.asp?Action=Email&CardID=" & objrs("CardID") & "&CardEID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-minus-mail""></i> Email</button>"
				'data-toggle="modal" data-target="#EmailModal"
				
			End Select
			
			'Create the Status list badge based on the status field
			If IsNull(objRS("Status")) Then
				strStatus = ""
			Else
				If objRS("Status") = "00" Then
					strStatus = "<span class=""badge badge-pill badge-success"">Active</span>"
				ElseIf objRS("Status") = "01" OR objRS("Status") = "02" Then
					strStatus = "<span class=""badge badge-pill badge-danger"">Cancelled</span>"
				Else
					strStatus = ""
				End If
			End If
			
			If IsNull(objRS("DateIssued")) Then
				dteDateSubmitted = ""
			Else
				dteDateSubmitted = FormatDateTime(objRS("DateIssued"),vbShortDate)
			End If
			
			If IsNull(objRS("DateIssued")) Then
				dteDateReviewed = ""
			Else
				dteDateReviewed = DateDiff("d",objRS("DateIssued"),now())
				If dteDateReviewed > 80 Then
					strDaysColour = "Style=""color:red; font-weight:bold;"""
				ElseIf dteDateReviewed > 45 Then
					strDaysColour = "Style=""color:orange; font-weight:bold;"""
				Else
					strDaysColour = "Style=""color:black"""
				End If
			End If
			
			'Format the Card number so it is masked depending on the card type
			If IsNull(objRS("CardNumber")) Then
				strCardNo = ""
			Else
				strCardNo = objRS("CardNumber")
				If mid(strCardNo,5,1)=0 Then 
					strCardNo = mid(strCardNo,6,2) & "****" & right(strCardNo,4)
				Else
					strCardNo = mid(strCardNo,4,2) & "****" & right(strCardNo,4)
				End If
				'If len(strCardNo)>8 Then strCardNo = mid(strCardNo,6,2) & "****" & right(strCardNo,4)
			End If
			
			If IsNull(objRS("ProcessStatus")) or objRS("ProcessStatus") = "" Then
				strProcessStatus = ""
			Else
				If objRS("ProcessStatus") = "Email Unactivated" Then
					strProcessStatus = "Emailed"
				Elseif objRS("ProcessStatus") = "Removed Unactivated" Then
					strProcessStatus = "Removed"
				Elseif objRS("ProcessStatus") = "Added to CS" Then
					strProcessStatus = "Added To CS"
				End If
			End If
			
			dteWarningDate = "title=" & objRS("Warning") & ""
			
			
			If isNull(objRS("CardTypeSub")) Then
				strHeader = ""
				strCardType = ""
			Else
				strHeader = objRS("CardTypeSub")
				strCardType  = objRS("CardType") & " " & objRS("CardTypeSub")
			End If
					
			'Determine the image and title based on the card type
			If Trim(strHeader) = "Diners" Then
				strHeader = "<img height=""20px"" width=""20px"" src=""../images/icon_diners.png"" title=""" & strCardType & """> " & strCardType
			ElseIf strHeader = "ANZ" Then
				strHeader = "<img height=""20px"" width=""20px"" src=""../images/logo_ANZ.png"" Title=""" & strCardType & """> " & strCardType
			ElseIf strHeader = "Mastercard" Then
				strHeader = "<img height=""20px"" width=""20px"" src=""../images/logo_mc.png"" Title=""" & strCardType & """> " & strCardType
			Else
				strHeader = "<img height=""20px"" width=""20px"" src=""../images/logo_coa.png"" Title=""" & strCardType & """> " & strCardType
			End If
			
			
			response.write "<TR><TD ><a Target=""_self"" HREF=""CardDetail.asp?CardID=" & objRS(0) & """>" & objRS(0) & "</a></TD><TD>" & objRS("EmployeeID") & "</a></TD>" & _
					"<TD style=""font-size:12px;""><a Target=""_self"" HREF=""CardDetail.asp?CardID=" & objRS(0) & """>" & objRS("FirstName") & " " & objRS("Surname") & "</a></TD><TD style=""font-size:12px;""><a Target=""_self"" HREF=""CardDetail.asp?CardID=" & objRS(0) & """>" & strHeader & "</a></TD>" & _
					"<TD >" & dteDateSubmitted & "</TD><TD " & strDaysColour & ">" & dteDateReviewed & "</TD></TR>"
					
			
		
		objRS.movenext
	Loop
	
	'If y > 0 Then
		'Response.Write "<TR><TH colspan=""4"">Total</TH>" & _
		'		"<TH colspan=""2"" style=""text-align:center;"">" & x & "</TH></TR></TABLE>"
		
		Response.Write "<TR><TH colspan=""4""><a href=""UnactivatedCards.asp"">More...</a></TH>" & _
				"<TH colspan=""2"" style=""text-align:center;"">" & x & "</TH></TR></TABLE>"
				
	'End If
	
	
objRS.Close

End Sub

Sub DisplayNAFile()

Dim strWhere

If Not IsEmpty(Request.QueryString("BatchNo")) Then
	If IsNull(Request.QueryString("BatchNo")) or Request.QueryString("BatchNo")= "" Then 
		strWhere = "WHERE [BatchNumber] = 0"
	Else
		strWhere = "WHERE FileSeqNum = " & Request.QueryString("BatchNo") & ""
	
		If Request.QueryString("BatchNo") = 0 Then strWhere = strWhere & " OR BatchNumber IS NULL"
	End If
Else
	strWhere = "WHERE [BatchNumber] = 0"
End If

objRS.Open "SELECT TOP 5 * FROM qryCAPSNAToNAB "  & strWhere,objCon
		
		    If objRS.eof Then
		        Response.Write"<table Class=""table"" WIDTH=""100%"" BORDER=""1"" CELLSPACING=""1"" CELLPADDING=""1"">" & _
                "<tr><th colspan=7 style=""text-align:left"">There is no NA To NAB data ready for export</B></th></tr>" & _
		        "<tr><td colspan=7 style=""text-align:left"">Files will be added as part of the admin and load functions and will appear here when appropriate</td></tr>"
		    Else
		    
		         Response.Write"<table Class=""table table-bordered table-hover table-compact mb-0 newd"" cellspacing=""0"" cellpadding=""0"">" & _
		        "<tr><th style=""text-align:center; font-size:12px;"">EmployeeID</th>" & _
				"<th style=""text-align:center; font-size:12px;"">Applicant Name</th>" & _	
		        "<th style=""text-align:center; font-size:12px;"">Status</th>" & _
	 	        "<th style=""text-align:center; font-size:12px;"">Card Type</th>" & _	
		        "<th style=""text-align:center; font-size:12px;"">Updated By</th></tr>"
		        
		    End If
		    
		    Do until objRS.eof
		        'If objRS("GLCode") <> 0 Then
			    Response.Write "<TR class='clickable-row' data-href='../Admin/CAPSAdmin/NATransactions.asp?Link=AD' style=""cursor: pointer;"">" & _
			                    "<TD style=""text-align:center; font-size:12px;"">" & objRS("EIDNo") & "</TD><TD style=""text-align:center; font-size:12px;"">" & objRS("FirstName") & " " & objRS("Surname") & "</TD>" & _
			                    "<TD style=""text-align:center; font-size:12px; font-size:12px;"">" & objRS("Status") & "</TD><TD style=""text-align:center"">" & objRS("CardTypeSub") & "</TD>" & _
			                    "<TD style=""text-align:center; font-size:12px;"">" & objRS("UpdatedByName") & "</TD>" & _
			                    "</TR>"
    			'End If
    			
			    objRS.movenext
		    Loop
    			
	    objRS.Close

        Response.Write "</table>"
		
End Sub


Public Sub GetSummaryStats()
'Procedure to get summary stats for one of the Summary panes

	objRS.Open "SELECT COUNT(*) as [CountAll] FROM qryCAPSCSToNAB WITH(NOLOCK) WHERE (FileSeqNum Is NULL OR FileSeqNum = '') AND [Status] <> 'Deleted'",objCon
		
		If objRS.eof Then
			lngCSRecords = 0
		Else
			lngCSRecords = objRS("CountAll")
		End If
		     
	 objRS.Close

	objRS.Open "SELECT COUNT(*) as [CountAll] FROM qryCAPSCSToNAB WITH(NOLOCK) WHERE ([CardStatus] = 'VX') AND [Status] = 'Awaiting Export'",objCon
		
		If objRS.eof Then
			lngCSCancels = 0
		Else
			lngCSCancels = objRS("CountAll")
		End If
		     
	 objRS.Close
	 
	 objRS.Open "SELECT COUNT(*) as [CountAll] FROM qryCAPSNAToNAB WITH(NOLOCK) WHERE ([BatchNumber] = 0 OR [BatchNumber] = '' OR [BatchNumber] Is NULL) AND [Status] <> 'Deleted'",objCon
		
		If objRS.eof Then
			lngNewCards = 0
		Else
			lngNewCards = objRS("CountAll")
		End If
		     
	 objRS.Close
	 
	 
	 objRS.Open "SELECT COUNT(*) as [CountAll] FROM qryCAPSNAToNAB WITH(NOLOCK) WHERE [Status] = 'Added To NA' AND CardType = 'DPC'",objCon
		
		If objRS.eof Then
			lngANZApps = 0
		Else
			lngANZApps = objRS("CountAll")
		End If
		     
	 objRS.Close
End Sub


Set objRS = Nothing
Set objCon = Nothing

%>