<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<%@ Language=VBScript %>

<%  
    'File:			Default.asp
	'Written By:	Michael Giacomin
	'Written On:	April 2016
	'Edit History:	
	'Purpose:	Report used in the dashboard

Response.Expires = -1500

	Dim objCon
	Dim objRS
	Dim strString

	Dim intCars, intDays, dblEmpCont
	Dim dblTotal
	Dim lngApplications
	
	Set objCon = Server.CreateObject("ADODB.Connection")
	Set objRS = Server.CreateObject("ADODB.Recordset")

	'Session("DBConnection") = "File Name=" & Server.MapPath("Database/IsidoreSME.udl") & ";"
	
	objCon.Open Session("DBConnection")

	If Not IsEmpty(Request.QueryString("Chart")) Then
		Session("Chart") = Request.QueryString("Chart")
	End If

   

%>
<html>
  <head>
    <title>
      Card Applications
    </title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
     <script src="../HighCharts/jquery-1.9.1.js" type="text/javascript"></script>
    <script src="../HighCharts/highcharts.js" type="text/javascript"></script>
    <script src="../HighCharts/exporting.js" type="text/javascript"></script>
    <script type="text/javascript">
        $(function() {
            $('#container').highcharts({
                title: {
                    text: "Card Applications",
                    x: -20 //center
                },
		subtitle: {
                    text: 'Credit Card Applications Awaiting Review',
                    x: -20
                },
		chart: {
                backgroundColor: '#f8f8f8',
            	type: 'column'
        	},
                
                xAxis: {
                    categories: ['Card Applications Awaiting Review']
                },
                yAxis: {
                    title: {
                        text: 'Applications'
                    },
                    plotLines: [{
                        value: 0,
                        width: 1,
                        color: '#808080'
                    }]
                },
                tooltip: {
                    valueSuffix: ' Applications'
                },
				plotOptions: {
        bar: {
            dataLabels: {
                enabled: true
            }
        }
    },
                legend: {
                    layout: 'vertical',
                    align: 'right',
                    verticalAlign: 'middle',
                    borderWidth: 0
                },
                series: [{
		<%
   
	 
	'Get the default BudgetID from system Defaults table.
   	 'Response.write "SELECT * FROM qryCarParkingStaffing WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND BusinessAreaID = " & Session("BusinessAreaID") & " AND StaffingClassificationID = " & Session("StaffingClassificationID") & " Order by CostCentreID"
   	 'objRS.Open "SELECT * FROM qryCarParkingStaffing WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND BusinessAreaID = " & Session("BusinessAreaID") & " Order by BenefitType",objCon
	 'objRS.Open "SELECT * FROM qryCarParkingStaffing WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND BusinessAreaID = " & Session("BusinessAreaID") & " AND StaffingClassificationID = " & Session("StaffingClassificationID") & " Order by BenefitType",objCon
	 
	 objRS.Open "SELECT * FROM qryCAPSApplicationSummary WITH(NOLOCK)",objCon
	 
	If objRS.EOF Then 
		strString = "] } , {    "

	Else
	
		Do Until objRS.EOF
			
		
			strString = strString & " name: '" & objRS("CardType") & "'," & _
				"data: [" & objRS("Applications") & "] } , {"

		objRS.Movenext
		Loop
	
	End If
	
	'strstring = "name: 'DTC', data: [234] }, {name: 'DPC', data: [125] }, {name: 'MC', data: [58] }...."
	
	objRS.close
	Response.Write Left(strString, Len(strString) -4)
%>

		]
            });
        });
    </script>
  </head>
  <body>

    <div id="container" style="min-width: 300px; height: 480px; margin: 0 auto"> </div>

  </body>
</html>