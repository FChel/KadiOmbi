<%


Dim objCon
Dim objRS
Dim strSQL
Dim strData
Dim strWHERE
Dim arrData(5)
Dim strMonth
Dim strYear

Set objCon = Server.CreateObject("ADODB.Connection")
Set objRS = Server.CreateObject("ADODB.Recordset")

objCon.Open Session("DBConnection")	

'Sub LoadDetails()

If IsNull(Session("ReportYear")) OR Session("ReportYear") = "" Then
	strYear = "2020"
Else
	strYear = Session("ReportYear")
End If

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
'strWHERE = "112020"


'strSQL = "SELECT COUNT([EmployeeID]) As Apps,CardType,[CardTypeSub],CAST(CAST(RIGHT('00'+CAST(Month(DateSubmitted) as varchar(2)),2) as varchar(2)) + '' + CAST(YEAR(DateSubmitted) as varchar(20)) AS Char(6)) As MonthYear " & _
	'"FROM tblCAPSApplication WITH(NOLOCK) " & _
	'"WHERE CAST(CAST(RIGHT('00'+CAST(Month(DateSubmitted) as varchar(2)),2) as varchar(2)) + '' + CAST(YEAR(DateSubmitted) as varchar(20)) AS Char(6)) = '" & strWHERE & "' " & _
	'"GROUP BY [CardType],[CardTypeSub],CAST(CAST(RIGHT('00'+CAST(Month(DateSubmitted) as varchar(2)),2) as varchar(2)) + '' + CAST(YEAR(DateSubmitted) as varchar(20)) AS Char(6))"
strSQL = "SELECT COUNT([EmployeeID]) As Apps,CardType,[ApplicationTypeName],[CardTypeSub],CAST(CAST(RIGHT('00'+CAST(Month(DateSubmitted) as varchar(2)),2) as varchar(2)) + '' + CAST(YEAR(DateSubmitted) as varchar(20)) AS Char(6)) As MonthYear " & _
	"FROM tblCAPSApplication WITH(NOLOCK) " & _
	"WHERE CAST(CAST(RIGHT('00'+CAST(Month(DateSubmitted) as varchar(2)),2) as varchar(2)) + '' + CAST(YEAR(DateSubmitted) as varchar(20)) AS Char(6)) = '" & strWHERE & "' " & _
	"GROUP BY [CardType],[ApplicationTypeName],[CardTypeSub],CAST(CAST(RIGHT('00'+CAST(Month(DateSubmitted) as varchar(2)),2) as varchar(2)) + '' + CAST(YEAR(DateSubmitted) as varchar(20)) AS Char(6))"

   'Description:	Loads Position details into page if applicable.
	objRS.Open strSQL,objCon

		arrData(1) = 0
		arrData(2) = 0
		arrData(3) = 0
		arrData(4) = 0
		arrData(5) = 0
		
		If Not objRS.EOF Then
			
			Do Until objRS.EOF
				
				''New NAB cards data
				If objRS("CardType") = "DTC" AND objRS("ApplicationTypeName") = "DTC NAB CiH" Then
					arrData(1) =  arrData(1) + objRS("Apps")
				End If
				
				If objRS("CardType") = "DTC" AND objRS("ApplicationTypeName") = "DTC NAB Lodge" Then
					arrData(2) =  arrData(2) + objRS("Apps")
				End If
				
				If objRS("CardType") = "DTC" AND objRS("ApplicationTypeName") = "DTC NAB Dual" Then
					arrData(3) =  arrData(3) + objRS("Apps")
				End If
				
				If objRS("CardType") = "DPC" AND Left(objRS("CardTypeSub"),3) = "NAB" Then
					arrData(4) =  arrData(4) + objRS("Apps")
				End If
				
				If Right(objRS("ApplicationTypeName"),12) = "Limit Change" Then
					arrData(5) =  arrData(5) + objRS("Apps")
				End If
				
				'If objRS("CardTypeSub") = "CTS" Then
				'	arrData(5) =  arrData(5) + objRS("Apps")
				'End If
				
				'''Old Diners and ANZ card data
				'If objRS("CardTypeSub") = "Diners" Then
				'	arrData(1) =  arrData(1) + objRS("Apps")
				'End If
				
				'If objRS("CardTypeSub") = "ANZ" Then
				'	arrData(2) =  arrData(2) + objRS("Apps")
				'End If
				''''New NOV 2023 to include a count for Diners DPC Mastercards
				'If objRS("CardType") = "DPC" AND objRS("CardTypeSub") = "Mastercard" Then
				'	arrData(2) =  arrData(2) + objRS("Apps")
				'End If
				
				'If objRS("CardType") = "DTC" AND objRS("CardTypeSub") = "Mastercard" Then
				'	arrData(3) =  arrData(3) + objRS("Apps")
				'End If
				
				'If objRS("CardTypeSub") = "CTS" Then
				'	arrData(4) =  arrData(4) + objRS("Apps")
				'End If
				
				strData = strData & "," & objRS("Apps")
			
				objRS.Movenext
			Loop
		Else
			strData = ","
	   End If
	 
	objRS.Close
	
	strData = arrData(1) & "," & arrData(2) & "," & arrData(3) & "," & arrData(4) & "," & arrData(5)
	'strData = Right(strData,Len(strData)-1)
	
	'response.write strData
	
'End Sub

Function PadDigits(val, digits)
  PadDigits = Right(String(digits,"0") & val, digits)
End Function

%>
<!doctype html>
<html>

<head>
	<title>Bar Chart</title>
	<script src="Chart.min.js"></script>
	<script src="utils.js"></script>
	<style>
	canvas {
		-moz-user-select: none;
		-webkit-user-select: none;
		-ms-user-select: none;
	}
	</style>
</head>

<body>
	<div id="container" style="width: 100%;">
		<canvas id="canvas"></canvas>
	</div>
	
	<script>
		var MONTHS = ['DTC CiH', 'Lodge', 'Dual', 'DPC', 'Limit Change'];
		var color = Chart.helpers.color;
		var barChartData = {
			labels: ['DTC CiH', 'Lodge', 'Dual', 'DPC', 'Limit Change'],
			datasets: [{
				label: 'Card Applications',
				backgroundColor: color(window.chartColors.blue).alpha(0.5).rgbString(),
				//borderColor: window.chartColors.blue,
				borderWidth: 1,
				data: [
					<%=strData%>
										
				],
				backgroundColor: [
						window.chartColors.blue,
						window.chartColors.purple,
						window.chartColors.green,
						window.chartColors.red,
						window.chartColors.yellow
						]
			}, ]

		};

		window.onload = function() {
			var ctx = document.getElementById('canvas').getContext('2d');
			window.myBar = new Chart(ctx, {
				type: 'bar',
				data: barChartData,
				options: {
					responsive: true,
					legend: {
						position: 'top',
					},
					title: {
						display: false,
						text: 'Cards'
					}
				}
			});

		};

		document.getElementById('randomizeData').addEventListener('click', function() {
			var zero = Math.random() < 0.2 ? true : false;
			barChartData.datasets.forEach(function(dataset) {
				dataset.data = dataset.data.map(function() {
					return zero ? 0.0 : randomScalingFactor();
				});

			});
			window.myBar.update();
		});

		var colorNames = Object.keys(window.chartColors);
		document.getElementById('addDataset').addEventListener('click', function() {
		
			var colorName = colorNames[barChartData.datasets.length % colorNames.length];
			var dsColor = window.chartColors[colorName];
			var newDataset = {
				label: 'Dataset ' + (barChartData.datasets.length + 1),
				backgroundColor: color(dsColor).alpha(0.5).rgbString(),
				borderColor: dsColor,
				borderWidth: 1,
				data: []
			};

			for (var index = 0; index < barChartData.labels.length; ++index) {
				newDataset.data.push(randomScalingFactor());
			}

			barChartData.datasets.push(newDataset);
			window.myBar.update();
		});

		document.getElementById('addData').addEventListener('click', function() {
			
			if (barChartData.datasets.length > 0) {
				var month = MONTHS[barChartData.labels.length % MONTHS.length];
				barChartData.labels.push(month);

				for (var index = 0; index < barChartData.datasets.length; ++index) {
					// window.myBar.addData(randomScalingFactor(), index);
					barChartData.datasets[index].data.push(randomScalingFactor());
				}

				window.myBar.update();
			}
		});

		document.getElementById('removeDataset').addEventListener('click', function() {

			barChartData.datasets.pop();
			window.myBar.update();
		});

		document.getElementById('removeData').addEventListener('click', function() {
			barChartData.labels.splice(-1, 1); // remove the label first

			barChartData.datasets.forEach(function(dataset) {
				dataset.data.pop();
			});

			window.myBar.update();
		});
	</script>
</body>

</html>
