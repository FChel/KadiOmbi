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

'Sub LoadDetails()

'strWHERE = "112020"

strSQL = "SELECT COUNT([EmployeeID]) As Apps,CardType, ApplicationTypeName,[CardTypeSub],CAST(CAST(RIGHT('00'+CAST(Month(DateSubmitted) as varchar(2)),2) as varchar(2)) + '' + CAST(YEAR(DateSubmitted) as varchar(20)) AS Char(6)) As MonthYear " & _
	"FROM tblCAPSApplication " & _
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
				
					'If objRS("CardTypeSub") = "Mastercard" Then
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


%>
<!doctype html>
<html>

<head>
	<title>Doughnut Chart</title>
	<script src="Chart.min.js"></script>
	<script src="utils.js"></script>
	<style>
	canvas {
		-moz-user-select: none;
		-webkit-user-select: none;
		-ms-user-select: none;
	}
	</style>
	
<script>
function clickGoto {
	alert('feerf fdf');
}

</script>
</head>

<body>
	<div id="canvas-holder" style="width:100%;">
		<canvas id="chart-area" style="width:100%; height:100%; align-content: left;"></canvas>
	</div>
	<script>
		var randomScalingFactor = function() {
			return Math.round(Math.random() * 100);
		};

		var config = {
			type: 'doughnut',
			data: {
				datasets: [{
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
				}], 
			labels: [
        'DTC CiH',
        'Lodge', 'Dual',
        'DPC','Credit Limit'
    ]
				
			},
			options: {
				responsive: false,

				title: {
					display: false,
					text: 'Applications'
				},
				animation: {
					animateScale: true,
					animateRotate: true
				},
				legend: { position: 'none' }
				
			}
		};

		window.onload = function() {
			var ctx = document.getElementById('chart-area').getContext('2d');
			window.myDoughnut = new Chart(ctx, config);
		};

		document.getElementById('randomizeData').addEventListener('click', function() {
			config.data.datasets.forEach(function(dataset) {
				dataset.data = dataset.data.map(function() {
					return randomScalingFactor();
				});
			});

			window.myDoughnut.update();
		});

		var colorNames = Object.keys(window.chartColors);
		document.getElementById('addDataset').addEventListener('click', function() {
			var newDataset = {
				backgroundColor: [],
				data: [],
				label: 'New dataset ' + config.data.datasets.length,
			};

			for (var index = 0; index < config.data.labels.length; ++index) {
				newDataset.data.push(randomScalingFactor());

				var colorName = colorNames[index % colorNames.length];
				var newColor = window.chartColors[colorName];
				newDataset.backgroundColor.push(newColor);
			}

			config.data.datasets.push(newDataset);
			window.myDoughnut.update();
		});

		document.getElementById('addData').addEventListener('click', function() {
			if (config.data.datasets.length > 0) {
				config.data.labels.push('data #' + config.data.labels.length);

				var colorName = colorNames[config.data.datasets[0].data.length % colorNames.length];
				var newColor = window.chartColors[colorName];

				config.data.datasets.forEach(function(dataset) {
					dataset.data.push(randomScalingFactor());
					dataset.backgroundColor.push(newColor);
				});

				window.myDoughnut.update();
			}
		});

		document.getElementById('removeDataset').addEventListener('click', function() {
			config.data.datasets.splice(0, 1);
			window.myDoughnut.update();
		});

		document.getElementById('removeData').addEventListener('click', function() {
			config.data.labels.splice(-1, 1); // remove the label first

			config.data.datasets.forEach(function(dataset) {
				dataset.data.pop();
				dataset.backgroundColor.pop();
			});

			window.myDoughnut.update();
		});

		document.getElementById('changeCircleSize').addEventListener('click', function() {
			if (window.myDoughnut.options.circumference === Math.PI) {
				window.myDoughnut.options.circumference = 2 * Math.PI;
				window.myDoughnut.options.rotation = -Math.PI / 2;
			} else {
				window.myDoughnut.options.circumference = Math.PI;
				window.myDoughnut.options.rotation = -Math.PI;
			}

			window.myDoughnut.update();
		});
		
	
	
	</script>
</body>

</html>