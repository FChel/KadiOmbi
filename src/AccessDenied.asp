<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<%@ Language=VBScript %>
<%

Dim strUserID
Dim objCon
Dim objRS
Dim strScreenName

Set objCon = Server.CreateObject("ADODB.Connection")
Set objRS = Server.CreateObject("ADODB.Recordset")	

objCon.Open Session("DBConnection")

    objRS.Open "SELECT * FROM tblScreens WHERE PageName = '" & Session("CurrentPage") & "'",objCon
    
        If Not objRS.EOF Then
            strScreenName = objRS("ScreenName")
        Else
            strScreenName = "NOT DEFINED."
        End IF
        
    objRS.Close	

  

strUserID = Right(Request.ServerVariables("Auth_User"),6)

 %>
<html>
<head>
<script LANGUAGE="javascript">
<!--


function GoToHome()

{
	
	//self.location="FinStatementConfig.asp?Action=Delete&LevelID="+frm.LevelID.value+"&Level2ID="+frm.RLLevel2ID.value+"&Level1ID="+frm.RLLevel1ID.value+"&TransactionID="+frm.TransactionID.value;
    top.location="Index.asp";
	//alert("../BERTFrameset.asp?CostCentreID="+frm.BidCostCentreID.value+"&CurrentPage=ProfitLoss/PLCostCentre1.asp&BusinessAreaID="+frm.BusinessAreaID.value+"&HeaderMenu=Header2.asp");

}

//-->
</script>
<title>Isidore SME</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="style/login_style.css" type="text/css" rel="stylesheet">
</head>
<body style="background-color:#f8f8f8">
<h1 Style="Font-Family:Arial;Color:Red;">&nbsp;&nbsp;ACCESS DENIED !</h1>
    <img alt="" src="images/lock.png" />
<h4 Style="Font-Family:Arial;Color:Black;">&nbsp;&nbsp;Your User Role does not allow access to the screen <U><%=strScreenName %>.</U></h4>
<br />
<h4 Style="Font-Family:Arial;Color:Black;">&nbsp;&nbsp;Press the Continue.. button to return to the Home Screen.</h4>

&nbsp;&nbsp;<input id="Submit1" type="submit" value="Continue..." onclick="GoToHome()">
</body
</html>