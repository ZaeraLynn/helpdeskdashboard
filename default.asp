<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"> 
	<head>
		<title>Help Desk Ticket Dashboard</title>
		<style type="text/css">
			body{
				font-family: helvetica;
				font-size: 13px;
			}
			
			.yellow{
				background-color: #ffff00;
			}
			
			.orange{
				background-color: #ff9900;
			}
			
			.red{
				--background-color: #ff3333;
				background-color: #FF8080;
			}
			
			.green{
				--background-color: #33ff66;
				background-color: #80FF9F;
			}
			
			
			table.tasklistleft{
				border: 1px #333333 solid;
				border-collapse: collapse;
				width: 48%;
				float: left;
			}
			
			table.tasklistright{
				border: 1px #333333 solid;
				border-collapse: collapse;
				width: 48%;
				float: right;
			}
			
			table.tasklistleft td, th{
				border-top: 1px #333333 solid;
				border-bottom: 1px #333333 solid;
				border-right: 1px #333333 solid;
				padding: 5px;
			}
			
			table.tasklistright td, th{
				border-top: 1px #333333 solid;
				border-bottom: 1px #333333 solid;
				border-right: 1px #333333 solid;
				padding: 5px;
			}
	
		</style>
		<meta http-equiv="Refresh" content="300">
	</head>
	<body>
<%


' ******* CONNECT TO CRM SQL DATABASE ********
'SERVER IP: 192.168.0.##
connString = "Provider=SQLNCLI10;Data Source=192.168.0.##\INSTANCENAME;Initial Catalog=DATABASENAME;User ID=SA;Password=********"
set Conn = Server.CreateObject("ADODB.Connection")
Conn.ConnectionString = dbString
Conn.Open

' Create recordset of MSP tickets
set rs = Server.CreateObject("ADODB.recordset")
query = "SELECT tblServiceOrders.SONumber, tblServiceOrders.Priority, tblServiceOrders.Status, tblAccounts.AccountName, tblReps.RepName, tblServiceOrders.BriefDescription, tblServiceOrders.DateReceived, tblServiceOrders.TimeReceived FROM tblServiceOrders, tblAccounts, tblReps WHERE (tblServiceOrders.Status LIKE 'MSP- Open' OR tblServiceOrders.Status LIKE 'MSP- New' OR tblServiceOrders.Status LIKE 'MSP- Alert' OR tblServiceOrders.Status LIKE 'MSP- Escalation') AND UPPER(tblServiceOrders.BriefDescription) NOT LIKE '%ALERT%' AND tblServiceOrders.DateClosed IS NULL AND tblServiceOrders.AccountNumber = tblAccounts.AccountNumber AND tblServiceOrders.TechAssigned = tblReps.RepNumber ORDER BY SONumber"
rs.Open query, Conn

if NOT rs.EOF then
	rs.MoveFirst
%>
		<h1>Help Desk Ticket Dashboard</h1>(Refreshes every 5 minutes)<br/><br/>
		<table class="tasklistleft">
			<tr>
				<td colspan="4">
					<h1>Open MSP Tickets:
<%
	rsArray = rs.GetRows()
	nr = UBound(rsArray, 2) + 1 
	Response.Write(nr)
	rs.MoveFirst
%>
					</h1><br/><br/>
					<h1>MSP Help Desk Tickets</h1>
				</td>
			</tr>
			<tr>
				<th>Status</th>
				<th>Account</th>
				<th>Description</th>
				<th>Opened Date</th>
			</tr>

<%
End If

' Construct list of tickets, status, client names, etc. with color coding based on status / urgency
Do While Not rs.EOF
	
	' Set row class based on status / urgency
	if(rs("Status") = "MSP- Escalation") Then
		Response.Write("<tr class=""red"">")
	elseif(rs("Status") = "HD- Open" OR rs("Status") = "MSP- Open") Then
		Response.Write("<tr class=""green"">")
	else
		Response.Write("<tr>")
	end if
	
	' Output ticket details
	Response.Write("<td>" & rs("Status") & "</td>")
	Response.Write("<td>" & rs("AccountName") & "</td>")
	Response.Write("<td>" & Server.HTMLEncode(rs("BriefDescription")) & "</td>")
	Response.Write("<td>" & rs("DateReceived") & "</td>")
	Response.Write("</tr>")
	rs.MoveNext
Loop

rs.Close
%>
		</table>
<%

' Create recordset of Break / Fix tickets
set rs = Server.CreateObject("ADODB.recordset")
query = "SELECT tblServiceOrders.SONumber, tblServiceOrders.Priority, tblServiceOrders.Status, tblAccounts.AccountName, tblReps.RepName, tblServiceOrders.BriefDescription, tblServiceOrders.DateReceived, tblServiceOrders.TimeReceived FROM tblServiceOrders, tblAccounts, tblReps WHERE (tblServiceOrders.Status LIKE 'HD- New' OR tblServiceOrders.Status LIKE 'HD- Open') AND UPPER(tblServiceOrders.BriefDescription) NOT LIKE '%ALERT%' AND tblServiceOrders.DateClosed IS NULL AND tblServiceOrders.AccountNumber = tblAccounts.AccountNumber AND tblServiceOrders.TechAssigned = tblReps.RepNumber ORDER BY SONumber"
rs.Open query, Conn

if NOT rs.EOF then
	rs.MoveFirst
%>
		<table class="tasklistright">
			<tr>
				<td colspan="4">
					<h1>Open HD Tickets:
<%
	rsArray = rs.GetRows()
	nr = UBound(rsArray, 2) + 1 
	Response.Write(nr)
	rs.MoveFirst
%> 
					</h1><br/><br/>
					<h1>Help Desk Tickets</h1>
				</td>
			</tr>
			<tr>
				<th>Status</th>
				<th>Account</th>
				<th>Description</th>
				<th>Opened Date</th>
			</tr>
<%
End If

' Construct list of tickets, status, client names, etc. with color coding based on status / urgency
Do While Not rs.EOF
	
	' Set row class based on status / urgency
	if(rs("Status") = "MSP- Escalation") Then
		Response.Write("<tr class=""red"">")
	elseif(rs("Status") = "HD- Open" OR rs("Status") = "MSP- Open") Then
		Response.Write("<tr class=""green"">")
	else
		Response.Write("<tr>")
	end if
	
	' Output ticket details
	Response.Write("<td>" & rs("Status") & "</td>")
	Response.Write("<td>" & rs("AccountName") & "</td>")
	Response.Write("<td>" & Server.HTMLEncode(rs("BriefDescription")) & "</td>")
	Response.Write("<td>" & rs("DateReceived") & "</td>")
	Response.Write("</tr>")
	rs.MoveNext
Loop

rs.Close
%>
		</table>

<%
Set rs = Nothing
%>
	</body>
</html>