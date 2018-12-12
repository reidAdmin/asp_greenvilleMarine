<!--#include file="administration/inc/db_conn_open.asp"-->
<%
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open sDSN

Set objRec = Server.CreateObject("ADODB.RecordSet")
objRec.Open "SELECT * FROM pics WHERE picId = " & SQLInject(Request.QueryString("id")), objConn
%>
<html>
<head><title>View Pic</title>
<LINK HREF="style.css" TYPE="text/css" REL="stylesheet"> 
</head>
<body style="background-color:#FFF;">
<center>
<table cellpadding="0" cellspacing="0" border="0" width="90%">
  <tr>
    <td rowspan="4">&nbsp;&nbsp;&nbsp;</td>
	<td>&nbsp;</td></tr>
  <tr>
    <td><%= objRec("picDescription") %></td></tr>
  <tr>
    <td><hr size="1" color="#000000"></td></tr>
  <tr>
    <td style="text-align:center;vertical-align:middle"><img src="<%= Replace(fPath,"\","") %>/<%= objRec("picFile") %>" border="0"></td></tr>
</table>
</center>
</body>
</html>
<%
objRec.Close
Set objRec = NOTHING

objConn.Close
Set objConn = NOTHING
%>