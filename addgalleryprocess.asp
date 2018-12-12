<!--#include file="inc/incsettings.asp"-->
<!--#include file="inc/db_conn_open.asp"-->
<%
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open sDSN

objConn.Execute "INSERT INTO galleries (galleryName) VALUES ('" & SQLInject(Request.Form("galleryName")) & "')"

objConn.Close
Set objConn = NOTHING

Response.Redirect "gallery.asp"
%>