<!--#include file="administration/inc/db_conn_open.asp"-->
<%
Set Jpeg = Server.CreateObject("Persits.Jpeg")

Jpeg.Open Server.MapPath(fPath & "\" & Request.QueryString("filename"))

L = 75

Jpeg.Width = L
Jpeg.Height = Jpeg.OriginalHeight * L / Jpeg.OriginalWidth

Jpeg.SendBinary
%> 