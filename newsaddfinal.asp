<!--#include file="Connections/NewsDataBase.asp" -->
<!--#include file="adovbs.inc" -->

<%
Dim article

if Request.Form("EditorValue") = "" then
article = Request.Form("text")
else
article = Request.Form("EditorValue")
end if

user = Request.Form("user")
text = article
related = Request.Form("related")
subject = Request.Form("subject")
category = Request.Form("category")

Set ConnObj = Server.CreateObject("ADODB.Connection")
ConnStr = MM_NewsDataBase_STRING
ConnObj.Open ConnStr

set rst = Server.CreateObject("ADODB.RecordSet")

rst.Open "tblnews",ConnObj,adOpenDynamic,adLockOptimistic

rst.AddNew 

With rst
	.Fields("USERCREATED") = user
	.Fields("BODY") = text
	.Fields("SUBJECT") = subject
	.Fields("RELATEDLINK") = related
	.Fields("CATEGORY") = category
End With

rst.Update

rst.Close

Set rst = nothing

ConnObj.Close

Set ConnObj  =nothing

Response.Redirect "admin.asp"
%>
