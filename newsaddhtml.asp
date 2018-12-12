<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/NewsDataBase.asp" -->
<%
' *** Restrict Access To Page: Grant or deny access to this page

MM_authorizedUsers=""
MM_authFailedURL="default.asp?p=badpassword"
MM_grantAccess=false
If Session("MM_Username") <> "" Then
  If (true Or CStr(Session("MM_UserAuthorization"))="") Or _
         (InStr(1,MM_authorizedUsers,Session("MM_UserAuthorization"))>=1) Then
    MM_grantAccess = true
  End If
End If
If Not MM_grantAccess Then
  MM_qsChar = "?"
  If (InStr(1,MM_authFailedURL,"?") >= 1) Then MM_qsChar = "&"
  MM_referrer = Request.ServerVariables("URL")
  if (Len(Request.QueryString()) > 0) Then MM_referrer = MM_referrer & "?" & Request.QueryString()
  MM_authFailedURL = MM_authFailedURL & MM_qsChar & "accessdenied=" & Server.URLEncode(MM_referrer)
  Response.Redirect(MM_authFailedURL)
End If
%>
<%
set rsNewNews = Server.CreateObject("ADODB.Recordset")
rsNewNews.ActiveConnection = MM_NewsDataBase_STRING
rsNewNews.Source = "SELECT *  FROM tblnews"
rsNewNews.CursorType = 0
rsNewNews.CursorLocation = 2
rsNewNews.LockType = 3
rsNewNews.Open()
rsNewNews_numRows = 0
%>
<%
Dim rsUser__MMColParam
rsUser__MMColParam = "1"
if (Session("MM_username") <> "") then rsUser__MMColParam = Session("MM_username")
%>
<%
set rsUser = Server.CreateObject("ADODB.Recordset")
rsUser.ActiveConnection = MM_NewsDataBase_STRING
rsUser.Source = "SELECT * FROM tbluser WHERE USER = '" + Replace(rsUser__MMColParam, "'", "''") + "'"
rsUser.CursorType = 0
rsUser.CursorLocation = 2
rsUser.LockType = 3
rsUser.Open()
rsUser_numRows = 0
%>
<%
set rsUser2 = Server.CreateObject("ADODB.Recordset")
rsUser2.ActiveConnection = MM_NewsDataBase_STRING
rsUser2.Source = "select * from CATEGORY"
rsUser2.CursorType = 0
rsUser2.CursorLocation = 2
rsUser2.LockType = 3
rsUser2.Open()
%>
<html>
<head>
<title>Add News Article</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="css_styles/site.css" rel="stylesheet" type="text/css">
<!-- START : EDITOR HEADER - INCLUDE THIS IN ANY FILES USING EDITOR -->
<script language="Javascript1.2" src="htmled/editor.js"></script>
<script>
_editor_url = "htmled/";
</script>
<style type="text/css"><!--
  .btn   { BORDER-WIDTH: 1; width: 26px; height: 24px; }
  .btnDN { BORDER-WIDTH: 1; width: 26px; height: 24px; BORDER-STYLE: inset; BACKGROUND-COLOR: buttonhighlight; }
  .btnNA { BORDER-WIDTH: 1; width: 26px; height: 24px; filter: alpha(opacity=25); }
--></style>
</head>
<body bgcolor="#FFFFFF" text="#000000">
<table width="100%" border="0" class="AdminTable">
  <tr> 
    <td width="95%" class="HeaderRow">Add News Article</td>
  </tr>
  <tr> 
    <form name="fHtmlEditor" method="POST" action="newsaddfinal.asp">
      <td width="95%"> <table width="54%" border="0">
          <tr> 
            <td width="40%" class="SubHeaderRow">Subject</td>
            <td width="30%"> 
            <input name="subject" type="text" class="inquiryform" size="20"> 
              <input type="hidden" name="user" value="<%=(rsUser.Fields.Item("USER").Value)%>"> 
            </td>

			 <tr> 
            <td width="77" class="SubHeaderRow" >Related Link</td>
            <td colspan="2"> 
            <input name="related" type="text" class="inquiryform" size="20"> 
            </td>
          </tr>
		  <tr>
		  <tr>  <td width="17%" class="SubHeaderRow"> <div align="left">Category</div></td>
			<td width="83%"><select name="Category">
			<%
			Set objConn=Server.CreateObject("ADODB.Connection")
			objConn.Provider="Microsoft.Jet.OLEDB.4.0"
			objConn.Open "d:\hosts\greenvillemarine.com\www\fpdb\news2.mdb"

			Set objRec = Server.CreateObject("ADODB.RecordSet")
			objRec.Open "SELECT * FROM CATEGORY ORDER BY CATEGORY", objConn
			While not objRec.EOF
				If InStr(rsUser("CATEGORY"),objRec("id")) Then
					Response.Write "<option value=""" & objRec("id") & """>" & objRec("CATEGORY")
				End If
				objRec.MoveNext
			Wend	
			objRec.Close
			Set objRec = NOTHING
			objConn.Close
			Set objConn = NOTHING
			%>
            </td>
          </tr>

            <td width="30%"><input name="Submit" type="submit" class="submit" value="Save"></td>
          </tr>
        </table>
        <textarea name="EditorValue" cols="80" rows="25"></textarea><script language="javascript1.2">
editor_generate('EditorValue'); // field, width, height
      </script> </td>
    </form>
  </tr>
</table>

<p align="center">
</p>
 
</center></body>
</html>
<%
rsNewNews.Close()
%>
<%
rsUser.Close()
%>