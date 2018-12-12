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
</head>
<body bgcolor="#FFFFFF" text="#000000">
<Font Face="<%=appFont%>" Size="-1">

<Center>
	<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber1">
      <tr>
        <td width="100%" background="images/gmarineadminback.gif">
        <p align="center"><a href="staff/default.htm">
        <img border="0" src="images/gmarineadmin.gif"></a></td>
      </tr>
      <tr>
        <td width="100%">&nbsp;</td>
      </tr>
    </table>
</Font>
</Center>


<table width="100%" border="0" class="AdminTable">
  <tr> 
    <td width="95%" class="HeaderRow" >Add News Article</td>
  </tr>
  <tr> 
    <td width="95%"> <form name="AddNews" method="POST" action="newsaddfinal.asp">
        <table width="54%" border="0">
          <!--DWLayoutTable-->
          <tr> 
            <td width="77" class="SubHeaderRow" >Subject</td>
            <td colspan="2"> 
            <input name="subject" type="text" class="inquiryform" size="20"> 
            </td>
          </tr>
		  			
            <tr>  <td width="17%" class="SubHeaderRow"> <div align="right">Category</div></td>
			<td width="83%">
			<%
			dim x
			x = rsUser("id")
			%>
			<select name="CATEGORY">
			<%While Not rsUser2.EOF %>
			<option value="<%= rsUser2("id") %>"> <% If InStr(rsUser("Category"),rsUser2("id")) Then response.write rsUser2("CATEGORY")  End If 
			    rsUser2.movenext
				wend
			%>
            </td>
          </tr>
          <tr> 
            <td width="77" class="SubHeaderRow" >Article Text<br>
              (supports HTML code)</td>
            <td colspan="2"> <textarea name="text" cols="65" rows="5" class="inquiryform"></textarea> 
            </td>
          </tr>
          <tr> 
            <td width="77" class="SubHeaderRow" >Related Link</td>
            <td colspan="2"> 
            <input name="related" type="text" class="inquiryform" size="20"> 
            </td>
          </tr>
          <tr> 
            <td width="77" height="20"> <font color="#FFFFFF"> 
              <input type="hidden" name="user" value="<%=(rsUser.Fields.Item("USER").Value)%>">
              </font></td>
            <td width="176" valign="top"><button onclick="this.innerHTML='Please wait. Processing...'; this.disabled=true; document.forms[0].submit();" type="button" class="submit" name="Submit">Take 
              The News to the People!</button></td>
            <td width="158" valign="top"><button onclick="self.location='admin.asp'" type="button" class="submit" name="Submit">Cancel</button></td>
          </tr>
        </table>
      </form></td>
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