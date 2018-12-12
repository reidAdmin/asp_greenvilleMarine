<%@LANGUAGE="VBSCRIPT"%>
<%
' *** Logout the current user.
MM_Logout = CStr(Request.ServerVariables("URL")) & "?MM_Logoutnow=1"
If (CStr(Request("MM_Logoutnow")) = "1") Then
  Session.Abandon
  MM_logoutRedirectPage = "default.asp"
  ' redirect with URL parameters (remove the "MM_Logoutnow" query param).
  if (MM_logoutRedirectPage = "") Then MM_logoutRedirectPage = CStr(Request.ServerVariables("URL"))
  If (InStr(1, UC_redirectPage, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
    MM_newQS = "?"
    For Each Item In Request.QueryString
      If (Item <> "MM_Logoutnow") Then
        If (Len(MM_newQS) > 1) Then MM_newQS = MM_newQS & "&"
        MM_newQS = MM_newQS & Item & "=" & Server.URLencode(Request.QueryString(Item))
      End If
    Next
    if (Len(MM_newQS) > 1) Then MM_logoutRedirectPage = MM_logoutRedirectPage & MM_newQS
  End If
  Response.Redirect(MM_logoutRedirectPage)
End If
%>
<!--#include file="Connections/NewsDataBase.asp" -->
<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers=""
MM_authFailedURL="default.asp?p=badlogin"
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
set rsUser = Server.CreateObject("ADODB.Recordset")
rsUser.ActiveConnection = MM_NewsDataBase_STRING
rsUser.Source = "SELECT *  FROM tbluser"
rsUser.CursorType = 0
rsUser.CursorLocation = 2
rsUser.LockType = 3
rsUser.Open()
rsUser_numRows = 0
%>
<%
Dim rsUserLevelAdmin__VarUser
rsUserLevelAdmin__VarUser = "1"
if (Session("MM_username") <> "") then rsUserLevelAdmin__VarUser = Session("MM_username")
%>
<%
set rsUserLevelAdmin = Server.CreateObject("ADODB.Recordset")
rsUserLevelAdmin.ActiveConnection = MM_NewsDataBase_STRING
rsUserLevelAdmin.Source = "SELECT *  FROM tbluser  WHERE USER = '" + Replace(rsUserLevelAdmin__VarUser, "'", "''") + "' AND RIGHTS = 1"
rsUserLevelAdmin.CursorType = 0
rsUserLevelAdmin.CursorLocation = 2
rsUserLevelAdmin.LockType = 3
rsUserLevelAdmin.Open()
rsUserLevelAdmin_numRows = 0
%>
<html>
<head>
<title>News Administration</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="css_styles/site.css" rel="stylesheet" type="text/css">
</head>
<body>
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
    <td width="95%" class="HeaderRow">News Administration</td>
  </tr>
  <tr> 
    <td width="95%" height="217"> <p>Welcome to the administration area for adding 
        and deleting news articles</p>
      <table width="100%" border="0">
        <tr> 
          <td class="SubHeaderRow">News Admin</td>
        </tr>
        <tr> 
          <td>- <a href="newsaddhtml.asp">Add News</a><br>
             - <a href="newsadd.asp">Add News - Basic</a><br>
            - <a href="newsedit.asp?p=1">Edit News</a><br>
           
            - <a href="newsdelete.asp">Delete News</a></td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
        </tr>
        <% If Not rsUserLevelAdmin.EOF Or Not rsUserLevelAdmin.BOF Then %>
        <tr> 
          <td bgcolor="#6699FF" class="SubHeaderRow">User 
            Admin</td>
        </tr>
        <tr> 
          <td>- 
            <a href="useradd.asp">Add Users</a><br>
            - <a href="useredit.asp">Change Users</a><br>
			- <a href="cate.asp">Add Category</a><br>
      
            - <a href="userdelete.asp">Delete Users</a></td>
        </tr>
        <% End If ' end Not rsUserLevelAdmin.EOF Or NOT rsUserLevelAdmin.BOF %>
        <tr> 
          <td><br>
            - <a href="userprefs.asp">Edit Prefences</a><br>
            <br>
            - <A HREF="<%=MM_Logout%>">Logout</A></td>
        </tr>
      </table></td>
  </tr>
</table>
<p><br>
  &nbsp;</p>
</body>
</html>
<%
rsUser.Close()
%>
<%
rsUserLevelAdmin.Close()
%>