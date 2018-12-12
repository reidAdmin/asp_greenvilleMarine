<%@LANGUAGE="VBSCRIPT"%>
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
' *** Edit Operations: declare variables

MM_editAction = CStr(Request("URL"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Request.QueryString
End If

' boolean to abort record edit
MM_abortEdit = false

' query string to execute
MM_editQuery = ""
%>
<%
' *** Delete Record: declare variables

if (CStr(Request("MM_delete")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_NewsDataBase_STRING
  MM_editTable = "tbluser"
  MM_editColumn = "ID"
  MM_recordId = Request.Form("MM_recordId")
  MM_editRedirectUrl = "admin.asp"

  ' append the query string to the redirect URL
  If (MM_editRedirectUrl <> "" And Request.QueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
    Else
      MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
    End If
  End If
  
End If
%>
<%
' *** Delete Record: construct a sql delete statement and execute it

If (CStr(Request("MM_delete")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  ' create the sql delete statement
  MM_editQuery = "delete from " & MM_editTable & " where " & MM_editColumn & " = " & MM_recordId

  If (Not MM_abortEdit) Then
    ' execute the delete
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If
%>
<%
set rsUser = Server.CreateObject("ADODB.Recordset")
rsUser.ActiveConnection = MM_NewsDataBase_STRING
rsUser.Source = "SELECT * FROM tbluser"
rsUser.CursorType = 0
rsUser.CursorLocation = 2
rsUser.LockType = 3
rsUser.Open()
rsUser_numRows = 0
%>
<%
Dim rsDeleteSelect__VarID
rsDeleteSelect__VarID = "1"
if (Request.Form("DELCODE") <> "") then rsDeleteSelect__VarID = Request.Form("DELCODE")
%>
<%
set rsDeleteSelect = Server.CreateObject("ADODB.Recordset")
rsDeleteSelect.ActiveConnection = MM_NewsDataBase_STRING
rsDeleteSelect.Source = "SELECT *  FROM tbluser  WHERE ID = " + Replace(rsDeleteSelect__VarID, "'", "''") + "  ORDER BY ID ASC"
rsDeleteSelect.CursorType = 0
rsDeleteSelect.CursorLocation = 2
rsDeleteSelect.LockType = 3
rsDeleteSelect.Open()
rsDeleteSelect_numRows = 0
%>
<html>
<head>
<title>News Administration</title>
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
    <td width="95%" class="HeaderRow">News Administration</td>
  </tr>
  <tr> 
    <td width="95%" height="64"> <form ACTION="<%=MM_editAction%>" METHOD="POST" name ="delete">
        <table width="100%" border="0">
          
        <!--DWLayoutTable-->
        <tr> 
          <td height="16" colspan="2" class="SubHeaderRow">Confirm Deletion</td>
        </tr>
        <tr> 
          <td height="31" colspan="2" valign="top">Are you sure you wish to delete 
            the user: <%=(rsDeleteSelect.Fields.Item("USER").Value)%> ?</td>
        </tr>
        <tr> 
          <td width="356" height="22" valign="top"><input name="Submit" type="submit" class="submit" value="Yes I am Sure!">
            </td>
          <td width="353" valign="top"><button onclick="self.location='admin.asp'" type="button" class="submit" name="Submit">Cancel</button></td>
        </tr>
      </table>
        <input type="hidden" name="MM_delete" value="delete">
        <input type="hidden" name="MM_recordId" value="<%= rsUser.Fields.Item("ID").Value %>">
      </form></td>
  </tr>
</table>
<p align="center">
</p>
  
</center></body>
</html>
<%
rsUser.Close()
%>
<%
rsDeleteSelect.Close()
%>