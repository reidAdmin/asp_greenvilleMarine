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
' *** Update Record: set variables

If (CStr(Request("MM_update")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_NewsDataBase_STRING
  MM_editTable = "tblnews"
  MM_editColumn = "ID"
  MM_recordId = Request.Form("MM_recordId")
  MM_editRedirectUrl = "admin.asp"
  MM_fieldsStr  = "subject|value|EditorValue|value|CATEGORY|value"
  MM_columnsStr = "SUBJECT|',none,''|BODY|',none,''|CATEGORY|',none,''"

  ' create the MM_fields and MM_columns arrays
  MM_fields = Split(MM_fieldsStr, "|")
  MM_columns = Split(MM_columnsStr, "|")
  
  ' set the form values
  For i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_fields(i+1) = CStr(Request.Form(MM_fields(i)))
  Next

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
' *** Update Record: construct a sql update statement and execute it

If (CStr(Request("MM_update")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  ' create the sql update statement
  MM_editQuery = "update " & MM_editTable & " set "
  For i = LBound(MM_fields) To UBound(MM_fields) Step 2
    FormVal = MM_fields(i+1)
    MM_typeArray = Split(MM_columns(i+1),",")
    Delim = MM_typeArray(0)
    If (Delim = "none") Then Delim = ""
    AltVal = MM_typeArray(1)
    If (AltVal = "none") Then AltVal = ""
    EmptyVal = MM_typeArray(2)
    If (EmptyVal = "none") Then EmptyVal = ""
    If (FormVal = "") Then
      FormVal = EmptyVal
    Else
      If (AltVal <> "") Then
        FormVal = AltVal
      ElseIf (Delim = "'") Then  ' escape quotes
        FormVal = "'" & Replace(FormVal,"'","''") & "'"
      Else
        FormVal = Delim + FormVal + Delim
      End If
    End If
    If (i <> LBound(MM_fields)) Then
      MM_editQuery = MM_editQuery & ","
    End If
    MM_editQuery = MM_editQuery & MM_columns(i) & " = " & FormVal
  Next
  MM_editQuery = MM_editQuery & " where " & MM_editColumn & " = " & MM_recordId

  If (Not MM_abortEdit) Then
    ' execute the update
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
Dim rsEditNews__VarCode
rsEditNews__VarCode = "1"
if (Request.Form("EDITCODE")  <> "") then rsEditNews__VarCode = Request.Form("EDITCODE") 
%>
<%
set rsEditNews = Server.CreateObject("ADODB.Recordset")
rsEditNews.ActiveConnection = MM_NewsDataBase_STRING
rsEditNews.Source = "SELECT *  FROM tblnews  WHERE ID = " + Replace(rsEditNews__VarCode, "'", "''")
rsEditNews.CursorType = 0
rsEditNews.CursorLocation = 2
rsEditNews.LockType = 3
rsEditNews.Open()
rsEditNews_numRows = 0
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
<title>Edit News Article</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="css_styles/site.css" rel="stylesheet" type="text/css">
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
    <td width="95%" class="HeaderRow">Edit News Article</td>
  </tr>
  <tr> 
    <form name="fHtmlEditor" method="POST" action="<%=MM_editAction%>">
      <td width="95%"> <table width="54%" border="0">
          <tr> 
            <td width="40%" class="SubHeaderRow">Subject</td>
            <td width="30%"> 
            <input name="subject" type="text" class="inquiryform" value="<%=(rsEditNews.Fields.Item("SUBJECT").Value)%>" size="20"> 
              <input type="hidden" name="user" value="<%=(Session(MM_username))%>"> 
                  <tr> 
            <td width="40%" class="SubHeaderRow">Related Links</td>
            <td width="30%"> 
            <input name="RelatedLink" type="text" class="inquiryform" value="<%=(rsEditNews.Fields.Item("RELATEDLINK").Value)%>" size="20"> 
            </td> 
			<td>
			<tr>
            <td align="left" nowrap class="SubHeaderRow">Category:</td>
			<td width="83%">
			<%
			dim x
			x = rsUser("id")
			%>
			<select name="CATEGORY">
			<%While Not rsUser2.EOF

				If InStr(rsUser("CATEGORY"),rsUser2("id")) Then
					Response.Write "<option value=""" & rsUser2("id") & """"
					If Int(rsEditNews("CATEGORY")) = rsUser2("id") Then
						Response.Write " selected"
					End If
					Response.Write ">" & rsUser2("CATEGORY")
		
				End If
				rsUser2.MoveNExt

			Wend
			%>
			

			</td>
            <td width="30%"><input name="Submit" type="submit" class="submit" value="Save"></td>
          </tr>
        </table>
        <style type="text/css">
<!--
.clsCursor {  cursor: hand}
-->
      </style> <textarea name="EditorValue" cols="80" rows="26"><%=(rsEditNews.Fields.Item("BODY").Value)%></textarea>
<script language="javascript1.2">
editor_generate('EditorValue','',''); // field, width, height
      </script> 
        
      </td>
      <input type="hidden" name="MM_update" value="true">
      <input type="hidden" name="MM_recordId" value="<%= rsEditNews.Fields.Item("ID").Value %>">
    </form>
  </tr>
</table>
<p align="center">
</p>
  
</center></body>
</html>
<%
rsEditNews.Close()
%>
<%
rsUser.Close()
%>