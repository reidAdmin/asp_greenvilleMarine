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
' *** Update Record: set variables

If (CStr(Request("MM_update")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_NewsDataBase_STRING
  MM_editTable = "tblnews"
  MM_editColumn = "ID"
  MM_recordId = "'" + Request.Form("MM_recordId") + "'"
  MM_editRedirectUrl = "admin.asp"
  MM_fieldsStr  = "SUBJECT|value|CATEGORY|value|TEXT|value|RELATEDLINK|value"
  MM_columnsStr = "SUBJECT|',none,''|CATEGORY|',none,''|BODY|',none,''|RELATEDLINK|',none,''"

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
Dim rsEditSelect__VarID
rsEditSelect__VarID = "1"
if (Request.Form("EDITCODE") <> "") then rsEditSelect__VarID = Request.Form("EDITCODE")
%>
<%
set rsEditSelect = Server.CreateObject("ADODB.Recordset")
rsEditSelect.ActiveConnection = MM_NewsDataBase_STRING
rsEditSelect.Source = "SELECT *  FROM tblnews  WHERE ID = '" + Replace(rsEditSelect__VarID, "'", "''") + "'"
rsEditSelect.CursorType = 0
rsEditSelect.CursorLocation = 2
rsEditSelect.LockType = 3
rsEditSelect.Open()
rsEditSelect_numRows = 0
%>
<%
set rsUser2 = Server.CreateObject("ADODB.Recordset")
rsUser2.ActiveConnection = MM_NewsDataBase_STRING
rsUser2.Source = "Select * from CATEGORY"
rsUser2.CursorType = 0
rsUser2.CursorLocation = 2
rsUser2.LockType = 3
rsUser2.Open()
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
    <td width="95%" height="98"> <table width="100%" border="0">
        <tr> 
          <td height="2" class="SubHeaderRow">Edit Post</td>
        </tr>
        <tr> 
          <td> <form method="POST" action="<%=MM_editAction%>" name="form2">
              <table align="center">
                <!--DWLayoutTable-->
                <tr valign="baseline"> 
                  <td align="right" nowrap class="SubHeaderRow">SUBJECT:</td>
                  <td colspan="2"> <input name="SUBJECT" type="text" class="inquiryform" value="<%=(rsEditSelect.Fields.Item("SUBJECT").Value)%>" size="32"> 
                  </td>
			
                <tr valign="baseline"> 
                  <td align="right" valign="top" nowrap class="SubHeaderRow">TEXT:</td>
                  <td colspan="2"> <textarea name="TEXT" cols="80" rows="10" class="inquiryform"><%=(rsEditSelect.Fields.Item("BODY").Value)%></textarea> 
                  </td>
                </tr>
                <tr valign="baseline"> 
                  <td align="right" nowrap class="SubHeaderRow">RELATEDLINK:</td>
                  <td colspan="2"> <input name="RELATEDLINK" type="text" class="inquiryform" value="<%=(rsEditSelect.Fields.Item("RELATEDLINK").Value)%>" size="32"> 
                  </td>
                </tr>
				<tr valign="baseline"> 
                  <td align="right" nowrap class="SubHeaderRow">CATEGORY:</td>
				
			<td width="83%">
			<%
			dim x
			x = rsUser2("id")
			%>
			<select name="CATEGORY">
			<%While Not rsUser2.EOF %>
			<option value="<%= rsUser2("id") %>"<% If rsEditSelect("CATEGORY") = rsUser2("id") Then Response.Write " selected" End If %>> <% If InStr(rsUser("Category"),rsUser2("id")) Then response.write rsUser2("CATEGORY")  End If 
			    rsUser2.movenext
				wend
			%>


                <tr valign="baseline"> 
                  <td height="20" align="right" nowrap>&nbsp;</td>
                  <td width="205" valign="top"><button onclick="this.innerHTML='Please wait. Processing...'; this.disabled=true; document.forms[0].submit();" type="button" class="submit" name="Submit">Save 
                    Changes</button></td>
                  <td width="201" valign="top"><button onclick="self.location='admin.asp'" type="button" class="submit" name="Submit">Cancel</button></td>
                </tr>
              </table>
              <input type="hidden" name="MM_update" value="true">
              <input type="hidden" name="MM_recordId" value="<%= rsEditSelect.Fields.Item("ID").Value %>">
            </form></td>
        </tr>
      </table></td>
  </tr>
</table>
<p align="center">
</p>
  
</center></body>
</html>
<%
rsUser.Close()
rsUser2.Close()
%>
<%
rsEditSelect.Close()
%>