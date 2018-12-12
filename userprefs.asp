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
  MM_editTable = "tbluser"
  MM_editColumn = "ID"
  MM_recordId = "'" + Request.Form("MM_recordId") + "'"
  MM_editRedirectUrl = "admin.asp"
  MM_fieldsStr  = "PASS|value|NAME|value|EMAIL|value"
  MM_columnsStr = "PASS|',none,''|NAME|',none,''|EMAIL|',none,''"

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
Dim rsEditSelect__VarID
rsEditSelect__VarID = "1"
If (Session("MM_username") <> "") Then 
  rsEditSelect__VarID = Session("MM_username")
End If
%>
<%
set rsEditSelect = Server.CreateObject("ADODB.Recordset")
rsEditSelect.ActiveConnection = MM_NewsDataBase_STRING
rsEditSelect.Source = "SELECT *  FROM tbluser  WHERE USER = '" + Replace(rsEditSelect__VarID, "'", "''") + "'  ORDER BY ID ASC"
rsEditSelect.CursorType = 0
rsEditSelect.CursorLocation = 2
rsEditSelect.LockType = 3
rsEditSelect.Open()
rsEditSelect_numRows = 0
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
    <td width="95%" height="64"> <table width="100%" border="0">
        <tr> 
          <td class="SubHeaderRow">Edit Prefences</td>
        </tr>
        <tr> 
          <td> <form ACTION="<%=MM_editAction%>" METHOD="POST" name="EditUser">
              <table width="100%" border="0">
                <!--DWLayoutTable-->
                <tr> 
                  <td width="83" class="SubHeaderRow"> <div align="right"><i><b>username:</b></i></div></td>
                  <td colspan="2"> <%=(rsEditSelect.Fields.Item("USER").Value)%> </td>
                </tr>
                <tr> 
                  <td width="83" class="SubHeaderRow"> <div align="right"><i><b>password:</b></i></div></td>
                  <td colspan="2"> 
                  <input name="PASS" type="password" class="inquiryform" value="<%=(rsEditSelect.Fields.Item("PASS").Value)%>" size="20"> 
                  </td>
                </tr>
                <tr> 
                  <td width="83" class="SubHeaderRow"> <div align="right"><i><b>Date 
                      Created:</b></i></div></td>
                  <td colspan="2"><%=(rsEditSelect.Fields.Item("DATECREATED").Value)%></td>
                </tr>
                <tr> 
                  <td width="83" class="SubHeaderRow"> <div align="right"><i><b>name:</b></i></div></td>
                  <td colspan="2"> 
                  <input name="NAME" type="text" class="inquiryform" value="<%=(rsEditSelect.Fields.Item("NAME").Value)%>" size="20"> 
                  </td>
                </tr>
                <tr> 
                  <td width="83" class="SubHeaderRow"> <div align="right"><i><b>email:</b></i></div></td>
                  <td colspan="2"> 
                  <input name="EMAIL" type="text" class="inquiryform" value="<%=(rsEditSelect.Fields.Item("EMAIL").Value)%>" size="20"> 
                  </td>
                </tr>
                <tr> 
                  <td width="83" height="16" class="SubHeaderRow"> <div align="right"></div></td>
                  <td width="202" valign="top"> <button onclick="this.innerHTML='Please wait. Updating...'; this.disabled=true; document.forms[0].submit();" type="button" class="submit" name="Submit">Change 
                    My Prefences</button></td>
                  <td width="397" valign="top"><button onclick="self.location='admin.asp'" type="button" class="submit" name="Submit">Cancel</button></td>
                </tr>
              </table>
              <input type="hidden" name="MM_update" value="EditUser">
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
rsEditSelect.Close()
%>