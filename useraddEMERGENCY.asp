<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/NewsDataBase.asp" -->
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
' *** Insert Record: set variables

If (CStr(Request("MM_insert")) <> "") Then

  MM_editConnection = MM_NewsDataBase_STRING
  MM_editTable = "tbluser"
  MM_editRedirectUrl = "admin.asp"
  MM_fieldsStr  = "USER|value|PASS|value|Name|value|EMAIL|value|CATEGORY|value|level|value"
  MM_columnsStr = "USER|',none,''|PASS|',none,''|NAME|',none,''|EMAIL|',none,''|CATEGORY|',none,''|LEVEL|',none,''"

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
' *** Insert Record: construct a sql insert statement and execute it

If (CStr(Request("MM_insert")) <> "") Then

  ' create the sql insert statement
  MM_tableValues = ""
  MM_dbValues = ""
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
      MM_tableValues = MM_tableValues & ","
      MM_dbValues = MM_dbValues & ","
    End if
    MM_tableValues = MM_tableValues & MM_columns(i)
    MM_dbValues = MM_dbValues & FormVal
  Next
  MM_editQuery = "insert into " & MM_editTable & " (" & MM_tableValues & ") values (" & MM_dbValues & ")"

  If (Not MM_abortEdit) Then
    ' execute the insert
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
set rsNewUser = Server.CreateObject("ADODB.Recordset")
rsNewUser.ActiveConnection = MM_NewsDataBase_STRING
rsNewUser.Source = "SELECT *  FROM tbluser"
rsNewUser.CursorType = 0
rsNewUser.CursorLocation = 2
rsNewUser.LockType = 3
rsNewUser.Open()
rsNewUser_numRows = 0
%>
<html>
<head>
<title>Add User</title>
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
    <td width="95%" height="11" class="HeaderRow">Add User - Emergency</td>
  </tr>
  <tr> 
    <td width="95%"> <form name="AddUser" method="POST" action="<%=MM_editAction%>">
        <table width="54%" border="0">
          <tr> 
            <td width="17%" class="SubHeaderRow"> <div align="right">Username</div></td>
            <td width="83%"> 
            <input name="USER" type="text" class="inquiryform" value="" size="20"> 
            </td>
          </tr>
          <tr> 
            <td width="17%" height="49" class="SubHeaderRow"> <div align="right">Password</div></td>
            <td width="83%" height="49"> 
            <input name="PASS" type="text" class="inquiryform" size="20"> 
            </td>
          </tr>
          <tr> 
            <td width="17%" height="2" class="SubHeaderRow"> <div align="right">Name</div></td>
            <td width="83%" height="2"> 
            <input name="Name" type="text" class="inquiryform" size="20"> 
            </td>
          </tr>
          <tr> 
            <td width="17%" class="SubHeaderRow"> <div align="right">Email</div></td>
            <td width="83%"> 
            <input name="EMAIL" type="text" class="inquiryform" size="20"> 
            </td>
          </tr>
		  <tr> 
            <td width="17%" class="SubHeaderRow"> <div align="right">Category</div></td>
            <td width="83%"> 
            <input name="CATEGORY" type="text" class="inquiryform" size="20"> 
            </td>
          </tr>
          <tr> 
            <td width="17%" class="SubHeaderRow"> <div align="right">Level</div></td>
            <td width="83%"> <select name="level" class="inquiryform">
                <option value="2">Editor</option>
                <option value="1">Admin</option>
              </select> </td>
          </tr>
          <tr> 
            <td width="17%"> <div align="right"><font color="#FFFFFF"><i><b><font color="#FFFFFF"></font></b></i></font> 
              </div></td>
            <td width="83%"> <input name="Submit" type="submit" class="submit" value="Add Me!"> 
            </td>
          </tr>
        </table>
        <input type="hidden" name="MM_insert" value="true">
      </form></td>
  </tr>
</table>
</body>
</html>
<%
rsNewUser.Close()
%>