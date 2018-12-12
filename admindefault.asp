<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/NewsDataBase.asp" -->
<% 
if Request.form("remember") = "1" then
Response.cookies("VarUsername") = Request.form("username")
end if

' *** Validate request to log in to this site.
MM_LoginAction = Request.ServerVariables("URL")
If Request.QueryString<>"" Then MM_LoginAction = MM_LoginAction + "?" + Request.QueryString
MM_valUsername=CStr(Request.Form("username"))
If MM_valUsername <> "" Then
  MM_fldUserAuthorization=""
  MM_redirectLoginSuccess="admin.asp"
  MM_redirectLoginFailed="default.asp?p=badpassword"
  MM_flag="ADODB.Recordset"
  set MM_rsUser = Server.CreateObject(MM_flag)
  MM_rsUser.ActiveConnection = MM_NewsDataBase_STRING
  MM_rsUser.Source = "SELECT USER, PASS"
  If MM_fldUserAuthorization <> "" Then MM_rsUser.Source = MM_rsUser.Source & "," & MM_fldUserAuthorization
  MM_rsUser.Source = MM_rsUser.Source & " FROM tbluser WHERE USER='" & MM_valUsername &"' AND PASS='" & CStr(Request.Form("password")) & "'"
  MM_rsUser.CursorType = 0
  MM_rsUser.CursorLocation = 2
  MM_rsUser.LockType = 3
  MM_rsUser.Open
  If Not MM_rsUser.EOF Or Not MM_rsUser.BOF Then 
    ' username and password match - this is a valid user
    Session("MM_Username") = MM_valUsername
    If (MM_fldUserAuthorization <> "") Then
      Session("MM_UserAuthorization") = CStr(MM_rsUser.Fields.Item(MM_fldUserAuthorization).Value)
    Else
      Session("MM_UserAuthorization") = ""
    End If
    if CStr(Request.QueryString("accessdenied")) <> "" And false Then
      MM_redirectLoginSuccess = Request.QueryString("accessdenied")
    End If
    MM_rsUser.Close
    Response.Redirect(MM_redirectLoginSuccess)
  End If
  MM_rsUser.Close
  Response.Redirect(MM_redirectLoginFailed)
End If
%>
<html>
<head>
<title>Login</title>
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
    <td width="100%" class="HeaderRow">News Admin</td>
  </tr>
  <tr> 
    <td width="93%"> <% 
	if request("p") = "badpassword" then
	response.write "Sorry your username or password are incorrect"
	end if
	%> <form name="LogMeIn" method="post" action="<%=MM_LoginAction%>">
        <table width="100%" border="0">
          <tr> 
            <td width="127" class="SubHeaderRow">Username:</td>
            <td> 
            <input name="username" type="text" class="inquiryform" value="<%=(Request.cookies("VarUsername"))%>" size="20"> 
            </td>
          </tr>
          <tr> 
            <td width="127" class="SubHeaderRow">Password:</td>
            <td> 
            <input name="password" type="password" class="inquiryform" size="20"> </font></b>
            </td>
          </tr>
          <tr> 
            <td width="127" class="SubHeaderRow">Remember Login?</td>
            <td> <input type="checkbox" name="remember" value="1">
              <br> </td>
          </tr>
          <tr> 
            <td width="127" class="SubHeaderRow"></td>
            <td>
<button onclick="this.innerHTML='Please wait. Logging In...'; this.disabled=true; document.forms[0].submit();" type="button" class="submit" name="Submit">Log Me In!</button></td>
          </tr>
        </table>
      </form></td>
  </tr>
</table>
<br>
&nbsp;<p align="center">
</p>
  
</center>
</body>
</html>