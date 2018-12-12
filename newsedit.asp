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
set rsEditSelect = Server.CreateObject("ADODB.Recordset")
rsEditSelect.ActiveConnection = MM_NewsDataBase_STRING
rsEditSelect.Source = "SELECT *  FROM tblnews  ORDER BY ID ASC"
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
    <td width="95%" class="HeaderRow" >News Administration</td>
  </tr>
  <tr> 
    <td width="95%" height="64"> <p>Please Select the news article you wish to 
        edit below</p>
      <form name="EditPost1" method="post" action="newsdoedit<% 			if Request("p") <> "" then			response.write "html"			end if			%>.asp"><table width="100%" border="0">
        <!--DWLayoutTable-->
        <tr> 
          <td colspan="2" class="SubHeaderRow" >Edit News</td>
        </tr>
        <tr> 
          <td colspan="2"> 
              <select name="EDITCODE" class="inquiryform">
                <%
While (NOT rsEditSelect.EOF)
%>
                <option value="<%=(rsEditSelect.Fields.Item("ID").Value)%>" selected ><%=(rsEditSelect.Fields.Item("SUBJECT").Value)%> 
                posted by <%=(rsEditSelect.Fields.Item("USERCREATED").Value)%> 
                on <%=(rsEditSelect.Fields.Item("CREATED").Value)%></option>
                <%
  rsEditSelect.MoveNext()
Wend
If (rsEditSelect.CursorType > 0) Then
  rsEditSelect.MoveFirst
Else
  rsEditSelect.Requery
End If
%>
              </select>
              <br>
              
            </td>
        </tr>
        <tr> 
          <td width="122" height="20" valign="top">                    <button onclick="this.innerHTML='Please wait. Processing...'; this.disabled=true; document.forms[0].submit();" type="button" class="submit" name="Submit">Edit Article</button>
</td>
          <td width="587" valign="top"><button onclick="self.location='admin.asp'" type="button" class="submit" name="Submit">Cancel</button></td>
        </tr>
      </table></form></td>
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
rsEditSelect.Close()
%>