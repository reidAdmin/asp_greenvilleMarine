<%	If Request.QueryString("action") = "add" Then
	Set objConn = Server.CreateObject ("ADODB.Connection")
	objConn.Provider="Microsoft.Jet.Oledb.4.0" 'For Access Database
    objConn.Open "D:\Hosts\greenvillemarine.com\www\fpdb\news2.mdb"

	Set objRec = server.createobject ("adodb.recordset")
	Const adOpenKeyset = 1
	Const adLockOptimistic = 3

	objRec.CursorType = adOpenKeyset
	objRec.LockType = adLockOptimistic

	objRec.Open "SELECT * FROM CATEGORY", objConn
	objRec.AddNew
	For Each x In objRec.Fields
	Response.Write "* " & x.Name & " = " & Request.Form(x.Name) & " *<br>"
		Select Case x.name
			Case "id"
				tmpVar = objRec(x.Name)
				'Skip
		    Case Else
				response.write x.Name
				objRec(x.Name) = Request.Form(x.Name)
		End Select
	Next
	objRec.Update
	objRec.Close

	objRec.Open "SELECT * FROM tbluser WHERE RIGHTS = 1", objConn
	WHILE NOT objRec.EOF
		tmpRights = objRec("Category")
		objRec("Category") = tmpRights & ", " & tmpVar
		objRec.MoveNext
	Wend
	objrec.Close
	Set objRec = Nothing 

    objConn.Close
    Set objConn = Nothing
	Response.Redirect "admin.asp"
Else
%>
<html>
<head>
<title>Add Category</title>
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
    <td width="95%" height="11" class="HeaderRow" >Add Category</td>
  </tr>
  <tr> 
   <FORM METHOD="POST" ACTION="cate.asp?action=add">
        <table width="20%" border="0">
          <tr> 
            <td width="17%" class="SubHeaderRow"> <div align="center">Category</div></td>
            <td width="20"> 
            <input name="CATEGORY" type="text" class="inquiryform" value="" size="20"> 
            </td>
          </tr>
         <td width="50%"> <input name="Submit" type="submit" class="submit" align=center value="Add Me!"> 
            </td>
      </form></td>
  </tr>
</table>

  
</center></body>
</html>
<%
End If
%>