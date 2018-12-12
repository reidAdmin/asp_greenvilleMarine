<!--#include file="inc/incsettings.asp"-->
<html>
<head>
<title><%= UCase(strSitename) %></title>
<meta http-equiv="content-style-type" content="text/css">
<link href="inc/css/style.css" type="text/css" rel="stylesheet">
</head>
<body bgcolor=#ffffff leftmargin=0 topmargin=0 marginwidth=0 marginheight=0>
<center>
<table cellpadding="0" cellspacing="0" border="0" bgcolor="#FFFFFF" width="700" height="100%">
  <tr>
    <td valign="top">
	  <table cellpadding="0" cellspacing="0" border="0" width="90%">
	    <tr>		  
		  <td colspan="2"><br></td></tr>
		<tr>
		  <td rowspan="3">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
		  <td><b><%= UCase(strSitename) %> PICTURE GALLERY ADMINISTRATION</b></td></tr>
		<tr>
		  <td><hr size="1" color="#000000"></td></tr>
		<tr>
		  <td>
		    <table cellpadding="0" cellspacing="0" border="0" width="100%">
			  <tr>
			    <td><br></td></tr>
			  <tr>
			    <td><!--#include file="inc/incaddgallery.asp"--></td></tr>
			</table></td></tr>
	  </table></td></tr>
</table>
</center>
</body>
</html>