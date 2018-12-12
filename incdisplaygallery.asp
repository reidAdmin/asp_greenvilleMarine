<!--#include file="administration/inc/db_conn_open.asp"-->
<%
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open sDSN

Set objRec = Server.CreateObject("ADODB.RecordSet")
%>
<script language="JavaScript">
function UpdateGallery(gid){
	if (gid != 0){
		location.href = '<%= Request.ServerVariables("PATH_INFO") %>?id='+gid
	}
}

function EnlargePic(pid){
	window.open('viewpic.asp?id='+pid,'picwin','width=800,height=600,top=10,left=10,location=no,menubar=no,resizable=yes,scrollbars=yes');
}
</script>
<table width="96%"% border="0" cellpadding="0" cellspacing="0">
  <tr>
	<td><select name="galleryId" style="font-size:11px" onChange="JavaScript:UpdateGallery(this.value)">
	  <%
	  If Request.QueryString("id") <> "" Then
		galleryId = CInt(Request.QueryString("id"))
	  Else
		galleryId = 0							
	  End If

	  objRec.Open "SELECT * FROM galleries ORDER BY galleryName", objConn
	  If NOT objRec.EOF Then
		While NOT objRec.EOF
			%>
			<option value="<%= objRec("galleryId") %>"<% If objRec("galleryId") = galleryId Then %> selected<% End If %>><%= objRec("galleryName") %>
			<%			
			If galleryId = 0 Then
				galleryId = objRec("galleryId")
			End If
			objRec.MoveNext
		Wend
	  End If
	  objRec.Close
	  %></select></td></tr>						  
  <tr>
	<td align="center">
	  <table cellpadding="0" cellspacing="0" border="0" width="98%">
		<tr>
		  <td rowspan="1000" width="10"></td>
		  <td colspan="4"><hr></td></tr>
		<%
		intPicCount = 1
		objRec.Open "SELECT * FROM pics WHERE picGalleryId = " & galleryId, objConn
		If NOT objRec.EOF Then
			While NOT objRec.EOF
				If intPicCount = 1 Then
					Response.Write "<tr>"
				End If

				Response.Write "<td style=""text-align:center;vertical-align:bottom"" width=""33%""><a href=""JavaScript:EnlargePic('" & objRec("picId") & "')""><img src=""thumbnail.asp?filename=" & objRec("picFile") & """ border=""0""></a><br><a href=""JavaScript:EnlargePic('" & objRec("picId") & "')""><b>" & objRec("picName") & "</b></a></td>"

				intPicCount = intPicCount + 1


				If intPicCount = 4 Then
					Response.Write "</tr>" & vbCrLf
					Response.Write "<tr><td colspan=""3""><br></td></tr>"
					intPicCount = 1
				End If
				objRec.MoveNext
			Wend
		Else
			Response.Write "<tr><td>No Pictures Available</td></tr>"
		End If
		objRec.Close

		If intPicCount = 2 Then
			Response.Write "<td width=""33%"">&nbsp;</td><td width=""33%"">&nbsp;</td>"
		Else
			If intPicCount = 3 Then
				Response.Write "<td width=""33%"">&nbsp;</td>"
			End If
		End If
		%>
	  </table></td></tr>
  <tr>
	<td><br><br><br><br></td></tr>
</table>
<%
Set objRec = NOTHING

objConn.Close
Set objConn = NOTHING
%>