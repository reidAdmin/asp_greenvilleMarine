<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/NewsDataBase.asp" -->
<%

Dim VarPageNo
if request("offset") <> "" then
if request("offset") <> "0" then
VarPageNo = "You are on page no. " & request("offset")
else
VarPageNo = ""
end if
else
VarPageNo = ""
end if

'*************************** CHANGE CATEGORY ID HERE **************************
CategoryID = 1
'******************************************************************************

set rsNews = Server.CreateObject("ADODB.Recordset")
rsNews.ActiveConnection = MM_NewsDataBase_STRING
rsNews.Source = "SELECT * FROM tblnews WHERE Category = '" & CategoryID & "' ORDER BY ID DESC"
rsNews.CursorType = 0
rsNews.CursorLocation = 2
rsNews.LockType = 3
rsNews.Open()
rsNews_numRows = 0
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = 10
Dim Repeat1__index
Repeat1__index = 0
rsNews_numRows = rsNews_numRows + Repeat1__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

' set the record count
rsNews_total = rsNews.RecordCount

' set the number of rows displayed on this page
If (rsNews_numRows < 0) Then
  rsNews_numRows = rsNews_total
Elseif (rsNews_numRows = 0) Then
  rsNews_numRows = 1
End If

' set the first and last displayed record
rsNews_first = 1
rsNews_last  = rsNews_first + rsNews_numRows - 1

' if we have the correct record count, check the other stats
If (rsNews_total <> -1) Then
  If (rsNews_first > rsNews_total) Then rsNews_first = rsNews_total
  If (rsNews_last > rsNews_total) Then rsNews_last = rsNews_total
  If (rsNews_numRows > rsNews_total) Then rsNews_numRows = rsNews_total
End If
%>
<%
' *** Move To Record and Go To Record: declare variables

Set MM_rs    = rsNews
MM_rsCount   = rsNews_total
MM_size      = rsNews_numRows
MM_uniqueCol = ""
MM_paramName = ""
MM_offset = 0
MM_atTotal = false
MM_paramIsDefined = false
If (MM_paramName <> "") Then
  MM_paramIsDefined = (Request.QueryString(MM_paramName) <> "")
End If
%>
<%
' *** Move To Record: handle 'index' or 'offset' parameter

if (Not MM_paramIsDefined And MM_rsCount <> 0) then

  ' use index parameter if defined, otherwise use offset parameter
  r = Request.QueryString("index")
  If r = "" Then r = Request.QueryString("offset")
  If r <> "" Then MM_offset = Int(r)

  ' if we have a record count, check if we are past the end of the recordset
  If (MM_rsCount <> -1) Then
    If (MM_offset >= MM_rsCount Or MM_offset = -1) Then  ' past end or move last
      If ((MM_rsCount Mod MM_size) > 0) Then         ' last page not a full repeat region
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' move the cursor to the selected record
  i = 0
  While ((Not MM_rs.EOF) And (i < MM_offset Or MM_offset = -1))
    MM_rs.MoveNext
    i = i + 1
  Wend
  If (MM_rs.EOF) Then MM_offset = i  ' set MM_offset to the last possible record

End If
%>
<%
' *** Move To Record: if we dont know the record count, check the display range

If (MM_rsCount = -1) Then

  ' walk to the end of the display range for this page
  i = MM_offset
  While (Not MM_rs.EOF And (MM_size < 0 Or i < MM_offset + MM_size))
    MM_rs.MoveNext
    i = i + 1
  Wend

  ' if we walked off the end of the recordset, set MM_rsCount and MM_size
  If (MM_rs.EOF) Then
    MM_rsCount = i
    If (MM_size < 0 Or MM_size > MM_rsCount) Then MM_size = MM_rsCount
  End If

  ' if we walked off the end, set the offset based on page size
  If (MM_rs.EOF And Not MM_paramIsDefined) Then
    If (MM_offset > MM_rsCount - MM_size Or MM_offset = -1) Then
      If ((MM_rsCount Mod MM_size) > 0) Then
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' reset the cursor to the beginning
  If (MM_rs.CursorType > 0) Then
    MM_rs.MoveFirst
  Else
    MM_rs.Requery
  End If

  ' move the cursor to the selected record
  i = 0
  While (Not MM_rs.EOF And i < MM_offset)
    MM_rs.MoveNext
    i = i + 1
  Wend
End If
%>
<%
' *** Move To Record: update recordset stats

' set the first and last displayed record
rsNews_first = MM_offset + 1
rsNews_last  = MM_offset + MM_size
If (MM_rsCount <> -1) Then
  If (rsNews_first > MM_rsCount) Then rsNews_first = MM_rsCount
  If (rsNews_last > MM_rsCount) Then rsNews_last = MM_rsCount
End If

' set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount <> -1 And MM_offset + MM_size >= MM_rsCount)
%>
<%
' *** Go To Record and Move To Record: create strings for maintaining URL and Form parameters

' create the list of parameters which should not be maintained
MM_removeList = "&index="
If (MM_paramName <> "") Then MM_removeList = MM_removeList & "&" & MM_paramName & "="
MM_keepURL="":MM_keepForm="":MM_keepBoth="":MM_keepNone=""

' add the URL parameters to the MM_keepURL string
For Each Item In Request.QueryString
  NextItem = "&" & Item & "="
  If (InStr(1,MM_removeList,NextItem,1) = 0) Then
    MM_keepURL = MM_keepURL & NextItem & Server.URLencode(Request.QueryString(Item))
  End If
Next

' add the Form variables to the MM_keepForm string
For Each Item In Request.Form
  NextItem = "&" & Item & "="
  If (InStr(1,MM_removeList,NextItem,1) = 0) Then
    MM_keepForm = MM_keepForm & NextItem & Server.URLencode(Request.Form(Item))
  End If
Next

' create the Form + URL string and remove the intial '&' from each of the strings
MM_keepBoth = MM_keepURL & MM_keepForm
if (MM_keepBoth <> "") Then MM_keepBoth = Right(MM_keepBoth, Len(MM_keepBoth) - 1)
if (MM_keepURL <> "")  Then MM_keepURL  = Right(MM_keepURL, Len(MM_keepURL) - 1)
if (MM_keepForm <> "") Then MM_keepForm = Right(MM_keepForm, Len(MM_keepForm) - 1)

' a utility function used for adding additional parameters to these strings
Function MM_joinChar(firstItem)
  If (firstItem <> "") Then
    MM_joinChar = "&"
  Else
    MM_joinChar = ""
  End If
End Function
%>
<%
' *** Move To Record: set the strings for the first, last, next, and previous links

MM_keepMove = MM_keepBoth
MM_moveParam = "index"

' if the page has a repeated region, remove 'offset' from the maintained parameters
If (MM_size > 0) Then
  MM_moveParam = "offset"
  If (MM_keepMove <> "") Then
    params = Split(MM_keepMove, "&")
    MM_keepMove = ""
    For i = 0 To UBound(params)
      nextItem = Left(params(i), InStr(params(i),"=") - 1)
      If (StrComp(nextItem,MM_moveParam,1) <> 0) Then
        MM_keepMove = MM_keepMove & "&" & params(i)
      End If
    Next
    If (MM_keepMove <> "") Then
      MM_keepMove = Right(MM_keepMove, Len(MM_keepMove) - 1)
    End If
  End If
End If

' set the strings for the move to links
If (MM_keepMove <> "") Then MM_keepMove = MM_keepMove & "&"
urlStr = Request.ServerVariables("URL") & "?" & MM_keepMove & MM_moveParam & "="
MM_moveFirst = urlStr & "0"
MM_moveLast  = urlStr & "-1"
MM_moveNext  = urlStr & Cstr(MM_offset + MM_size)
prev = MM_offset - MM_size
If (prev < 0) Then prev = 0
MM_movePrev  = urlStr & Cstr(prev)
%>
<link href="css_styles/site.css" rel="stylesheet" type="text/css">



<table width="100%" border="0">
  <tr> 
    <td width="35%"> 
      <table width="100%" border="0">
        <tr> 
          <% If rsNews.EOF And rsNews.BOF Then %>
          <td class="ContentBody"><font face="Verdana, Arial, Helvetica, sans-serif">Sorry 
            but there are no news articles</font></td>
          <% End If ' end rsNews.EOF And rsNews.BOF %>
        </tr>
        <% 
While ((Repeat1__numRows <> 0) AND (NOT rsNews.EOF)) 
%>
        <tr> 
          <td>&nbsp;</td>
        </tr>
        <tr> 
          <td class="ContentHead"><%=(rsNews.Fields.Item("SUBJECT").Value)%></td>
        </tr>
        <tr> 
          <td class="ContentBody"><%=(rsNews.Fields.Item("BODY").Value)%><br> <%
			if rsNews.Fields.Item("RELATEDLINK").Value <> "" then
			%>
            Related Link: <a href="<%=(rsNews.Fields.Item("RELATEDLINK").Value)%>"><%=(rsNews.Fields.Item("RELATEDLINK").Value)%></a><br><% end if %> <br>
            Posted by <%=(rsNews.Fields.Item("USERCREATED").Value)%> at <%=(rsNews.Fields.Item("CREATEDTIME").Value)%> on <%= day((rsNews.Fields.Item("CREATED").Value)) & " " & left(monthname(month((rsNews.Fields.Item("CREATED").Value))),3) & " " & year((rsNews.Fields.Item("CREATED").Value)) %></td>
          
        </tr>
        <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsNews.MoveNext()
Wend
%>
        <tr> 
          <td> <table width="100%" border="0">
              <tr> 
                <td><font face="Verdana, Arial, Helvetica, sans-serif"> 
                  <% If MM_offset <> 0 Then %>
                  <A HREF="<%=MM_movePrev%>">&lt;&lt; 
                  Previous Page</A> 
                  <% End If ' end MM_offset <> 0 %>
                  </font></td>
                <td> <div align="right"> 
                    <% If Not MM_atTotal Then %>
                    <font face="Verdana, Arial, Helvetica, sans-serif"><A HREF="<%=MM_moveNext%>">Next 
                    Page &gt;&gt;</A></font> 
                    <% End If ' end Not MM_atTotal %>
                  </div></td>
              </tr>
            </table></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<%
rsNews.Close()
%>