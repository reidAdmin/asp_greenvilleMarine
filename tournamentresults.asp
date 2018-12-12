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
CategoryID = 4
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
Repeat1__numRows = 1
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

<head>
<title>Greenville Marine - Fountain, Sea Pro, Sea Ray, May-Craft, &amp; Triton Boat 
Dealer - Tournament Results</title>
<meta name="description" content="Fountain powerboat dealer for new fountain powerboats and used fountain powerboats and new boats and used boats with sales and service of many major lines">
<meta name="keywords" content="boat dealer, new fountain powerboats, used fountain powerboats, fountains, used fountains, new boats, used boats, north carolina boat dealer, high performance boats, fishing boats">
<style>
body{scrollbar-base-color: #999999}
</style>

<style fprolloverstyle>A:hover {color: #646565; font-family: Verdana}
</style>

</head>

<link href="css_styles/site.css" rel="stylesheet" type="text/css">

<body topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" bgcolor="#010098" link="#AA0100" vlink="#AA0100" alink="#AA0100">

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber1" bgcolor="#FFFFFF">
  <tr>
    <td width="100%" valign="top">
    <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber2" background="images/gmtopback.jpg">
      <tr>
        <td width="63%" valign="top">
        <a href="http://www.greenvillemarine.com/">
        <img border="0" src="images/gmtopleft.jpg"></a></td>
        <td width="37%" valign="top">
        <p align="right">
        <img border="0" src="images/gmtopright.jpg"></td>
      </tr>
    </table>
    </td>
  </tr>
  <tr>
    <td width="100%" valign="top">
    <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber4">
      <tr>
        <td width="180" valign="top" background="images/gmsideback.jpg">
        <!--webbot bot="Include" U-Include="navside.htm" TAG="BODY" startspan -->

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="167" id="AutoNumber1">
  <tr>
    <td width="100%" valign="top">
    </td>
  </tr>
  <tr>
    <td width="100%" valign="top">
    <a href="http://www.greenvillemarine.com/inventory/search.asp">
    <img border="0" src="images/navslice2.jpg" alt="Browse Our Current Inventory Of New &amp; Used Sea Ray, Sea Pro, G3 Boats, May-Craft, Carolina Skiff, Swamp Duck,  &amp; Fountain Powerboats." width="167" height="49"></a></td>
  </tr>
  <tr>
    <td width="100%" valign="top">
    <a href="http://www.greenvillemarine.com/tournaments.asp">
    <img border="0" src="images/navslice3.jpg" alt="Click Here For Tournaments, Events, &amp; Photo Gallery" width="167" height="49"></a></td>
  </tr>
    <tr>
    <td width="100%" valign="top">
    <a href="http://www.greenvillemarine.com/trout.asp">
    <img border="0" src="images/navslice3-1.jpg" alt="Click Here For Tournaments, Events, &amp; Photo Gallery" width="167" height="49"></a></td>
  </tr>
  <tr>
    <td width="100%" valign="top">
    <a href="http://www.greenvillemarine.com/tackle.htm">
    <img border="0" src="images/navslice4.jpg" alt="Click Here For Greenville Marine Tackle Shop" width="167" height="30"></a></td>
  </tr>
  <tr>
    <td width="100%" valign="top">
    <a href="http://www.greenvillemarine.com/boats.htm">
    <img border="0" src="images/navslice5.jpg" alt="Click Here For Greenville Marine Boat Lines" width="167" height="30"></a></td>
  </tr>
  <tr>
    <td width="100%" valign="top">
    <a href="http://www.greenvillemarine.com/motors.htm">
    <img border="0" src="images/navslice6.jpg" alt="Click Here For Greenville Marine Motor Lines" width="167" height="34"></a></td>
  </tr>
  <tr>
    <td width="100%" valign="top">
    <a href="http://www.greenvillemarine.com/trailers.htm">
    <img border="0" src="images/navslice7.jpg" alt="Click Here For Greenville Marine Trailer Lines" width="167" height="30"></a></td>
  </tr>
  <tr>
    <td width="100%" valign="top">
    <a href="http://www.greenvillemarine.com/boatparts.htm">
    <img border="0" src="images/navslice7b.jpg" alt="Click Here For Greenville Marine Boat and Engine Parts" width="167" height="49"></a></td>
  </tr>
  <tr>
    <td width="100%" valign="top">
    <a href="http://www.greenvillemarine.com/map.htm">
    <img border="0" src="images/navslice8.jpg" alt="Click Here For Greenville Marine Location" width="167" height="30"></a></td>
  </tr>
  <tr>
    <td width="100%" valign="top">
    <a href="http://www.greenvillemarine.com/contact.asp">
    <img border="0" src="images/navslice9.jpg" alt="Click Here For Greenville Marine Contact Information" width="167" height="34"></a></td>
  </tr>
  <tr>
    <td width="100%" valign="top"><a href="http://www.greenvillemarine.com">
    <img border="0" src="images/navslice10.jpg" alt="Click Here To Return To Greenville Marine Homepage" width="167" height="31"></a></td>
  </tr>
</table>

<!--webbot bot="Include" i-checksum="36622" endspan --><p>&nbsp;</td>
        <td valign="top">
        <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="97%" id="AutoNumber5">
          <tr>
            <td width="100%"><img border="0" src="images/space.gif"></td>
          </tr>
          <tr>
            <td width="100%" background="images/headerback.gif">
            <img border="0" src="images/headertournaments.gif"></td>
          </tr>
          <tr>
            <td width="100%">



<table border="1" cellspacing="1" style="border-collapse: collapse" bordercolor="#FFFFFF" width="100%" id="AutoNumber8" cellpadding="0">
  <tr>
    <td width="33%" bgcolor="#D9D9FF" align="center">
    <span lang="en-us"><b><font face="Verdana" size="1">
    <a style="text-decoration: none" href="tournaments.asp">Tournaments &amp; Events</a></font></b></span></td>
    <td width="33%" bgcolor="#D9D9FF" align="center">
    <span lang="en-us"><b><font face="Verdana" size="1">
    <a style="text-decoration: none" href="tournamentresults.asp">Tournament Results</a></font></b></span></td>
    <td width="34%" bgcolor="#D9D9FF" align="center">
    <span lang="en-us"><b><font face="Verdana" size="1">
    <a style="text-decoration: none" href="http://www.greenvillemarine.com/photogallery/index.asp?CatID=6">
    Photo Gallery</a></font></b></span></td>
  </tr>
</table>



<table width="100%" border="0">
  <tr> 
    <td width="35%"> 
      &nbsp;<table width="100%" border="0">
        <tr> 
          <% If rsNews.EOF And rsNews.BOF Then %>
          <td class="ContentBody"><font face="Verdana, Arial, Helvetica, sans-serif">Sorry 
            but there are no <span lang="en-us">tournament results posted</span></font><span lang="en-us">.</span></td>
          <% End If ' end rsNews.EOF And rsNews.BOF %>
        </tr>
        <% 
While ((Repeat1__numRows <> 0) AND (NOT rsNews.EOF)) 
%>
        <tr> 
          <td class="ContentHead"><%=(rsNews.Fields.Item("SUBJECT").Value)%></td>
        </tr>
        <tr> 
          <td class="ContentBody"><%=(rsNews.Fields.Item("BODY").Value)%><br> <%
			if rsNews.Fields.Item("RELATEDLINK").Value <> "" then
			%>
            Related Link: <a href="<%=(rsNews.Fields.Item("RELATEDLINK").Value)%>"><%=(rsNews.Fields.Item("RELATEDLINK").Value)%></a><br><% end if %> </td>
          
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
            <p align="center">&nbsp;<p align="center">&nbsp;<p>&nbsp;</p>



<table border="1" cellspacing="1" style="border-collapse: collapse" bordercolor="#FFFFFF" width="100%" id="AutoNumber7" cellpadding="0">
  <tr>
    <td width="33%" bgcolor="#D9D9FF" align="center">
    <span lang="en-us"><b><font face="Verdana" size="1">
    <a style="text-decoration: none" href="tournaments.asp">Tournaments &amp; Events</a></font></b></span></td>
    <td width="33%" bgcolor="#D9D9FF" align="center">
    <span lang="en-us"><b><font face="Verdana" size="1">
    <a style="text-decoration: none" href="tournamentresults.asp">Tournament Results</a></font></b></span></td>
    <td width="34%" bgcolor="#D9D9FF" align="center">
    <span lang="en-us"><b><font face="Verdana" size="1">
    <a style="text-decoration: none" href="http://www.greenvillemarine.com/photogallery/index.asp?CatID=6">
    Photo Gallery</a></font></b></span></td>
  </tr>
</table>



            </td>
          </tr>
          <tr>
            <td width="100%" background="images/footerb2.gif">
            <img border="0" src="images/footerb1.gif"></td>
          </tr>
          <tr>
            <td width="100%">
            <!--webbot bot="Include" U-Include="navbot.htm" TAG="BODY" startspan -->

<p align="center"><span lang="en-us">
<font face="Verdana" size="1">
<a href="fountain/default.htm" style="text-decoration: none">Boat Inventory</a> || 
<a href="tournaments.asp" style="text-decoration: none">Bass Tournaments &amp; Events</a> || 
<a href="trout.asp" style="text-decoration: none">Trout Tournaments &amp; Events</a><br>
<a href="tackle.htm" style="text-decoration: none">Tackle</a> || 
<a href="boats.htm" style="text-decoration: none">Boats</a> || 
<a href="motors.htm" style="text-decoration: none">Motors</a> || 
<a href="trailers.htm" style="text-decoration: none">Trailers</a> || 
<a href="map.htm" style="text-decoration: none">Location</a> || 
<a href="contact.htm" style="text-decoration: none">Contact</a> || 
<a href="default.asp" style="text-decoration: none">Home</a></font></span></p>

<!--webbot bot="Include" i-checksum="22608" endspan --></td>
          </tr>
        </table>
        </td>
      </tr>
    </table>
    </td>
  </tr>
  <tr>
    <td width="100%" valign="top">
    <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber3" background="images/gmbotback.jpg">
      <tr>
        <td width="44%" valign="top">
        <img border="0" src="images/gmbotleft6.jpg"></td>
        <td width="56%" valign="top">
        <p align="right">
        <img border="0" src="images/gmbotright2.jpg" usemap="#FPMap1"></td>
      </tr>
    </table>
    </td>
  </tr>
</table>


<%
rsNews.Close()
%>