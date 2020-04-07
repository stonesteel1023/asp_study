<!--#include file="../inc/connection.asp"-->
<%
code = request("code")
id = Request.Cookies("user")("id")

Set Dbcon= Server.CreateObject("ADODB.Connection")
Dbcon.Open strConnect

sql = "select Cc_boss, Cc_title, Cc_date, Cc_memcount, Cc_description from Circle "
sql = sql & " where Cc_Code='" & code & "'"

set Rs = Dbcon.Execute(sql)

boss = rs(0)
title = rs(1)
dd = rs(2)
memcount = rs(3)
content = rs(4)
content = replace(content, chr(13)&chr(10), "<br>")
%>
<HTML>
<HEAD>
<link rel="stylesheet" href="../css/main.css">
</HEAD>
<BODY leftmargin="30">
<table width="400" border=1 bgcolor=#ffffff cellspacing=0>
<tr>
  <td colspan=2><br><center><h3><%=title%></h3></center>
  </td></tr>
<tr bgcolor="#2f4f5c">
  <td colspan=2 width=400 height=30 >
  <center><font class="blue">커뮤니티 참여</center>
  </td></tr>
<tr height=30>
  <td width=100 align="center">써클장</td>
  <td width=300 style="padding-left:20"><%=boss%>
  </td></tr>
<tr height=30>
  <td width=100 align="center">모임주제</td>
  <td width=300 style="padding-left:20">   <%=title%>
  </td></tr>
<tr height=30>
  <td width=100 align="center">탄생일자</td>
  <td width=300 style="padding-left:20"><%=dd%>
  </td></tr>
<tr height=30>
  <td width=100 align="center">현재인원</td>
  <td width=300 style="padding-left:20">
   <%=memcount%> &nbsp;&nbsp;
  </td></tr>
<tr height=30>
  <td colspan=2 bgcolor="#BBCAE8" align="center">모임주제 설명
  </td></tr>
<tr>  
  <td colspan=2 style="padding-left:20; padding-right:20; padding-top:20; padding-bottom:20">
   <%=content%>
  </td></tr>
</table>
</BODY></HTML>
