<!--#include file="../inc/connection.asp"-->
<%
code = request("code")
BoardName = request("BoardName")
id = Request.Cookies("user")("id")

Set Dbcon= Server.CreateObject("ADODB.Connection")
Dbcon.Open strConnect

sql = "select mem_id from CircleMember Where Cc_BoardName ='" & BoardName & "'"
sql = sql & " Group by Cc_BoardName, mem_id"

set Rs = Dbcon.Execute(sql)
%>
<HTML>
<HEAD>
<link rel="stylesheet" href="../css/main.css">
</HEAD>
<BODY leftmargin="30">
<table width="400" border=1 bgcolor=#ffffff cellspacing=0>
<tr bgcolor="#2f4f5c">
  <td colspan=2 width=400 height=30 >
  <center><font class="blue">커뮤니티 참여회원 목록</center>
  </td></tr>
<tr>
  <td align="center" style="padding-top:30;padding-bottom:30">
  <% Do Until Rs.EOF 
		mem_id = rs(0)
		if mem_id = id then
			Response.Write "<img src='images/join.gif' align=absmiddle> "
		end if
  %>
			<%=mem_id%> 님<br>
  <% Rs.movenext
	 Loop
  %>
      </td>
  </tr>
</table>
</BODY></HTML>
