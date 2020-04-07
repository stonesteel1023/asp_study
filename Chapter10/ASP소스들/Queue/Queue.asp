<% Option Explicit %>
<!--#include file="../inc/connection.asp"-->
<%
Dim id, sql
Dim dbcon, Rs

id = Request.Cookies("user")("id")

Set Dbcon= Server.CreateObject("ADODB.Connection")
Dbcon.Open strConnect

sql = "select Que_code, Que_boss, Que_title, Que_date, Que_memcount, Que_member from Queue"
set Rs = Dbcon.Execute(sql)
%>
	<html>
	<head><link rel="stylesheet" href="../css/main.css"></head>
	<body>
	<div align="center">

<!-- 신청된 커뮤티니 목록 출력하기 시작 -->
	<table width="600" border="0" bgcolor="#fffff4" cellspacing="0">
	<tr bgcolor="#BBCAE8">
	  <td width="100">&nbsp;</td>
	  <td width="350" height="30"><center><font>현재 대기중인 커뮤니티..</font></center></td>
	  <td align="center" width="150">
	  <input type="button" name="MakeQueue" value="커뮤니티 신청" onClick="javascript:location.href='Que_form.asp'" style="font-family:돋움;color:#004080;background-color:#efeff4;border:1 solid black;height:20">
	  </td>
	</tr>
	</table><br>
	<table width="600" border="1" bgcolor="#ffffff" cellspacing="0">
<% if Rs.BOF and Rs.EOF then %>
	<tr>
	  <td height="70" bgcolor="#efefff"><center><font><font>현재는 없습니다.</font></font></center></td>
	</tr>
<% else %>
	<tr height="25" bgcolor="khaki">
	<td align="center">요청자</td>
	<td align="center">커뮤니티 제목</td>
	<td align="center">인원</td>
	<td align="center">신청일자</td>
	</tr>
<% 
		Dim code, boss, title, dd, memcount, mem, BoardName
		Do until Rs.EOF 
			code = rs(0)
			boss = rs(1)
			title = rs(2)
			dd = year(rs(3)) & "/" & month(rs(3)) & "/" & day(rs(3))
			memcount = rs(4)
			mem = Rs(5)
%>
			<tr height="25">
			<td align="center"><%=boss%></td>
			<td align="center">
			<a href="Que_description.asp?code=<%=code%>"><%=title%></a></td>
			<td align="center"><%=memcount%></td>
			<td align="center"><%=dd%></td>
			</tr>
<% 
			Rs.movenext
		Loop

End if
Rs.close
%></table><br><br>
<!-- 신청된 커뮤티니 목록 출력하기 끝 -->
<hr color="silver" size="2" width="600"><br><br>
<!-- 운영중인 커뮤티니 목록 출력하기 시작 -->
<%
sql = "select CC_code, CC_boss, CC_title, CC_date, CC_memcount, Cc_BoardName from Circle"
set Rs = Dbcon.Execute(sql)
%>
<table width="600" border="0" bgcolor="#fffff4" cellspacing="0">
<tr bgcolor="#2f4f5c">
  <td width="100">&nbsp;</td>
  <td width="350" height="30"><center><font class="blue">현재 운영중인 커뮤니티..</font></center>
</tr>
</table><br>
<table width="600" border="1" bgcolor="#ffffff" cellspacing="0">
<% if Rs.BOF and Rs.EOF then %>
<tr>
  <td height="70" bgcolor="#efefff"><center>현재는 없습니다.</center></td>
</tr>
<% else %>
	<tr height="25" bgcolor="#BBCAE8">
	<td align="center">써클장</td>
	<td align="center">커뮤니티 제목</td>
	<td align="center">인원</td>
	<td align="center">신청일자</td>
	</tr>

<% 
		Do until Rs.EOF 
			code = rs(0)
			boss = rs(1)
			title = rs(2)
			dd = year(rs(3)) & "/" & month(rs(3)) & "/" & day(rs(3))
			memcount = rs(4)
			BoardName = rs(5)
%>
			<tr height="25">
			<td align="center"><%=boss%></td>
			<td align="center">
			<a href="../commu/join.asp?code=<%=code%>&BoardName=<%=BoardName%>"><%=title%>
			</a></td>
			<td align="center"><%=memcount%></td>
			<td align="center"><%=dd%></td>
			</tr>
<% 
			Rs.movenext
		Loop

End if
Rs.close
set Rs = nothing
Dbcon.Close
Set Dbcon = nothing
%>
<!-- 운영중인 커뮤티니 목록 출력하기 끝 -->
</table>
</div></body>
</html>
