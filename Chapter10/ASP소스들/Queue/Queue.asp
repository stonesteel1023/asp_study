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

<!-- ��û�� Ŀ��Ƽ�� ��� ����ϱ� ���� -->
	<table width="600" border="0" bgcolor="#fffff4" cellspacing="0">
	<tr bgcolor="#BBCAE8">
	  <td width="100">&nbsp;</td>
	  <td width="350" height="30"><center><font>���� ������� Ŀ�´�Ƽ..</font></center></td>
	  <td align="center" width="150">
	  <input type="button" name="MakeQueue" value="Ŀ�´�Ƽ ��û" onClick="javascript:location.href='Que_form.asp'" style="font-family:����;color:#004080;background-color:#efeff4;border:1 solid black;height:20">
	  </td>
	</tr>
	</table><br>
	<table width="600" border="1" bgcolor="#ffffff" cellspacing="0">
<% if Rs.BOF and Rs.EOF then %>
	<tr>
	  <td height="70" bgcolor="#efefff"><center><font><font>����� �����ϴ�.</font></font></center></td>
	</tr>
<% else %>
	<tr height="25" bgcolor="khaki">
	<td align="center">��û��</td>
	<td align="center">Ŀ�´�Ƽ ����</td>
	<td align="center">�ο�</td>
	<td align="center">��û����</td>
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
<!-- ��û�� Ŀ��Ƽ�� ��� ����ϱ� �� -->
<hr color="silver" size="2" width="600"><br><br>
<!-- ����� Ŀ��Ƽ�� ��� ����ϱ� ���� -->
<%
sql = "select CC_code, CC_boss, CC_title, CC_date, CC_memcount, Cc_BoardName from Circle"
set Rs = Dbcon.Execute(sql)
%>
<table width="600" border="0" bgcolor="#fffff4" cellspacing="0">
<tr bgcolor="#2f4f5c">
  <td width="100">&nbsp;</td>
  <td width="350" height="30"><center><font class="blue">���� ����� Ŀ�´�Ƽ..</font></center>
</tr>
</table><br>
<table width="600" border="1" bgcolor="#ffffff" cellspacing="0">
<% if Rs.BOF and Rs.EOF then %>
<tr>
  <td height="70" bgcolor="#efefff"><center>����� �����ϴ�.</center></td>
</tr>
<% else %>
	<tr height="25" bgcolor="#BBCAE8">
	<td align="center">��Ŭ��</td>
	<td align="center">Ŀ�´�Ƽ ����</td>
	<td align="center">�ο�</td>
	<td align="center">��û����</td>
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
<!-- ����� Ŀ��Ƽ�� ��� ����ϱ� �� -->
</table>
</div></body>
</html>
