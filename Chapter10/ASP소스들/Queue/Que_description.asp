<% Option Explicit %>
<!--#include file="../inc/connection.asp"-->
<%
Dim code, id, sql
Dim Dbcon, Rs

code = request("code")
id = Request.Cookies("user")("id")

Set Dbcon= Server.CreateObject("ADODB.Connection")
Dbcon.Open strConnect

sql = "select Que_boss, Que_title, Que_date, Que_memcount, Que_Member, Que_description from Queue "
sql = sql & " where Que_code=" & code

set Rs = Dbcon.Execute(sql)

Dim boss, title, dd, memcount, member, mem, joinCheck, content

boss = rs(0)
title = rs(1)
dd = rs(2)
memcount = rs(3)
member = rs(4)
joinCheck  = instr(member, id)
content = rs(5)
content = replace(content, chr(13)&chr(10), "<br>")
%>
<HTML>
<HEAD>
<link rel="stylesheet" href="../css/main.css">
<script language="javascript">
function createComm()
{
   document.queue.action = "Que2Comm.asp";
   document.queue.submit();
}
</script>   
</HEAD>
<BODY>
<div align="center">
<form name="queue" Method="post" action="Que_join.asp">
<input type="hidden" name="code" value="<%=code%>">

<table width="400" border=1 bgcolor=#ffffff cellspacing=0>
<tr bgcolor="#2f4f5c">
  <td colspan=2 width=400 height=30 >
  <center><font class="blue">Ŀ�´�Ƽ ����</center></td></tr>
<tr height=30>
  <td width=100 align="center">��û��</td>
  <td width=300 style="padding-left:20"><%=boss%></td></tr>
<tr height=30>
  <td width=100 align="center">��������</td>
  <td width=300 style="padding-left:20"><%=title%></td></tr>
<tr height=30>
  <td width=100 align="center">��û����</td>
  <td width=300 style="padding-left:20"><%=dd%></td></tr>
<tr height=30>
  <td width=100 align="center">�����ο�</td>
  <td width=300 style="padding-left:20">
   <%=memcount%> &nbsp;&nbsp;(5�� �̻��� �����ؾ� �����˴ϴ�)</td></tr>
<tr height=30>
  <td width=100 align="center">������</td>
  <td width=300 style="padding-left:20"><%=member%></td></tr>
<tr height=30>
  <td colspan=2 bgcolor="#BBCAE8" align="center">�������� ����</td></tr>
<tr>  
  <td colspan=2 style="padding-left:20; padding-right:20; padding-top:20; padding-bottom:20">
   <%=content%></td></tr>
<tr height=40>
  <td colspan=2><center>
  <% if int(joinCheck) = 0 then%>
  <input type="submit" value="Ŀ�´�Ƽ ����" style="background-color:#efeff4" >
  <%end if%>
  <input type="button" value="�ڷ� ���ư���" OnClick="javascript:history.back();"
  style="background-color:#efeff4" >
  <% if id ="admin" then%>
  <input type="button" value="Ŀ�´�Ƽ ����" OnClick="javascript:createComm();"
  style="background-color:#efeff4" >
  <% end if%>
  </center>
  </td></tr>  
</table>
</form></div>
</BODY></HTML>
