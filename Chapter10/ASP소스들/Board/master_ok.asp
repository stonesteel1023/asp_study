<!--#include file="../inc/connection.asp"-->
<%
table=request("table")
id = Request.Cookies("user")("id")

Set Dbcon= Server.CreateObject("ADODB.Connection")
Dbcon.Open strConnect

sql = "select Cc_boss from Circle "
sql = sql & " where Cc_BoardName='" & table & "'"

set Rs = Dbcon.Execute(sql)
boss = rs(0)
Rs.close
Set Rs = nothing
%>
<html>
<head>
<title>Board</title>
</head>
<body>
<% if boss = id then %>
<p align="center"><font face="Century Gothic" size="6" color="#0000FF"><strong>Login 
Success</strong></font></p>

<p align="center"><font face="����" size="2" color="#004080">�α��� ���������� 
����Ǿ����ϴ�..</font></p>

<p align="center"><font face="����" size="2" color="#004080">�����ڴ��� 
�Խ��ǿ����� �۵��� ������ ���Ƕ� </font></p>

<p align="center"><font face="����" size="2" color="#004080">�� �۵��� 
��й�ȣ�� ���� ������ ���Դϴ�..</font></p>

<p align="center"><font face="����" size="2" color="#004080">�켱�� �װ����� 
����,���� �Ͻñ� �ٶ��ϴ�..</font></p>

<p align="center"><font face="����" size="2">��</font></p>

<p align="center"><font face="����" size="2">&lt; <a href="List.asp?table=<%=table%>"><font color="#008080">�������� �ź����� �Խ������� ���ϴ�</font></a><font color="#808080">.</font> &gt;</font></p>

<p align="center"><font face="����" size="2">&lt; <font color="#408080"><a href="master_logoff.asp?table=<%=table%>">�Ϲ����� �ź����� �Խ������� ���ϴ�</a></font><font color="#808080">.</font> &gt;</font></p>

<% else %>
<p align="center"><font face="Century Gothic" size="6" color="#FF0000"><strong>Login 
Failed</strong></font></p>

<p align="center"><font face="����" size="2" color="#004080">��</font></p>

<p align="center"><font face="����" size="2" color="#004080">�α��� 
&nbsp; ���еǾ����ϴ�..</font></p>

<p align="center"><font face="����" size="2" color="#004080">�Ϲ� ����ڴԵ��� 
������ �޴��� ����Ͻ� �� �����ϴ�.</font></p>

<p align="center"><font face="����" size="2" color="#004080">�ڼ��� ������ 
��ڿ��� ���� �Ͻʽÿ�...</font></p>

<p align="center">��</p>

<p align="center"><font face="����" size="2">&lt; <font color="#008080"><a href="List.asp?table=<%=table%>">�Խ������� ���ư��ϴ�</a>.</font> &gt;</font></p>
<% end if %>
</body>
</html>
