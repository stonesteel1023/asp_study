<%@ Language=VBScript %>
<%
if session("id")<>"admin" then
  response.redirect "../default.htm"
end if 
%>
<html>

<head>
 <title></title>
</head>

<body bgcolor="#FFFFFF">

<table border="0" width="500">
  <tr>
    <td></td>
  </tr>
  <tr>
    <td><p align="center">��</p>
    <form method="POST" action>
      <div align="center"><center><p><a href="../goods/new.htm"><font face="����" size="2">��ǰ 
      ���</font></a></p>
      </center></div><div align="center"><center><p>
      <a href="list_buy.asp">
      <font face="����" size="2">���� 
      ����Ʈ ����</font></a></p>
      </center></div><p>��</p>
      <p>��</p>
    </form>
    <p align="center">��</td>
  </tr>
</table>
</body>
</html>


<a href="../goods/new.htm"><font face="����" size="2">��ǰ 
      ���</font></a>
      <a href="list_buy.asp">
      <font face="����" size="2">���� 
      ����Ʈ ����</font></a>