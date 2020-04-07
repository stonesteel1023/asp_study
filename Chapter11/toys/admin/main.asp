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
    <td><p align="center">　</p>
    <form method="POST" action>
      <div align="center"><center><p><a href="../goods/new.htm"><font face="돋움" size="2">상품 
      등록</font></a></p>
      </center></div><div align="center"><center><p>
      <a href="list_buy.asp">
      <font face="돋움" size="2">구매 
      리스트 보기</font></a></p>
      </center></div><p>　</p>
      <p>　</p>
    </form>
    <p align="center">　</td>
  </tr>
</table>
</body>
</html>


<a href="../goods/new.htm"><font face="돋움" size="2">상품 
      등록</font></a>
      <a href="list_buy.asp">
      <font face="돋움" size="2">구매 
      리스트 보기</font></a>