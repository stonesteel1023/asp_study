<%@ Language=VBScript %>
<html>

<head>
<title></title>
<style type="text/css">  
 A            {text-decoration: none; color:white }  
 A:hover      {text-decoration: none }  
</style>
</head>

<body bgcolor="#00274f" text="#eefde3">

<table border="0" width="120" height="30" cellpadding="2" cellspacing="0">
  <tr>
    <td align="left" height="80" width="120"><p align="center"><font face="Comic Sans MS"
    color="#ffff00"><strong><big><big><big>&quot;toys&quot;</big></big></big></strong></font></td>
  </tr>
  <tr>
    <td align="left" bgcolor="#eefde3" height="40" width="120"><strong><a href="toy/cart.asp" target="right">&nbsp; <font face="µ¸¿ò" size="2"
    color="#00274f">ShoppingCart</font></a></strong></td>
  </tr>
  <tr>
    <td align="left" height="20" width="120">&nbsp; </td>
  </tr>
  <tr>
    <td align="left" bgcolor="#eefde3" height="40" width="120"><font face="µ¸¿ò" size="2"
    color="#00274f"><strong>&nbsp; ¼îÇÎ</strong></font></td>
  </tr>
  <tr>
    <td align="left"
    style="BORDER-LEFT: rgb(238,253,227) 1px dashed; BORDER-RIGHT: rgb(238,253,227) 1px dashed"
    height="25" width="120"><font face="µ¸¿ò" size="2"><strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
    </strong></font><font color="#ffff80">&gt;</font><font face="µ¸¿ò" size="2"><strong> <a
    href="toy/list.asp?part=Doll" target="right"><font color="#ffffff">ÀÎÇü</font></a></strong></font></td>
  </tr>
  <tr>
    <td align="left"
    style="BORDER-LEFT: rgb(238,253,227) 1px dashed; BORDER-RIGHT: rgb(238,253,227) 1px dashed"
    height="25" width="120"><font face="µ¸¿ò" size="2"><strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
    </strong></font><font color="#80ffff">&gt;</font><font face="µ¸¿ò" size="2"><strong> <a
    href="toy/list.asp?part=Game" target="right"><font color="#ffffff">°ÔÀÓ</font></a></strong></font></td>
  </tr>
  <tr>
    <td align="left"
    style="BORDER-BOTTOM: rgb(238,253,227) 1px dashed; BORDER-LEFT: rgb(238,253,227) 1px dashed; BORDER-RIGHT: rgb(238,253,227) 1px dashed"
    height="25" width="120"><font face="µ¸¿ò" size="2"><strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
    </strong></font><font color="#80ff80">&gt;</font> <a href="toy/list.asp?part=Rc"
    target="right"><font color="#ffffff" face="µ¸¿ò" size="2"><strong>RCÄ«</strong></font></a></td>
  </tr>
  <tr>
    <td align="left" height="20" width="120">&nbsp; </td>
  </tr>
<% if session("id")<>"admin" then %>
  <tr>
    <td align="left" bgcolor="#eefde3" height="40" width="120"><strong><a
    href="admin/login.asp" target="right"><font face="µ¸¿ò" size="2" color="#00274f">&nbsp; 
    °ü¸®ÀÚ ·Î±ä</font></a></strong></td>
  </tr>
<% else%>
  <tr>
    <td align="left" height="30" width="120"><strong><font face="µ¸¿ò" size="2">&nbsp; <a
    href="admin/main.asp" target="right">°ü¸®ÀÚ ±¸¿ª</a></font></strong></td>
  </tr>
<% end if%>
  <tr>
    <td align="left" height="20" width="120">&nbsp; </td>
  </tr>
</table>
</body>
</html>
