<%

Response.Expires=0

 Set db = Server.CreateObject("ADODB.Connection")
 db.Open("dsn=toys;uid=sa;pwd=;")
	
 sql = "select imsi_ea,g_code, g_name, g_sellprice from imsi_buy A, goods B  where A.imsi_memid ='" & session.SessionID & "'"
 sql = sql & " and A.imsi_goodscode = B.g_code"
 	
 Set rs = Server.CreateObject("ADODB.RecordSet")
 rs.Open sql, db ,1
 %>
<html>

<head>
<title></title>
</head>
<script language="javascript">
function update()
{
   document.myform.action = "update.asp";
   document.myform.submit();
}

function order()
{
   document.myform.action = "order.asp";
   document.myform.submit();
}
</script>
<body bgcolor="#ffffff">
<% if rs.EOF then %>

<p><br><br>
<p align="center"><font face="Comic Sans MS" size="7" color="#000040"><strong>Empty !!</strong></font></p>

<p align="center"><font face="Comic Sans MS" size="7" color="#008080"><strong>¡¡</strong></font></p>

<p align="center"><font face="µ¸¿ò" size="2" color="#004080">ÇöÀç </font><strong><font face="µ¸¿ò" color="#004080" size="4">Cart</font></strong><font face="µ¸¿ò" size="2" color="#004080"> ¿¡´Â ¾Æ¹«°Íµµ ¾ø½À´Ï´Ù.</font></p>

<p align="center"><font face="µ¸¿ò" size="2" color="#004080">Áñ°Å¿î ¼îÇÎÀÌ 
µÇ½Ê½Ã¿ä</font></p>


<% else %> 

<p>¡¡</p>
<div align="center"><center>

<table border="0" cellpadding="0" width="300">
  <tr>
    <td width="0" height="30"><p align="center"><font face="±¼¸²" size="4"><strong>Àå¹Ù±¸´Ï 
    º¸±â (ShoppingCart)</strong></font></td>
  </tr>
</table>
</center></div>

<form method="POST" name="myform">
  <div align="center"><center><table border="0" cellpadding="0" width="550" cellspacing="0" height="64">
    <tr>
      <td height="30" width="63" align="center" bgcolor="#00356A"><strong><font face="µ¸¿ò" size="2" color="#FFFFFF">»èÁ¦</font></strong></td>
      <td height="30" width="299" align="center" bgcolor="#00356A"><strong><font face="µ¸¿ò" size="2" color="#FFFFFF">±¸ÀÔÇ°</font></strong></td>
      <td height="30" width="73" align="center" bgcolor="#00356A"><strong><font face="µ¸¿ò" size="2" color="#FFFFFF">¼ö·®</font></strong></td>
      <td height="30" width="115" align="center" bgcolor="#00356A"><strong><font face="µ¸¿ò" size="2" color="#FFFFFF">°¡°Ý</font></strong></td>
    </tr>
    <tr>
      <td height="15" width="550" align="center" colspan="4" bgcolor="#FBF9C6"></td>
    </tr>
<% do until rs.EOF %>
    <tr>
      <td height="30" width="63" align="center" style="border-bottom: 1 dashed rgb(128,128,128)"><font face="µ¸¿ò" size="2">  
      <a href="del.asp?code=<%=rs("g_code")%>"><img src="../images/del.gif" border="0" WIDTH="25" HEIGHT="15"></a>
        </font></td>
      <td height="30" width="299" align="center" style="border-bottom: 1 dashed rgb(128,128,128)"><font face="µ¸¿ò" size="2"><%=rs("g_name")%></font></td>
      <td height="30" width="73" align="center"><font face="µ¸¿ò" size="2">
      <input type="hidden" name="code" value="<%=rs("g_code")%>">
      <input type="text" style="font-family: µ¸¿ò; text-align: right; border: 1px dashed" size="5" name="ea" value="<%=rs("imsi_ea")%> "> EA</font></td>
      <td height="30" width="115" align="center" style="border-bottom: 1 dashed rgb(128,128,128); padding-right: 15"><font face="µ¸¿ò" size="2"><div align="right"><p><%=formatcurrency(rs("g_sellprice"))%></font></td>
    </tr>
<% 
sum = sum + (rs("g_sellprice")*rs("imsi_ea"))
rs.MoveNext
loop
%>
    
    <input type="hidden" name="sum" value="<%=sum%>">
    <tr>
      <td height="30" width="550" align="center" colspan="4"></td>
    </tr>
    <tr>
      <td height="30" width="362" align="center" colspan="2"><font face="Arial Rounded MT Bold" size="3">
      <a href="javascript:update();"><img src="../images/update.gif" border="0" WIDTH="70" HEIGHT="30"></a>&nbsp;&nbsp; 
      <a href="reset.asp"><img src="../images/reset.gif" border="0" WIDTH="70" HEIGHT="30"></a>&nbsp;&nbsp;
      <a href="javascript:order();"><img src="../images/order.gif" border="0" WIDTH="70" HEIGHT="30"></a></font></td>
      <td height="30" width="73" align="center"><font face="µ¸¿ò" size="2">ÇÕ°è</font></td>
      <td height="30" width="115" align="center" style="padding-right: 15"><font face="µ¸¿ò" size="2"><%=formatcurrency(sum)%> </font></td>
    </tr>
  </table>
  </center></div>
</form>
<% end if%>
</body>
</html>
