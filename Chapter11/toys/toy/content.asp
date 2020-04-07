<%

 code = request("code")

 Set db = Server.CreateObject("ADODB.Connection")
 db.Open("dsn=toys;uid=sa;pwd=;")
	
 sql = "select * from goods where g_code='" & code & "'"
	
 set rs = db.Execute(sql)
 
 content = replace(rs("g_content"),chr(13),"<br>")
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
<title>content</title>
</head>
<script language="javascript">
function buyit()
{
   document.myform.submit();
}
</script>
<body bgcolor="#FFFFFF">

<form method="POST" name="myform" action="buy_temp.asp">
<input type="hidden" name="code" value="<%=rs("g_code")%>">
  <div align="left"><table border="1" cellpadding="0" cellspacing="0" width="650">
    <tr>
      <td align="center">　<div align="center"><center><table border="0" cellpadding="0" cellspacing="0" width="550">
        <tr>
          <td width="50"></td>
          <td width="450" bgcolor="#004080" style="border: 1 dashed rgb(0,64,128)" height="40"><div align="center"><center><table border="0" cellpadding="0" cellspacing="0" width="440" height="28" bgcolor="#D3D9F3">
            <tr>
              <td><span class="link"><font face="Comic Sans MS" size="2"><div align="center"><center><p></font><font size="3" color="#004080"><strong><%=rs("g_name")%></strong></font></span></td>
            </tr>
          </table>
          </center></div></td>
          <td width="50" align="center"></td>
        </tr>
      </table>
      </center></div><div align="center"><center><table border="0" cellpadding="0" cellspacing="0" width="509">
        <tr>
          <td height="25" width="610" colspan="2"></td>
        </tr>
        <tr>
          <td height="50" width="150" style="border-left: 1 dashed rgb(190,198,237); border-bottom: 1 dashed rgb(190,198,237)" rowspan="2"><div align="center"><center><p><img src="../images/<%=rs("g_image")%>.jpg" border="0"></td>
          <td height="25" width="460" style="border-left: 1 dashed rgb(190,198,237); border-right: 1 dashed rgb(190,198,237); border-top: 1 none rgb(190,198,237); border-bottom: 1 dashed rgb(190,198,237); padding-left: 10; padding-right: 10; padding-top: 5; 
        padding-bottom: 5" bgcolor="#FFFFFF" align="center"><div align="center"><center><p><input name="ea" style="font-family: 돋움; text-align: right; border: 1px dashed" size="5" value="1 "> <font size="2">EA&nbsp;&nbsp;&nbsp;&nbsp; 
        </font><a href="javascript:buyit();"><img src="../images/buy.gif" border="0" WIDTH="28" HEIGHT="18"></a></td>
        </tr>
        <tr align="center">
          <td height="25" width="460" style="border-left: 1 dashed rgb(190,198,237); border-right: 1 dashed rgb(190,198,237); border-top: 1 none rgb(190,198,237); border-bottom: 1 dashed rgb(190,198,237); padding-left: 10; padding-right: 10; padding-top: 5; padding-bottom: 5" bgcolor="#F0F2FB"><div align="left"><p><font size="2" color="navy"><%=rs("g_name")%></font><br>
          <font face="Comic Sans MS" size="2"></font><font size="2">제조회사 </font>
          <font face="Comic Sans MS" size="2">: </font><font size="2"> <%=rs("g_maker")%> </font><font face="Comic Sans MS" size="2"><br>
          <br>
          </font><font size="2">기존가 </font><font face="Comic Sans MS" size="2">:</font><font face="돋움" size="2"> <strike> <%=formatcurrency(rs("g_origin_price"))%></strike></font><br>
          <font size="2">판매가 </font><font face="Comic Sans MS" size="2">:</font> <font face="돋움" size="2"> <%=formatcurrency(rs("g_sellprice"))%></font><p>
          <font size="2">데이터 수정일 </font><font face="Comic Sans MS" size="2">: </font><font size="2"><%=rs("g_updateday")%></font></td>
        </tr>
        <tr align="center">
          <td height="25" width="610" colspan="2"></td>
        </tr>
      </table>
      </center></div><div align="center"><center><table border="0" cellpadding="0" cellspacing="0" width="509">
        <tr>
          <td height="80" style="border: 1 dashed rgb(196,196,255); padding-left: 20; padding-right: 20; padding-top: 10; padding-bottom: 10"><span class="blk-txt"><font face="Arial" size="2" color="#000040">&lt; </font><font size="2" color="#000040">제품설명 </font><font face="Arial" size="2" color="#000040">&gt;</font></span><p><font face="Arial" size="2" color="#000040"><%=content%></font></td>
        </tr>
        <tr>
          <td height="30"></td>
        </tr>
      </table>
      </center></div></td>
    </tr>
  </table>
  </div>
</form>
</body>
</html>
