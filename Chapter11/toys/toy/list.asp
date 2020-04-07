<%

 part = request("part")

 Set db = Server.CreateObject("ADODB.Connection")
 db.Open("dsn=toys;uid=sa;pwd=;")
	
 sql = "select g_code,g_part,g_name,g_maker,g_sellprice,g_image,g_updateday from goods "
 sql = sql & " where g_part='" & part & "'"
	
 set rs = Server.CreateObject("ADODB.RecordSet")
 
 rs.open sql,db,1
 
%>
<html>

<head>
<title><%=part%></title>
</head>

<body bgcolor="#FFFFFF">
<% if rs.EOF then %>
<% else %>
<div align="left">

<table border="0" width="510" cellpadding="0" style="margin-left: 10">
  <tr>
    <td height="40">　<div align="center"><center><table border="0" cellpadding="0" cellspacing="0" width="300">
      <tr>
        <td bgcolor="#004080" style="border: 1 dashed rgb(0,64,128)" height="40"><p align="center"><font color="#FFFFFF"><big><strong><big><%=part%></big></strong></big></font></td>
      </tr>
    </table>
    </center></div><div align="center"><center><table border="0" cellpadding="0" cellspacing="0" width="500">
<!-- 시작  //-->
<%
 Do until rs.EOF 

%>
      <tr>
        <td height="25" width="500" colspan="2"></td>
      </tr>
      <tr>
        <td height="50" width="150" style="border-left: 1 dashed rgb(190,198,237); border-bottom: 1 dashed rgb(190,198,237)" rowspan="2"><p align="center"><a href="content.asp?code=<%=rs("g_code")%>"><img src="../images/<%=rs("g_image")%>_s.jpg" border="0"></a></td>
        <td height="25" width="350" style="border-left: 1 dashed rgb(190,198,237); border-right: 1 dashed rgb(190,198,237); border-top: 1 none rgb(190,198,237); border-bottom: 1 dashed rgb(190,198,237); padding-left: 10; padding-right: 10; padding-top: 5; 
        padding-bottom: 5" bgcolor="#FFFFFF"><a href="content.asp?code=<%=rs("g_code")%>"><img src="../images/zoom.gif" border="0" WIDTH="30" HEIGHT="18"></a> </td>
      </tr>
      <tr>
        <td height="25" width="350" style="border-left: 1 dashed rgb(190,198,237); border-right: 1 dashed rgb(190,198,237); border-top: 1 none rgb(190,198,237); border-bottom: 1 dashed rgb(190,198,237); padding-left: 10; padding-right: 10; padding-top: 5; padding-bottom: 5" bgcolor="#F0F2FB"><font size="2" color="navy"><%=rs("g_name")%>&nbsp; </font><font face="Comic Sans MS" size="2">(</font><font size="2">제조회사 </font><font face="Comic Sans MS" size="2">: </font><font size="2"><%=rs("g_maker")%> </font><font face="Comic Sans MS" size="2">)<br>
        <br>
        </font><font size="2">판매가 </font><font face="Comic Sans MS" size="2">:</font><font face="돋움" size="2"> <%=formatcurrency(rs("g_sellprice"))%> </font><br>
        <font size="2">데이터 마지막 수정일 </font><font face="Comic Sans MS" size="2">: </font><font size="2"><%=rs("g_updateday")%></font></td>
      </tr>
<%
rs.MoveNext
loop
%>
<!-- 끝  //-->
    </table>
    </center></div><p><br>
    </p>
    <p><br>
    </td>
  </tr>
  <tr>
    <td height="40"></td>
  </tr>
</table>
</div><% end if%>

</body>
</html>
