<%
 Set db = Server.CreateObject("ADODB.Connection")
 db.Open("dsn=toys;uid=sa;pwd=;")
	
 sql = "select C.c_address,C.c_tel,C.c_sex,C.c_name,B.b_totalprice,B.b_date,B.b_bank,B.b_summary from buy B,customer C where B.c_code = C.c_code order by b_date desc"
	
 Set rs = Server.CreateObject("ADODB.RecordSet")	
 'set rs = db.Execute(sql)
 rs.open sql, db  
 %>
<html>

<head>
<title>구매 내역</title>
</head>

<body bgcolor="#ffffff">

<table border="0" width="650" cellpadding="0" cellspacing="1">
  <tr>
    <td width="100%" height="40"><p align="center"><strong>구매결과 리스트</strong></td>
  </tr>
<% 
Do until rs.EOF
summary = replace(rs("b_summary"),chr(13) & chr(10),"<br>")
%>
  <tr>
    <td height="40">　<div align="center"><center><table border="1" cellpadding="0"
    cellspacing="0" width="600" height="40">
      <tr>
        <td width="166" height="30" align="center" bgcolor="#002142"><font face="돋움" size="2"
        color="#FFFFFF">구매자</font></td>
        <td width="249" height="30" align="center" bgcolor="#002142"><font face="돋움" size="2"
        color="#FFFFFF">거래내역</font></td>
        <td width="92" height="30" align="center" bgcolor="#002142"><font face="돋움" size="2"
        color="#FFFFFF">거래액</font></td>
        <td width="83" height="30" align="center" bgcolor="#002142"><font face="돋움" size="2"
        color="#FFFFFF">입금은행</font></td>
      </tr>
      <tr>
        <td width="166" height="30" align="center"><font face="돋움" size="2"><%=rs("c_name")%></font></td>
        <td width="249" height="30" align="left" style="text-align: left; padding-left: 10"><font
        face="돋움" size="2"><%=summary%></font></td>
        <td width="92" height="30" align="center" style="text-align: right; padding-right: 10"><font
        face="돋움" size="2"><%=formatcurrency(rs("b_totalprice"))%></font></td>
        <td width="83" height="30" align="center"><font face="돋움" size="2"><%=rs("b_bank")%></font></td>
      </tr>
      <tr>
        <td width="166" height="30" align="center" bgcolor="#002142"><font face="돋움" size="2"
        color="#FFFFFF">구매일</font></td>
        <td width="249" height="30" align="center" style="text-align: center; padding-left: 10"
        bgcolor="#002142"><font face="돋움" size="2" color="#FFFFFF">배달처</font></td>
        <td width="92" height="30" align="center" style="text-align: right; padding-right: 10"
        bgcolor="#002142"><font face="돋움" size="2" color="#FFFFFF">연락처</font></td>
        <td width="83" height="30" align="center" bgcolor="#002142"><font face="돋움" size="2"
        color="#FFFFFF">성별</font></td>
      </tr>
      <tr>
        <td width="166" height="30" align="center"><font face="돋움" size="2"><%=rs("b_date")%></font></td>
        <td width="249" height="30" align="left" style="text-align: center; padding-left: 10"><font
        face="돋움" size="2"><%=rs("c_address")%></font></td>
        <td width="92" height="30" align="center" style="text-align: center; padding-right: 10"><font
        face="돋움" size="2"><%=rs("c_tel")%></font></td>
        <td width="83" height="30" align="center"><font face="돋움" size="2"><%=rs("c_sex")%></font></td>
      </tr>
<% 
rs.movenext
loop      
%>
    </table>
    </center></div></td>
  </tr>
</table>
</body>
</html>
