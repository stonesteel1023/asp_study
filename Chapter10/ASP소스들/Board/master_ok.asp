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

<p align="center"><font face="돋움" size="2" color="#004080">로긴이 성공적으로 
수행되었습니다..</font></p>

<p align="center"><font face="돋움" size="2" color="#004080">관리자님은 
게시판에서의 글들의 내용을 보실때 </font></p>

<p align="center"><font face="돋움" size="2" color="#004080">그 글들의 
비밀번호가 같이 보여질 것입니다..</font></p>

<p align="center"><font face="돋움" size="2" color="#004080">우선은 그것으로 
수정,삭제 하시기 바랍니다..</font></p>

<p align="center"><font face="돋움" size="2">　</font></p>

<p align="center"><font face="돋움" size="2">&lt; <a href="List.asp?table=<%=table%>"><font color="#008080">관리자의 신분으로 게시판으로 갑니다</font></a><font color="#808080">.</font> &gt;</font></p>

<p align="center"><font face="돋움" size="2">&lt; <font color="#408080"><a href="master_logoff.asp?table=<%=table%>">일반인의 신분으로 게시판으로 갑니다</a></font><font color="#808080">.</font> &gt;</font></p>

<% else %>
<p align="center"><font face="Century Gothic" size="6" color="#FF0000"><strong>Login 
Failed</strong></font></p>

<p align="center"><font face="돋움" size="2" color="#004080">　</font></p>

<p align="center"><font face="돋움" size="2" color="#004080">로긴이 
&nbsp; 실패되었습니다..</font></p>

<p align="center"><font face="돋움" size="2" color="#004080">일반 사용자님들은 
관리자 메뉴를 사용하실 수 없습니다.</font></p>

<p align="center"><font face="돋움" size="2" color="#004080">자세한 사항은 
운영자에게 문의 하십시요...</font></p>

<p align="center">　</p>

<p align="center"><font face="돋움" size="2">&lt; <font color="#008080"><a href="List.asp?table=<%=table%>">게시판으로 돌아갑니다</a>.</font> &gt;</font></p>
<% end if %>
</body>
</html>
