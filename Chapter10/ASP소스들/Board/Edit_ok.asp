<!--#include file="../inc/connection.asp"-->
<%
Response.Buffer = false
GoTopage= request("GotoPage")
table = request("table")

Set db = Server.CreateObject("ADODB.Connection")
db.Open strconnect

SQL = "select pwd from " & table & " where board_idx = " & request("gant")

set result = db.execute(SQL)


if request("pwd") = result(0) then
		
	name = Request.Form("name")
	mail = Request.Form("mail")
	title = Request.Form("title")

	if title = "" then
		title = "�������"
	end if

	memo = Request.Form("memo")
	memo = replace(memo,"|","l")

	if request("tag")="ON" then
		info = "tagok"
	end if

	URL = Request.Form("URL")

	SQL = "Update " & table & " set name = '" & name & "'"
	SQL = SQL & ", title = '" & title & "'"
	SQL = SQL & ", mail = '" & mail & "'"
	SQL = SQL & ", content = '" & memo & "'"
	SQL = SQL & ", url = '" & URL & "'"
	SQL = SQL & ", writeday = getdate()"'" & now() & "'"
	SQL = SQL & ", part ='" & info & "'"
	SQL = SQL & " where board_idx = " & Clng(request("gant"))

    'Response.Write sql
    db.execute SQL
	    
	Response.Redirect "list.asp?gotopage=" & gotopage & "&table=" & table

else %>
<html>
<head>
<title>Document Title</title>
</head>

<body>
<p align="center">��</p>
<p align="center">��</p>
<p align="center">��</p>
<p align="center"><big><strong><font color="#FF8040"><img src="images/noway2.gif" WIDTH="32" HEIGHT="32"> ��й�ȣ</font>�� <font color="#000040">��ġ���� �ʽ��ϴ�</font>.</strong></big></p>
<p align="center">��</p>
<p align="center"><font face="����">�ڼ��� ������ <strong><font color="#004080">���</font></strong>���� 
���ǹٶ��ϴ�.</font></p>
<p align="center"><strong><font face="����"><font size="2" color="#808080">�Ʒ��� 
��ư�� ������</font><font color="#800080" size="2">&nbsp; </font></font><font size="2" color="#408080" face="����">��������</font><font face="����"><font color="#0080C0" size="2"> </font><font color="#808080"><font size="2"> �̵��մϴ�.</font>.</font></font></strong></p>
<p align="center"><a href="javascript:history.back();"><img src="images/top.gif" alt="�������� �̵�" border="0" width="44" height="41"></a></p>
<% end if %>
<% 
   db.Close
   Set db = Nothing
%>
</body>
</html>
