<%
Option Explicit
Dim id

id = request("id")

if id <> "" then
	Response.Cookies("user").expires = #12/30/2010 00:00:00#
	Response.Cookies("user")("id") = id
	Response.redirect "Queue/Queue.asp"
else
%>
<html><body>
<form Method="post" action="default.asp" id=form1 name=form1>
인증 아이디의 입력 <input name="id">
<input type="submit" value="Connect" id=submit1 name=submit1>
</form>
</body></html>
<% end if%>
