<% Option Explicit %>
<!--#include file="../inc/connection.asp"-->
<%
Dim code, BoardName, id, sql
Dim Dbcon, Rs

code = request("code")
BoardName = request("BoardName")
id = Request.Cookies("user")("id")

Set Dbcon= Server.CreateObject("ADODB.Connection")
Dbcon.Open strConnect

sql = "select mem_id from CircleMember where Cc_BoardName='" & BoardName & "'"
sql = sql & " and mem_id='" & id & "'"

set Rs = Dbcon.Execute(sql)

if Rs.EOF then
%>
	<script language="javascript">
	var bool = confirm("이 커뮤니티에 가입되어 있지 않습니다.\n\n지금 가입하시겠습니까?");
	if (bool)
		location.href = "Join_ok.asp?code=<%=code%>&BoardName=<%=BoardName%>";
	else
		history.back();
	</script>
<% 
	Response.End
else
	Response.Redirect "main.asp?code=" & code & "&BoardName=" & BoardName
end if
%>	