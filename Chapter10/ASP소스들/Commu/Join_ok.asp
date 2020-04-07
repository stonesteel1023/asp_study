<% Option Explicit %>
<!--#include file="../inc/connection.asp"-->
<%
Dim code, BoardName, id, sql
Dim Dbcon

code = request("code")
BoardName = request("BoardName")
id = Request.Cookies("user")("id")

Set Dbcon= Server.CreateObject("ADODB.Connection")
Dbcon.Open strConnect

sql = "insert into CircleMember (Cc_BoardName, mem_id) values "
sql = sql & "('" & BoardName & "','" & id & "')"

Dbcon.Execute(sql)

sql = "Update Circle set Cc_memcount = Cc_memcount + 1"
sql = sql & " where Cc_code = " & code

Dbcon.Execute(sql)

Response.Redirect "main.asp?code=" & code & "&BoardName=" & BoardName
%>	