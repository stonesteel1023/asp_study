<% Option Explicit %>
<!--#include file="../inc/connection.asp"-->
<%
Dim id, code, sql
Dim Dbcon
id = Request.Cookies("user")("id") & ","
code = Request.Form("code")

Set Dbcon= Server.CreateObject("ADODB.Connection")
Dbcon.Open strConnect

SQL = "update Queue set Que_memcount = Que_memcount + 1 "
SQL = SQL & " , Que_member = Que_member + '" & id & "'"
SQL = SQL & " where Que_code=" & code

Dbcon.Execute sql
Dbcon.Close
Set Dbcon = nothing

Response.redirect "Queue.asp"
%>
