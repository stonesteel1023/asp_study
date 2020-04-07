<% Option Explicit %>
<!--#include file="../inc/connection.asp"-->
<%
Dim id, title, content, member, SQL
Dim Dbcon

id = Request.Form("id")
title = Request.Form("title")
content = Request.Form("content")

member = id & ","
title = replace(title,"'","''")
content = replace(content,"'","''")

Set Dbcon= Server.CreateObject("ADODB.Connection")
Dbcon.Open strConnect

SQL = "Insert into Queue (Que_boss, Que_title, Que_member, Que_description) values "
SQL = SQL & "('" & id & "'"
SQL = SQL & ",'" & title & "'"
SQL = SQL & ",'" & member & "'"
SQL = SQL & ",'" & content & "')"

Dbcon.Execute sql
Dbcon.Close
Set Dbcon = nothing

Response.redirect "Queue.asp"
%>
