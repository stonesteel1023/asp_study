<!--#include file="../inc/connection.asp"-->
<%
table = request("table")
GoTopage= request("GotoPage")
id = Request.Cookies("user")("id")

Set Dbcon = Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.RecordSet")
Dbcon.Open strconnect


b_id = request("id")
if b_id <> "" then '�� �亯������

myref = CDbl(request("ref"))
mystep = CDbl(request("step"))
mylevel = CDbl(request("level"))

'********************  ������ ������ �߰��κ� ����***********************
SQL = " SELECT * FROM " & table 
SQL = SQL & " WHERE  ref = " & myref
SQL = SQL & " AND step > " & mystep
SQL = SQL & " AND re_level <= " & mylevel & " ORDER BY step "

Rs.Open SQL, Dbcon
If Not Rs.EOF then
 Bro_Step = Rs("step")
End if
Rs.Close

If Bro_Step <> "" then
	'STR = "�θ�� ���� ������ ���� �ִ�"
	SQL = " SELECT * FROM " & table
	SQL = SQL & " WHERE ref = " & myref
	SQL = SQL & " AND step > " & myStep
	SQL = SQL & " AND step < " & Bro_Step
	SQL = SQL & " AND re_level > " & mylevel & " ORDER BY step DESC"
Else
	'STR = "�θ�� ���� ������ ���� ����"
	SQL = " SELECT * FROM " & table
	SQL = SQL & " WHERE  ref = " & myref
	SQL = SQL & " AND step > " & myStep
	SQL = SQL & " AND re_level > " & mylevel
	SQL = SQL & " ORDER BY step DESC"
End if
  
Rs.Open SQL, Dbcon
If Not Rs.EOF then
  myStep = Rs("step")
End if
Rs.Close
Set Rs = Nothing
'********************  ������ ������ �߰��κ� �� ***********************

Application.Lock
SQLString = "UPDATE " & table & " SET step = step + 1 WHERE ref=" & myref & " AND step > " & mystep
Dbcon.Execute(SQLString)

newstep = mystep + 1
newlevel = mylevel + 1

else ' ÷ �۾�����...

myref = 0
newstep=0
newlevel=0

end if


homep = request("homepage")
title = request("title")
title = replace(title,"'","''")

if title = "" then
	title = "�������"
end if

P_name = request("name")
pwd = request("pwd")
mail = request("mail")
memo = request("memo")

If memo <> "" then
	memo = replace(memo,"'","''")
end if

if request("tag")="ON" then
	info = "tagok"
end if


 	SQL = "INSERT INTO " & table & " (name,title,ref,step,re_level,mail,content,readnum,pwd,url,part,writeday) VALUES "
	SQL = SQL & "('" & id & "'"
	SQL = SQL & ",'" & title & "'"
	SQL = SQL & "," & myref & " "
	SQL = SQL & "," & newstep & " "
	SQL = SQL & "," & newlevel & " "
	SQL = SQL & ",'" & mail & "'"
	SQL = SQL & ",'" & memo & "'"
    SQL = SQL & "," & 0 & " "
	SQL = SQL & ",'" & pwd & "'"
	SQL = SQL & ",'" & homep & "'"
	SQL = SQL & ",'" & info & "'"
    SQL = SQL & ",'" & now() & "')"

    Dbcon.Execute SQL

	if b_id = "" then '÷ ���̶�� ref�� �ڽ��� identity�� �����ش�
		set rs = Dbcon.execute("select @@identity")
		newref = rs(0)
		rs.close
		Set rs = Nothing
	    
		sqlup = "update " & table & " set ref =" & newref & " where board_idx =" & newref
		Dbcon.Execute SQLup
	end if
    
	Dbcon.close
	Set Dbcon = Nothing


	Response.Redirect "list.asp?table=" & table & "&Gotopage=" & gotopage

%>
