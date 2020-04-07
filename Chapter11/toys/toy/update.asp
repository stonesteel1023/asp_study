<%
 num = request("code").count
 
 Set db = Server.CreateObject("ADODB.Connection")
 db.Open("dsn=toys;uid=sa;pwd=;")

 for i=1 to num
	
 sql = "Update imsi_buy set "
 sql = sql & "imsi_ea = " & request("ea")(i)
 sql = sql & " where imsi_memid='" & session.SessionID & "'"
 sql = sql & " and imsi_goodscode='" & request("code")(i) & "'"
 
 db.Execute sql
 
 next
 
 Response.Redirect "cart.asp"
' Set db = Server.CreateObject("ADODB.Connection")
' db.Open("toys")
	
' sql = "select * from goods where g_code='" & code & "'"
	
' set rs = db.Execute(sql)
%>
