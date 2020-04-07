<%
 code = request("code")
 
 Set db = Server.CreateObject("ADODB.Connection")
 db.Open("dsn=toys;uid=sa;pwd=;")
	
 sql = "delete * from imsi_buy where imsi_memid ='" & session.SessionID & "'"
 sql = sql & " and imsi_goodscode ='" & code & "'"
 
 db.Execute sql
 
 Response.Redirect "cart.asp"
 %>
	