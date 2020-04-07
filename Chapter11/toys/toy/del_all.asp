<%
 
 Set db = Server.CreateObject("ADODB.Connection")
 db.Open("dsn=toys;uid=sa;pwd=;")
	
 sql = "delete from imsi_buy where imsi_memid ='" & session.SessionID & "'"
 
 'Response.Write sql
 db.Execute sql
 %>
