<%
code = request("code")
ea = request("ea")

'buy_temp = code & "/" & sellprice & "/" & ea

 Set db = Server.CreateObject("ADODB.Connection")
 db.Open ("dsn=toys;uid=sa;pwd=;")

 sql = "select * from imsi_buy where imsi_goodscode ='" & code & "'"
 sql = sql & " and imsi_memid='" & session.SessionID & "'"
 
 set rs = db.Execute(sql)
 
 if rs.EOF or rs.BOF then
	
   sql = "insert into imsi_buy (imsi_memid,imsi_goodscode,imsi_ea) values "
   sql = sql & "('" & session.SessionID & "'"
   sql = sql & ",'" & code & "'"
   sql = sql & "," & ea & ")"
	
   db.Execute sql
  
 else
   
   sql = "Update imsi_buy set "
   sql = sql & " imsi_ea= imsi_ea +" & ea 
   sql = sql & " where imsi_goodscode ='" & code & "'"
   
   db.Execute sql
   
 end if
 
 Response.Redirect "cart.asp"
%>
