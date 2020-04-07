<%
 'request객체를 통해 넘어온 값들을 변수에 저장한다. (추천되는 방법이다)
 code = request("code")
 part = request("part")
 name = request("name")
 maker = request("maker")
 origin_price = request("origin_price")
 sellprice = request("sellprice")
 image = request("image")
 content = request("content")
  
  
 Set db = Server.CreateObject("ADODB.Connection")
 db.Open("dsn=toys;uid=sa;pwd=;")

 SQL = "INSERT INTO goods (g_code,g_part,g_name,g_maker,g_origin_price,"
 SQL = SQL & " g_sellprice,g_image,g_updateday,g_content) VALUES "
 SQL = SQL & "('" & code & "'"
 SQL = SQL & ",'" & part & "'" 
 SQL = SQL & ",'" & name & "'" 
 SQL = SQL & ",'" & maker & "'" 
 SQL = SQL & "," & origin_price 
 SQL = SQL & "," & sellprice 
 SQL = SQL & ",'" & image & "'"
 SQL = SQL & ",'" & now() & "'"  
 SQL = SQL & ",'" & content & "')"
 
 'Response.Write sql
 db.Execute sql

 Response.redirect "new_result.htm"
%>