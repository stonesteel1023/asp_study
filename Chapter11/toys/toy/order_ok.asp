<%'@TRANSACTION="Required"%>
<!--#include file="../pkval.asp"//-->
<%
 sum = request("sum")
 name = request("name")
 address = request("address")
 sex = request("sex")
 tel = request("tel")
 bank = request("bank")
 dd = now()
  
 Set db = Server.CreateObject("ADODB.Connection")
 db.Open("dsn=toys;uid=sa;pwd=;")
	
 sql = "select imsi_ea, g_code, g_name, g_sellprice from imsi_buy A, goods B  where A.imsi_memid ='" & session.SessionID & "'"
 sql = sql & " and A.imsi_goodscode = B.g_code"
 
 Set rs = Server.CreateObject("ADODB.RecordSet")
 rs.Open sql, db
 
 Do until rs.EOF
   price = rs("g_sellprice")
   ea = rs("imsi_ea")
   code = rs("g_code")
   g_name = rs("g_name")
   
   '구입물품 요약정리
   str = g_name & "(" & code & ") : " & ea & "개" & chr(13) & chr(10)
   summary = summary & str
   
   rs.MoveNext   
 Loop 
 
 rs.Close
 
 '고객의 정보 입력, 배달처등..
 sql = "insert into customer (c_code,c_name,c_tel,c_sex,c_address) values "
 sql = sql & "('" & primaryval & name & "'"
 sql = sql & ",'" & name & "'"
 sql = sql & ",'" & tel & "'"
 sql = sql & ",'" & sex & "'"
 sql = sql & ",'" & address & "')"
 db.Execute sql
 
'그리고 난 뒤에 실제구매 테이블에 
 sql = "insert into buy (b_code,c_code,b_date, b_summary,b_totalprice,b_bank) values "
 sql = sql & "('" & primaryval & name & "'"
 sql = sql & ",'" & primaryval & name & "'"
 sql = sql & ",'" & dd & "'" 
 sql = sql & ",'" & summary & "'"
 sql = sql & "," & sum 
 sql = sql & ",'" & bank & "')" 
 db.Execute sql
 

 '임시 테이블의 내역을 삭제
 sql = "delete from imsi_buy where imsi_memid='" & session.SessionID & "'"
 
 db.Execute sql
 
Sub OnTransactionCommit()   'Commit시 하고 싶은 처리..

 Response.Redirect "result.htm"

End Sub

Sub OnTransactionAbort()   'Commit시 하고 싶은 처리..

    response.write  "입력에 문제발생... 다시 입력하세요"
%> 
<script language="javascript">
alert("오류가 발생해서 모든 작업을 취소했습니다.\n\n다시 시도해 주십시요\n\n다시 문제가 생기면 관리자에게 연락을 주시기 바랍니다");
history.back();
</script>
<%
End Sub
%>