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
   
   '���Թ�ǰ �������
   str = g_name & "(" & code & ") : " & ea & "��" & chr(13) & chr(10)
   summary = summary & str
   
   rs.MoveNext   
 Loop 
 
 rs.Close
 
 '���� ���� �Է�, ���ó��..
 sql = "insert into customer (c_code,c_name,c_tel,c_sex,c_address) values "
 sql = sql & "('" & primaryval & name & "'"
 sql = sql & ",'" & name & "'"
 sql = sql & ",'" & tel & "'"
 sql = sql & ",'" & sex & "'"
 sql = sql & ",'" & address & "')"
 db.Execute sql
 
'�׸��� �� �ڿ� �������� ���̺� 
 sql = "insert into buy (b_code,c_code,b_date, b_summary,b_totalprice,b_bank) values "
 sql = sql & "('" & primaryval & name & "'"
 sql = sql & ",'" & primaryval & name & "'"
 sql = sql & ",'" & dd & "'" 
 sql = sql & ",'" & summary & "'"
 sql = sql & "," & sum 
 sql = sql & ",'" & bank & "')" 
 db.Execute sql
 

 '�ӽ� ���̺��� ������ ����
 sql = "delete from imsi_buy where imsi_memid='" & session.SessionID & "'"
 
 db.Execute sql
 
Sub OnTransactionCommit()   'Commit�� �ϰ� ���� ó��..

 Response.Redirect "result.htm"

End Sub

Sub OnTransactionAbort()   'Commit�� �ϰ� ���� ó��..

    response.write  "�Է¿� �����߻�... �ٽ� �Է��ϼ���"
%> 
<script language="javascript">
alert("������ �߻��ؼ� ��� �۾��� ����߽��ϴ�.\n\n�ٽ� �õ��� �ֽʽÿ�\n\n�ٽ� ������ ����� �����ڿ��� ������ �ֽñ� �ٶ��ϴ�");
history.back();
</script>
<%
End Sub
%>