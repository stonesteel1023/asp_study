<!--#include file="../inc/connection.asp"-->
<%
table = request("table")


    Set db = Server.CreateObject("ADODB.Connection")
	db.Open strconnect

    SQL = "select board_idx from " & table & " where  board_idx = " & request("dog") 	
    SQL = SQL & " and pwd ='" & request("pwd") & "'"
    
        
    Set rs = Server.CreateObject("ADODB.RecordSet")
    rs.Open SQL,db
    
	if rs.EOF or rs.BOF then
%>
<script language="javascript">
	alert("��й�ȣ�� ��ġ���� �ʽ��ϴ�.");
	history.back();
</script>
<%
      Response.End
    else
     SQL = " DELETE FROM " & table & " where board_idx = " & request("dog") 	
	 db.execute SQL
     Response.Redirect "list.asp?table=" & table
    end if

%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual InterDev 1.0">
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=ks_c_5601-1987">
<title>taeyo's Board</title>
</head>
<body>

<p align="center">��</p>

<p align="center">��</p>

<p align="center">��</p>

<p align="center"><big><strong><big><big>���� <font color="#0000FF">����</font>�Ǿ����ϴ�.</big></big></strong></big></p>

<p align="center">��</p>

<p align="center"><strong><font size="3" color="#808080">�Ʒ��� ��ư�� ����</font><font color="#800080" size="3">&nbsp; </font><font size="3" color="#FF8040">����Ʈ ����</font><font color="#0080C0" size="3"> </font><font color="#808080"><font size="3">�� ���ż� 
Ȯ���Ͻʽÿ�</font>.</font></strong></p>

<p align="center"><a href="List.asp"><img src="../../images/top.gif" alt="����Ʈ ����� �̵�" border="0" width="44" height="41"></a></p>

</body>
</html>
<% 
   rs.Close
   db.Close
   Set rs = Nothing
   Set db = Nothing
%>