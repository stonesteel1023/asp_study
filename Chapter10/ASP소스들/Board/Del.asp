<!--#include file="../inc/connection.asp"-->
<%
GoTopage= request("GotoPage")
table = request("table")

%>
<html>

<head>
<link rel="stylesheet" href="../css/main.css">
</head>
<script language="javascript">
<!--
function sendit() {
       
          //��й�ȣ
        if (document.forms[0].pwd.value == "" ) {
                alert("��й�ȣ�� ������ �ֽʽÿ�");
                return false;
        } 
		document.forms[0].submit();
        
	}
//-->
</script>


<body bgcolor="#ffffff" link="#000080" vlink="#000080" alink="#000080" OnLoad="document.delform.pwd.focus();">

<br>
</p>

<form method="POST" name="delform"  action="del_ok.asp">
<input type="hidden" name="table" value="<%=table%>">
  <input type="hidden" name="dog" value="<%=request("dog")%>">
  <input type="hidden" name="gotopage" value="<%=gotopage%>">
  <p>��</p>
  <div align="center"><center><table border="0" cellspacing="0" width="450">
    <tr>
      <td bgcolor="#000080"><div align="center"><center><p><font color="#FFFF00"><big><big><big><strong>Delete</strong></big></big></big></font></td>
    </tr>
    <tr align="center">
      <td bgcolor="#F8FCDC" align="center">��<div align="center"><center><p><font face="����"
      size="2">���� : <font color="#004080"><%=request("name")%></font>����&nbsp; ���� <font color="#004080"><%=request("title")%></font></font></p>
      </center></div><div align="center"><center><p><font face="����" size="2">���� 
      ��й�ȣ�� ������ �ֽʽÿ�...</font></p>
      </center></div><div align="center"><center><p><font face="����" size="2" color="#000000">(÷�ε�</font><font
      face="����" size="2" color="#008080"> <strong>ȭ��</strong></font><font face="����"
      size="2" color="#000000">�� ������� ����</font><font face="����" size="2"
      color="#008080"> </font><strong><font face="����" size="2" color="#FF8040">����</font></strong><font
      face="����" size="2" color="#008080"> </font><font face="����" size="2"
      color="#000000">�˴ϴ�)</font></p>
      </center></div><div align="center"><center><p><br>
      </td>
    </tr>
    <tr align="center">
      <td bgcolor="#E1E1E1"><div align="center"><center><p><br>
      <font face="����" size="2">��й�ȣ :&nbsp; <input type="password" name="pwd"
      size="8"></font><br>
      <br>
      </td>
    </tr>
  </table>
  </center></div><div align="center"><center><p>��</p>
  </center></div><div align="center"><center><p><input type="button" value=" �� �� "
  OnClick="sendit()" name="del" style="color:black;background-color:#ffcc00;border:1 solid black;height:20">&nbsp;&nbsp; <font face="����" size="2">&lt;<a
  href="javascript:history.back();">���� ���ҷ���!!</a>&gt;</font></p>
  </center></div>
</form>

<p align="center">��</p>
</body>
</html>
