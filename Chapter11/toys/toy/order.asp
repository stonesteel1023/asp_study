<html>

<head>
<title>ȸ������</title>
<script language="JavaScript">
<!-- 
function sendit()
{
   document.orderform.submit();
}
//-->
</script>
</head>

<body bgcolor="#ffffff" MARGINHEIGHT="0" MARGINWIDTH="0" TOPMARGIN="0" LEFTMARGIN="0">

<form name="orderform" action="order_ok.asp" method="post">
  <table cellpadding="0" cellspacing="0" width="600" border="0">
    <tr>
      <td width="600" align="center">��<p><font face="����" size="5" color="#00264D"><strong>��� 
      ����� �帱���?</strong></font></p>
     </td>
    </tr>
    <tr>
      <td width="600"><table align="left" width="600" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td align="center"><table width="500" cellpadding="1" cellspacing="0" border="0"
          bgcolor="#FFFFFF">
            <tr>
              <td height="16" colspan="3"></td>
            </tr>
            <tr>
            <input type="hidden" name="sum" value="<%=request("sum")%>">
              <td width="30" align="center" height="25"></td>
              <td width="100" height="25" bgcolor="#E3F2F9" style="padding-left: 10"><font face="����"
              size="2" color="#002C57">�̸�</font></td>
              <td width="370" height="25" style="border: 1 dashed; padding-left: 10"><input TYPE="text"
              size="10" name="name" maxlength="8"><font face="����" size="2" color="#002C57"> </font></td>
            </tr>
            <tr>
              <td align="center" height="25" width="30"></td>
              <td height="25" bgcolor="#E3F2F9" style="padding-left: 10" width="100"><font face="����"
              size="2" color="#002C57">���ó</font></td>
              <td height="25" style="border: 1 dashed; padding-left: 10" width="370"><input TYPE="text"
              size="40" name="address"></td>
            </tr>
            <tr>
              <td align="center" height="25" width="30"></td>
              <td height="25" bgcolor="#E3F2F9" style="padding-left: 10" width="100"><font face="����"
              size="2" color="#002C57">����ó</font></td>
              <td height="25" style="border: 1 dashed; padding-left: 10" width="370"><input TYPE="text"
              size="10" name="tel"></td>
            </tr>
            <tr align="center">
              <td height="18" colspan="3"></td>
            </tr>
            <tr align="center">
              <td height="21" width="30"></td>
              <td height="25" style="padding-left: 10" bgcolor="#E3F2F9" width="100"><div align="left"><p><font
              face="����" size="2" color="#002C57">�� ��</font></td>
              <td height="25" style="border: 1 dashed; padding-left: 10" align="left" width="370"><div
              align="left"><p><font face="����" size="2" color="#002C57"><input TYPE="radio"
              name="sex" value="��" checked>���� <input TYPE="radio" name="sex" value="��">����</font></td>
            </tr>
            <tr align="center">
              <td height="26" width="30"></td>
              <td height="25" style="padding-left: 10" bgcolor="#E3F2F9" width="100"><div align="left"><p><font
              face="����" size="2" color="#002C57">�� ��</font></td>
              <td height="25" style="border: 1 dashed; padding-left: 10" width="370"><div align="left"><p><font
              face="����" size="2" color="#002C57"><select name="bank" size="1"
              style="font-family: ����; font-size: 12; background-color: rgb(255,255,255); color: rgb(0,0,0)">
                <option value="��������">�������� : 20-1234-123-3123</option>
                <option value="��������">�������� : 212-21-1321-1212</option>
              </select></font></td>
            </tr>
          </table>
          <div align="center"><center><p><input type="button" value="�����մϴ�." name="B2"
          OnClick="sendit();">&nbsp; <input type="button" value="���(��������)" name="B1"
          OnClick="history.back();"></td>
        </tr>
      </table>
      </td>
    </tr>
  </table>
</form>
</body>
</html>
