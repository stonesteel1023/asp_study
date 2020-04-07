<html>

<head>
<title>회원가입</title>
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
      <td width="600" align="center">　<p><font face="돋움" size="5" color="#00264D"><strong>어떻게 
      배달해 드릴까요?</strong></font></p>
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
              <td width="100" height="25" bgcolor="#E3F2F9" style="padding-left: 10"><font face="돋움"
              size="2" color="#002C57">이름</font></td>
              <td width="370" height="25" style="border: 1 dashed; padding-left: 10"><input TYPE="text"
              size="10" name="name" maxlength="8"><font face="돋움" size="2" color="#002C57"> </font></td>
            </tr>
            <tr>
              <td align="center" height="25" width="30"></td>
              <td height="25" bgcolor="#E3F2F9" style="padding-left: 10" width="100"><font face="돋움"
              size="2" color="#002C57">배달처</font></td>
              <td height="25" style="border: 1 dashed; padding-left: 10" width="370"><input TYPE="text"
              size="40" name="address"></td>
            </tr>
            <tr>
              <td align="center" height="25" width="30"></td>
              <td height="25" bgcolor="#E3F2F9" style="padding-left: 10" width="100"><font face="돋움"
              size="2" color="#002C57">연락처</font></td>
              <td height="25" style="border: 1 dashed; padding-left: 10" width="370"><input TYPE="text"
              size="10" name="tel"></td>
            </tr>
            <tr align="center">
              <td height="18" colspan="3"></td>
            </tr>
            <tr align="center">
              <td height="21" width="30"></td>
              <td height="25" style="padding-left: 10" bgcolor="#E3F2F9" width="100"><div align="left"><p><font
              face="돋움" size="2" color="#002C57">성 별</font></td>
              <td height="25" style="border: 1 dashed; padding-left: 10" align="left" width="370"><div
              align="left"><p><font face="돋움" size="2" color="#002C57"><input TYPE="radio"
              name="sex" value="남" checked>남자 <input TYPE="radio" name="sex" value="여">여자</font></td>
            </tr>
            <tr align="center">
              <td height="26" width="30"></td>
              <td height="25" style="padding-left: 10" bgcolor="#E3F2F9" width="100"><div align="left"><p><font
              face="돋움" size="2" color="#002C57">은 행</font></td>
              <td height="25" style="border: 1 dashed; padding-left: 10" width="370"><div align="left"><p><font
              face="돋움" size="2" color="#002C57"><select name="bank" size="1"
              style="font-family: 돋움; font-size: 12; background-color: rgb(255,255,255); color: rgb(0,0,0)">
                <option value="서울은행">서울은행 : 20-1234-123-3123</option>
                <option value="국민은행">국민은행 : 212-21-1321-1212</option>
              </select></font></td>
            </tr>
          </table>
          <div align="center"><center><p><input type="button" value="구입합니다." name="B2"
          OnClick="sendit();">&nbsp; <input type="button" value="취소(이전으로)" name="B1"
          OnClick="history.back();"></td>
        </tr>
      </table>
      </td>
    </tr>
  </table>
</form>
</body>
</html>
