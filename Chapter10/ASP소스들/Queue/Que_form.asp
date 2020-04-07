<HTML>
<HEAD><link rel="stylesheet" href="../css/main.css"></HEAD>
<BODY>
<div align="center">
<form name="queue" Method="post" action="Que_make.asp">
<table width="400" border=1 bgcolor=#ececc9 cellspacing=0>
<tr bgcolor="#2f4f5c">
  <td colspan=2 width=400 height=30 align="center">
  <font class="blue">커뮤니티 신청</font></td></tr>
<tr>
  <td width=100 height=30 align="center">신청자</td>
  <td width=300 height=30 align="center">
  <input type="text"  name="ID" size="35" readonly value="<%=Request.Cookies("user")("id")%>">
  </td></tr>
<tr>
  <td width=100 height=30 align="center">모임주제</td>
  <td width=300 height=30 align="center">
  <input type="text"  name="title" size="35"></td></tr>
<tr>
  <td colspan=2 height=30 align="center">모임주제 설명</td></tr>
<tr>  
  <td colspan=2 height=30 align="center">
  <textarea name="content" rows="10" cols="53" wrap="hard"
  style="font-family:돋움;font-size:10pt;border-width:1;border-style:solid"></textarea>
  </td></tr>
<tr>
  <td colspan=2 align="center">
  <input type="submit" value="커뮤니티 신청" style="background-color:#efeff4">
  <input type="button" value="뒤로 돌아가기" style="background-color:#efeff4"
  OnClick="javascript:history.back();"></td></tr>  
</table>
</form>
</div>
</BODY>
</HTML>
