<!--#include file="../inc/connection.asp"-->
<%
GoTopage= request("GotoPage")
table = request("table")
%>
<html>

<head>
<title>��������</title>
<link rel="stylesheet" href="../css/main.css">
<script language="javascript">
<!--
function send_re()
{
    document.re.submit();
}
function myst()
{
  window.self.status = "÷��ȭ���� �ֽ��ϴ�.";
}
//-->
</script>
</head>
<%
   	Set db = Server.CreateObject("ADODB.Connection")
	db.Open strconnect

    UpdateSQL = "Update " & table & " Set readnum = readnum+1 where board_idx = " & request.Querystring("id")
    db.Execute UpdateSQL

	SQL = "SELECT board_idx,name,title,mail,url,writeday,pwd,ref,step,re_level,readnum, content,part from " & table  
	SQL = SQL & " where board_idx = " & request.Querystring("id") 

    Set Rs = Server.CreateObject("ADODB.Recordset")

	Rs.Open SQL,db
%>

<body bgcolor="#FFFFFF" link="#000080" vlink="#000080" alink="#000080">
<%  if Rs.BOF or Rs.EOF then 
%>
<p>����� ������ ���.. </p>
<%
else 

	con_id = Rs("board_idx")
	name = (Rs("name"))
	title = replace(Rs("title"),"&quot;","'")
	mail = Rs("mail")
	URL = Rs("url")
	writeday = Rs("writeday")
	pword = Rs("pwd")
	refer = Rs("ref")
	c_step = Rs("step")
	c_level = Rs("re_level")
	readnum = Rs("readnum")
	content = Rs("content")
	info = Rs("part")

	if info <>"tagok" then
	  content = UseText(content)
	end if
%>
<div align="center"><div align="center">

<p align="center"><font color="#004080" face="����"><big><big><strong>�Խù� ����<%=test%></strong></big></big></font></p>
<div align="center"><center>

<table border="0" cellspacing="0" width="480" style="border: medium none" cellpadding="0"
bordercolor="#C0C0C0" bordercolorlight="#C0C0C0" bordercolordark="#C0C0C0">
  <tr>
    <td align="center" width="100" height="25" bgcolor="#B6C9E4"
    style="border-bottom: 1 dashed rgb(182,201,228); padding-top: 3; padding-bottom: 3"><font
    face="����" size="2" color="#003A75"><strong>�۾���</strong></font></td>
    <td align="center" width="100" height="25"
    style="border-bottom: 1 dashed rgb(182,201,228); padding-top: 3; padding-bottom: 3"
    bgcolor="#FFFFFF"><font face="����" size="2"><%=name%></font></td>
    <td align="center" width="100" height="25" bgcolor="#B6C9E4"
    style="border-bottom: 1 dashed rgb(182,201,228); padding-top: 3; padding-bottom: 3"><font
    face="����" size="2" color="#003A75"><strong>��¥</strong></font></td>
    <td align="center" width="180" height="25"
    style="border-bottom: 1 dashed rgb(182,201,228); padding-top: 3; padding-bottom: 3"
    bgcolor="#FFFFFF"><font face="����" size="2"><%=writeday%></font></td>
  </tr>
  <tr>
    <td align="center" width="100" height="25" bgcolor="#B6C9E4"
    style="border-bottom: 1 dashed rgb(182,201,228); padding-top: 3; padding-bottom: 3"><font
    face="����" size="2" color="#003A75"><strong>E-mail</strong></font></td>
    <td align="center" width="100" height="25"
    style="border-bottom: 1 dashed rgb(182,201,228); padding-top: 3; padding-bottom: 3"
    bgcolor="#FFFFFF"><font face="����" size="2"><a href="mailto:<%=mail%>"><strong>E-mail</strong></a></font></td>
    <td align="center" width="100" height="25" bgcolor="#B6C9E4"
    style="border-bottom: 1 dashed rgb(182,201,228); padding-top: 3; padding-bottom: 3"><font
    face="����" size="2" color="#003A75"><strong>Homepage</strong></font></td>
    <td align="center" width="180" height="25"
    style="border-bottom: 1 dashed rgb(182,201,228); padding-top: 3; padding-bottom: 3"
    bgcolor="#FFFFFF"><font face="����" size="2"><a href="<%=URL%>" TARGET="_blank"><strong><%=name%>�� Homepage</strong></a></font></td>
  </tr>
  <tr>
    <td colspan="4" style="padding-top: 5; padding-bottom: 5" height="40"><p align="right"><strong><font
    face="����" size="2" color="#004080">��ȸ�� : <%=readnum%>&nbsp;&nbsp;&nbsp;</font></strong></td>
  </tr>
  <tr>
    <td colspan="4" style="padding-top: 10; padding-bottom: 10" bgcolor="#5A69A7"><p
    align="center"><font face="����" size="2" color="#ffffff"><strong>�� �� : <font
    color="#ffffff"><%=title%></font></strong></font></td>
  </tr>
  <tr>
    <td colspan="4"
    style="border-left: 1 dashed rgb(90,105,167); border-right: 1 dashed rgb(90,105,167); border-bottom: 1 dashed rgb(90,105,167); padding-left: 30; padding-right: 30; padding-top: 10"
    bgcolor="#FFFFFF"><pre><font face="����" size="2"><%=content%></font></pre>
    </td>
  </tr>
</table>
</center></div></div><% end if %>

<p><font face="����" size="2">
&lt;<a href="javascript:send_re()">�亯�ϱ�</a>&gt; 
&lt;<a href="list.asp?GoTopage=<%=GotoPage%>&table=<%=table%>">����Ʈ�� ���ư���</a>&gt; 

<% if request("found")="is" then%> 
	&lt;<a href="javascript:window.history.back(1);"> �˻������ ���ư���</a> &gt;
<%end if%> 

&lt;<a href="edit.asp?dog=<%=con_id%>&amp;GoTopage=<%=GotoPage%>&table=<%=table%>">����</a>&gt; 

&lt;<a href="del.asp?name=<%=name%>&amp;dog=<%=con_id%>&amp;GoTopage=<%=GotoPage%>&table=<%=table%>&amp;step=<%=c_step%>&amp;title=<%=title%>">����</a>&gt; 

</p>
</div>

<p>��</p>
<div align="center"><center>

<table border="0" cellspacing="0" width="550">
  <tr>
    <td><font face="����ü" color="#004080" size="2">������ �۰� ���õ� �۵��� 
    ����Ʈ����</font></td>
  </tr>
</table>
</center></div>

<form name="re" method="post" action="write.asp">
  <input type="hidden" name="table" value="<%=table%>">
  <input type="hidden" name="id" value="<%=con_id%>">
  <input type="hidden" name="head"  value="<%=title%>">
  <input type="hidden" name="step" value="<%=C_step%>">
  <input  type="hidden" name="level" value="<%=C_level%>">
  <input type="hidden" name="ref"  value="<%=refer%>">
  <input type="hidden" name="GoToPage" value="<%=gotoPage%>">
  <input  type="hidden" name="block" value="<%=block%>">
  <input  type="hidden" name="ques_mail" value="<%=mail%>">

<%
'Ÿ��Ʋ�� " ���� �ذ�κ�
title = replace(title,chr(34),"&#34")
%>

</form>
<%
   Rs.Close
   
	SQL = "SELECT * FROM " & table
	SQL = SQL & " where ref = " & refer
	SQL = SQL & " order by ref desc,step asc"

    Rs.Open SQL, db
%>
<%  if Rs.BOF or Rs.EOF then 
    else
%>
<div align="center"><div align="center"><center>

<table border="0" width="630">
  <tr>
    <td width="20" style="padding-top: 3; padding-bottom: 3"></td>
    <td width="350" bgcolor="#BCBCDE" style="padding-top: 3; padding-bottom: 3" align="center"><p
    align="center"><font face="����" size="2" color="#37376F">��&nbsp; ��</font></td>
    <td width="96" bgcolor="#BCBCDE" style="padding-top: 3; padding-bottom: 3" align="center"><p
    align="center"><font face="����" size="2" color="#37376F">�۾���</font></td>
    <td width="93" bgcolor="#BCBCDE" style="padding-top: 3; padding-bottom: 3" align="center"><p
    align="center"><font face="����" size="2" color="#37376F">��¥</font></td>
    <td width="71" bgcolor="#BCBCDE" style="padding-top: 3; padding-bottom: 3" align="center"><p
    align="center"><font face="����" size="2" color="#37376F">��ȸ</font></td>
  </tr>
<%
Do until Rs.EOF 

	name = Rs("name")
	'��¥�� �����ؼ� �����ִ� �κ�
	yymmdd = rs("writeday")
	yy= year(yymmdd)
	mm = right("0" & month(yymmdd),2)
	dd = right("0" & day(yymmdd),2)
	wdate = yy & "/" & mm & "/" & dd

	If Len(title) > 20 Then
		title = Mid(title,1,21) & "..."
	End If
%>
  <tr>
    <td width="20" style="padding-top: 2; padding-bottom: 2"><%if DateDiff("h",Rs("writeday"),Now()) < 20 then%>
<p><img src="images/dot.gif"
    width="13" height="12"><% end if %> </td>
    <td width="350" <% if Rs("board_idx") = con_id then %>bgcolor="#fffff1"<%end if%> style="padding-left: 10; padding-top: 2; padding-bottom: 2"><font
    face="����" size="2"><% if Rs("re_level") > 0 then %> <% wid = 4 * Rs("re_level") %> <img src="images/x.gif" width="<%=wid%>" height="16"> <img
    src="images/re.gif" width="26" height="16"> <% end if %> 
 <% if Rs("board_idx") = con_id then %> <font color="gray"><%=title%></font> 
 <%else%> 
    <a
    href="content.asp?id=<%=Rs("board_idx")%>&amp;GoTopage=<%=GotoPage%>&amp;table=<%=table%>"><%=title%></a> 
<%end if%>    </font></td>
    <td width="96" style="padding-top: 2; padding-bottom: 2"><p align="center"><font
    face="����" size="2"><%=name%></font></td>
    <td width="93" style="padding-top: 2; padding-bottom: 2"><p align="center"><font
    face="����" size="2"><%=wdate%></font></td>
    <td width="71" style="padding-top: 2; padding-bottom: 2"><p align="center"><font
    face="����" size="2"><%=Rs("readnum")%></font></td>
  </tr>

<%
   Rs.Movenext
   Loop
%>
<% end if%>  
</table>
</center></div>

<hr width="600" align="center">
</div>
</body>
<% 
   Rs.Close
   db.Close
   Set Rs = Nothing
   Set db = Nothing
%>
</html>
