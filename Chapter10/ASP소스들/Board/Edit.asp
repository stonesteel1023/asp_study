<!--#include file="../inc/connection.asp"-->
<%
GoTopage= request("GotoPage")
table = request("table")

%>

<html>

<head>
<link rel="stylesheet" href="../css/main.css">
<script language="javascript">
<!--
function sendit() 
{
        //�̸�
        if (document.forms[0].name.value == "")  {
               alert("�̸��� �Է��� �ֽʽÿ�.. please.");
               return false;
        }
        //����
        if (document.forms[0].title.value == "") {
                alert("������ �Է��� �ֽʽÿ�.., please.");
                return false;
        }
        //�� ����
        if (document.forms[0].memo.value == "" ) {
                alert("���� �ۼ� ���ϼ̽��ϴ�. ���� �ۼ��� �ֽʽÿ�");
                return false;
        } 
        //��й�ȣ
        if (document.forms[0].pwd.value == "" ) {
                alert("����,������ �ʿ��մϴ�. ��й�ȣ�� ������ �ֽʽÿ�");
                return false;
        } 
		document.forms[0].submit();
        
}
//-->
</script>
</head>
<%
if request("dog") <> "" then

   	Set db = Server.CreateObject("ADODB.Connection")
	db.Open strconnect

	SQL = "SELECT board_idx,name,title,mail,content,url,pwd,part FROM " & table & " where board_idx = " & request("dog")

    Set Grs = Server.CreateObject("ADODB.Recordset")

	Grs.Open SQL,db
	
	
	name = Grs("name")

	mail = Grs("mail")

	title = Grs("title") 'replace(SearchString,chr(34),"&#34")
	title = replace(title,chr(34),"&#34") '�ֵ���ǥ ������ �ذ��ϴ�.

	content = replace(Grs("content"),"&quot;","'")

	URL = Grs("url")
	info = Grs("part")
%>


<body bgcolor="white" link="#000080" vlink="#000080" alink="#000080">

<form method="POST" action="edit_ok.asp" id=form1 name=form1>
  <input type="hidden" name="gant" value="<%=Grs("board_idx")%>">
  <input type="hidden" name="gotopage" value="<%=gotopage%>">
  <input type="hidden" name="table" value="<%=table%>">
  
  <div align="center"><center><p><h3>Edit Contents</h3></p>
  </center></div><div align="center"></div><div align="center"><center><table border="0"
  cellspacing="0" width="600">
    <tr>
      <td width="100" height="30" align="center" bgcolor="#CBE0F8"><font color="#00376F"><font
      face="����" size="2">�۾���</font></font></td>
      <td width="400" height="030" align="left"
      style="border-top: 1 dashed rgb(210,233,255); padding-left: 20; padding-right: 30"
      bgcolor="#FFFFFF"><input type="text" name="name"
      value="<%=name%>" size="10" readonly style="border: 1 dashed"></font></td>
    </tr>
    <tr>
      <td width="100" height="30" align="center" bgcolor="#CBE0F8"><font color="#00376F"><font
      face="����" size="2">E-mail</font></font></td>
      <td width="500" height="30" align="left" style="padding-left: 20; padding-right: 30"
      bgcolor="#FFFFFF"><font face="����" size="2"><input type="text" name="mail"
      value="<%=mail%>" size="30" style="border: 1 dashed"></font></td>
    </tr>
    <tr>
      <td width="100" height="30" align="center" bgcolor="#CBE0F8"><font color="#00376F"><font
      face="����" size="2">Homepage</font></font></td>
      <td width="500" height="30" align="left" style="padding-left: 20; padding-right: 30"
      bgcolor="#FFFFFF"><font face="����" size="2"><input type="text" name="URL"
      value="<%=URL%>" size="30" style="border: 1 dashed"></font></td>
    </tr>
    <tr>
      <td width="100" height="30" align="center" bgcolor="#CBE0F8"><font color="#00376F"><font
      face="����" size="2">������</font></font></td>
      <td width="500" height="30" align="left" style="padding-left: 20; padding-right: 30"
      bgcolor="#FFFFFF"><font face="����" size="2"><input type="text" name="title" size="30"
      value="<%=title%>"
      style="border-left: 1px dashed; border-right: 1 dashed; border-top: 1 dashed; border-bottom: 1 dashed"></font></td>
    </tr>
    <tr>
      <td width="100" height="100" align="center" bgcolor="#CBE0F8"><font color="#00376F"><font
      face="����" size="2">��</font></font></td>
      <td width="500" height="100" align="left"
      style="padding-left: 20;  padding-top: 10; padding-bottom: 10" bgcolor="#FFFFFF"><font
      face="����" size="2"><textarea rows="10" name="memo" cols="60" style="font-family:����; border: 1 dashed"><%=content%></textarea></font></td>
    </tr>
    <tr>
      <td width="100" height="30" align="center" bgcolor="#CBE0F8"><font face="����" size="2"
      color="#00376F">��й�ȣ</font></td>
      <td width="500" height="30" align="left" style="padding-left: 20; padding-right: 30"
      bgcolor="#FFFFFF"><font face="����" size="2"><input type="password" name="pwd" size="8"
      style="border: 1 dashed"> &nbsp; (������ ������ ����)</font>
      &nbsp;&nbsp;&nbsp;
      <input type="checkbox" name="tag" value="ON"
      <% if info = "tagok" then%> checked<%end if%>><font face="����" size="2">�±�ȿ��</font> 
</td>
    </tr>
  </table>
  </center></div><div align="center"></div><div align="center"><center><p><input
  type="button" value="���� �Ϸ�" name="write" onclick="sendit()"
  style="color:navy;background-color:#E1F0FF;border:1 solid black;height:20">&nbsp; 
  &nbsp;&nbsp; <font color="#000080"><strong>&lt;<a href="javascript:history.back();">���� 
  ���ҷ���!!</a>&gt;</strong></font></p>
  </center></div><div align="center"><center><p>��</p>
  </center></div>
</form>
</body>
<% end if %>
<% 
   Grs.Close
   db.Close
   Set Grs = Nothing
   Set db = Nothing
%>
</html>
