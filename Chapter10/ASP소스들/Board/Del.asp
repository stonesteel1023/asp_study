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
       
          //비밀번호
        if (document.forms[0].pwd.value == "" ) {
                alert("비밀번호를 기입해 주십시요");
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
  <p>　</p>
  <div align="center"><center><table border="0" cellspacing="0" width="450">
    <tr>
      <td bgcolor="#000080"><div align="center"><center><p><font color="#FFFF00"><big><big><big><strong>Delete</strong></big></big></big></font></td>
    </tr>
    <tr align="center">
      <td bgcolor="#F8FCDC" align="center">　<div align="center"><center><p><font face="돋움"
      size="2">삭제 : <font color="#004080"><%=request("name")%></font>님이&nbsp; 쓰신 <font color="#004080"><%=request("title")%></font></font></p>
      </center></div><div align="center"><center><p><font face="돋움" size="2">글의 
      비밀번호를 기입해 주십시요...</font></p>
      </center></div><div align="center"><center><p><font face="돋움" size="2" color="#000000">(첨부된</font><font
      face="돋움" size="2" color="#008080"> <strong>화일</strong></font><font face="돋움"
      size="2" color="#000000">이 있을경우 같이</font><font face="돋움" size="2"
      color="#008080"> </font><strong><font face="돋움" size="2" color="#FF8040">삭제</font></strong><font
      face="돋움" size="2" color="#008080"> </font><font face="돋움" size="2"
      color="#000000">됩니다)</font></p>
      </center></div><div align="center"><center><p><br>
      </td>
    </tr>
    <tr align="center">
      <td bgcolor="#E1E1E1"><div align="center"><center><p><br>
      <font face="돋움" size="2">비밀번호 :&nbsp; <input type="password" name="pwd"
      size="8"></font><br>
      <br>
      </td>
    </tr>
  </table>
  </center></div><div align="center"><center><p>　</p>
  </center></div><div align="center"><center><p><input type="button" value=" 삭 제 "
  OnClick="sendit()" name="del" style="color:black;background-color:#ffcc00;border:1 solid black;height:20">&nbsp;&nbsp; <font face="돋움" size="2">&lt;<a
  href="javascript:history.back();">삭제 안할래요!!</a>&gt;</font></p>
  </center></div>
</form>

<p align="center">　</p>
</body>
</html>
