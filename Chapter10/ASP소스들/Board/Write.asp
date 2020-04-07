<!--#include file="../inc/connection.asp"-->
<% table = request("table") %>
<html>

<head>
<link rel="stylesheet" href="../css/main.css">
</head>
<script language="javascript">
<!--
function sendit() 
{
        //제목
        if (document.all.title.value == "") {
                alert("제목을 입력해 주십시오.., please.");
                return false;
        }
        //글 내용
        if (document.all.memo.value == "" ) {
                alert("글을 작성 안하셨습니다. 글을 작성해 주십시요");
                return false;
        } 
        //비밀번호
        if (document.all.pwd.value == "" ) {
                alert("수정,삭제시 필요합니다. 비밀번호를 기입해 주십시요");
                return false;
        } 
		document.myform.submit();
        
}

function toggle(iObject)
{ 

	   if (iObject.style.display != "none") 
           iObject.style.display = "none" 
           else 
           iObject.style.display = "" 
} 
function focus_it()
{  
   document.all.title.focus();
}

function mysearch(object)
{
   if(object.Find.value=="")
   {
     alert("검색어를 넣어주세요");
     return;
   }
     
   object.submit();
}

//-->
</script>


<body bgcolor="#ffffff" OnLoad="document.all.name.focus();">

 
  <% if request("id") = "" then %>
<div align="center"><center><p><img
  src="images/write_new.gif" align="middle" WIDTH="31" HEIGHT="30"><font face="돋움"
  size="4" color="#004080"><strong> 글 올리기</strong></font> </p>
     
  </center></div>
   
<% else %>
<%
   	Set db = Server.CreateObject("ADODB.Connection")
	db.Open strconnect
	
	SQL = "SELECT title,content FROM " & table
	SQL = SQL & " where board_idx = " & request("id") 

    Set Rs = Server.CreateObject("ADODB.Recordset")

	Rs.Open SQL,db

title = request("head")
title = replace(title,chr(34),"&#34") '쌍따옴표 문제를 해결하다.

content = Rs("content")
%>

<div align="center"><div align="center"><center><p><img
  src="images/write_new.gif" align="middle" WIDTH="31" HEIGHT="30"><font face="돋움"
  size="4" color="#004080"><strong> 답변 올리기</strong></font> </p>
  </center></div>
  <dd align="center"><br>
  </dd>
  <font face="돋움"><div align="center"><center><p><img src="images/m_star.gif"
  WIDTH="20" HEIGHT="20"> &nbsp;<a href="javascript:toggle(document.all.quest)">이전의 질문 내용을 봅니다</a> 
  </font></p>
  </center></div><div ID="quest" ><div align="center"><div
  align="center"><center><table border="1" cellspacing="0" width="500">
    <tr>
      <td style="padding-left: 20; padding-right: 10; padding-top: 5; padding-bottom: 5"
      bgcolor="#FFFFFF"><pre><font color="navy" face="돋움" size="2"><%=content%></font></pre>
      </td>
    </tr>
  </table>
  </center></div>
  <dd align="center"><br>
  </dd>
  </div></div><% 
   Rs.Close
   db.Close
   Set Rs = Nothing
   Set db = Nothing
%>
<% end if %>
 <form method="POST" action="write_ok.asp" name="myform">
  <input type="hidden" name="table" value="<%=table%>">
  <input type="hidden" name="ref" value="<%=request("ref")%>">
  <input type="hidden" name="id"  value="<%=request("id")%>">
  <input type="hidden" name="step" value="<%=request("step")%>">
  <input type="hidden" name="level" value="<%=request("level")%>">
  <input type="hidden" name="GoTopage" value="<%=request("GoTopage")%>">
  

<div align="center"><center><table border="0" cellspacing="0" width="500"
  cellpadding="0">
    <tr>
      <td width="100" height="030" align="center" bgcolor="#E1F0FF"><font color="#004080"><font
      face="돋움" size="2">글쓴이</font></font></td>
      <td width="400" height="030" align="left"
      style="border-top: 1 dashed rgb(210,233,255); padding-left: 30; padding-right: 30"
      bgcolor="#FFFFFF"><%=Request.Cookies("user")("id")%></td>
    </tr>
    <tr>
      <td height="030" align="center" bgcolor="#E1F0FF"><font color="#004080"><font
      face="돋움" size="2">E-mail</font></font></td>
      <td height="030" align="left" style="padding-left: 30; padding-right: 30"
      bgcolor="#FFFFFF"><input type="text" name="mail" size="30" style="border: 1 dashed"
      value="<%=Request.Cookies("user")("mail")%>"></td>
    </tr>
    <tr>
      <td height="030" align="center" bgcolor="#E1F0FF"><font color="#004080"><font
      face="돋움" size="2">Homepage</font></font></td>
      <td height="030" align="left" style="padding-left: 30; padding-right: 30"
      bgcolor="#FFFFFF"><input type="text" name="homepage"
      <% if Request.Cookies("user")("site") = "" then%>value="http://" <%else%>value="<%=Request.Cookies("user")("site")%>"
      <%end if %> size="30" style="border: 1 dashed"></td>
    </tr>

    <tr>
      <td height="030" align="center" bgcolor="#E1F0FF"><font color="#004080"><font
      face="돋움" size="2">글제목</font></font></td>
      <td  height="030" align="left" style="padding-left: 30; padding-right: 30"
      bgcolor="#FFFFFF"><input type="text" name="title" size="30"
      <% if request("id") <> "" then %> value="<%=title%>" <% end if %>
      style="border-left: 1px dashed; border-right: 1 dashed; border-top: 1 dashed; border-bottom: 1 dashed"></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#E1F0FF"><font color="#004080"><font face="돋움"
      size="2">글</font></font></td>
      <td  align="left"
      style="border-bottom: 1 none rgb(210,233,255); padding-left: 20; padding-top: 5; padding-bottom: 5"
      bgcolor="#FFFFFF"><textarea wrap="hard" rows="10" name="memo" cols="55"
      style="font-family:돋움; border: 1 dashed" ></textarea></td>
    </tr>
    <tr>
      <td  height="030" align="center" bgcolor="#E1F0FF"><font face="돋움" size="2"
      color="#004080">비밀번호</font></td>
      <td  height="030" align="left"
      style="border-bottom: 1 dashed rgb(225,240,255); padding-left: 30; padding-right: 30"
      bgcolor="#FFFFFF"><input type="password" name="pwd" size="8" style="border: 1 dashed"> 
      &nbsp; <font face="돋움" size="2">(수정과 삭제시 필요!! )&nbsp;&nbsp;&nbsp; </font>
  
      <input
      type="checkbox" name="tag" value="ON"><font face="돋움" size="2">태그효과</font>
  </td>
    </tr>
  </table>
  </center></div><div align="center"></div><div align="center"><center><p><input
  type="button" value="글을 다 썼어요" name="write" OnClick="sendit()"
  style="background-color: rgb(225,240,255); color: rgb(0,53,106); font-weight: bolder">&nbsp; 
  <input type="reset" value="글 다시 쓸래요" name="reset"
  style="background-color: rgb(225,240,255); color: rgb(0,53,106); font-weight: bolder"></p>
  </center></div><div align="center"><center><p>　</p>
  </center></div></div>
</form>
</body>
</html>

  