<HTML>
<HEAD><link rel="stylesheet" href="../css/main.css"></HEAD>
<BODY bgcolor="#BBCAE8">
<table border=0>
<tr><td><a href="../Queue/Queue.asp" target="_top">동호회 메인</a></P>
<P>&nbsp;</P>
<P>&nbsp;</P>
<P><img src="images/join.gif" align="absmiddle" WIDTH="14" HEIGHT="14">
<A href="Content.asp?code=<%=request("code")%>" target="board">Main</a></p>
<P><img src="images/join.gif" align="absmiddle" WIDTH="14" HEIGHT="14">
<A href="../board/list.asp?table=<%=request("BoardName")%>" target="board">
우리 게시판</a></p>
<p><img src="images/join.gif" align="absmiddle" WIDTH="14" HEIGHT="14">
<A href="memlist.asp?code=<%=request("code")%>&BoardName=<%=request("BoardName")%>" target="board">
회원 목록</a>
<P>&nbsp;</P>
</td></tr></table>
</BODY>
</HTML>
