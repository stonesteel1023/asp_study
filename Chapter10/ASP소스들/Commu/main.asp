<html>
<head>
<title>동호회 메인 페이지</title>
</head>
<frameset cols="120,*" frameborder="0" border="0" framespacing="0"> 
	<frame src="menu.asp?code=<%=request("code")%>&BoardName=<%=request("BoardName")%>" name="Menu" noresize scrolling="no">
	<frame src="Content.asp?code=<%=request("code")%>" name="board">
</frameset>
<noframes>
  <body bgcolor="#FFFFFF">
  </body>
</noframes>
