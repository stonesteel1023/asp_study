<%@TRANSACTION=required Language=VBScript %>
<%
str = "taeyo"
%>
<HTML>
<HEAD></HEAD>
<BODY>
<center><font face="돋움" size="2">
필자의 WebID는 <%=str%>입니다.
<P>
트랜잭션의 결과 : 
<%
Sub OnTransactionCommit()   'Commit시 하고 싶은 처리..

    Response.write "모든 업무가 성공적으로 수행"

End Sub

Sub OnTransactionAbort()   'Commit시 하고 싶은 처리..

    response.write "업무 수행시 문제 발생 모든 작업을 취소"

End Sub
%>
</BODY>
</HTML>
