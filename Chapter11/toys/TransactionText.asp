<%@TRANSACTION=required Language=VBScript %>
<%
str = "taeyo"
%>
<HTML>
<HEAD></HEAD>
<BODY>
<center><font face="����" size="2">
������ WebID�� <%=str%>�Դϴ�.
<P>
Ʈ������� ��� : 
<%
Sub OnTransactionCommit()   'Commit�� �ϰ� ���� ó��..

    Response.write "��� ������ ���������� ����"

End Sub

Sub OnTransactionAbort()   'Commit�� �ϰ� ���� ó��..

    response.write "���� ����� ���� �߻� ��� �۾��� ���"

End Sub
%>
</BODY>
</HTML>
