<% Option Explicit %>
<!--#include file="../inc/connection.asp"-->
<%
Dim code, Cmd
code = Request.Form("code")

Set Cmd = Server.CreateObject("ADODB.Command")

'�������� ������ �����, ����� ������ ���ν����� �����Ѵ�.
Cmd.ActiveConnection = strConnect
Cmd.CommandText = "usp_MakeCircle" 
Cmd.CommandType = adCmdStoredProc

Cmd.Parameters.Append  Cmd.CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)
Cmd.Parameters.Append Cmd.CreateParameter("@code",adInteger,adParamInput, ,code)

Dim ReturnValue
'������ ���ν����� �����ϴµ�, ���ڵ���� ������ �ʵ��� �Ѵ�.
Cmd.Execute , , adExecuteNoRecords 
ReturnValue = Cmd.Parameters("RETURN_VALUE").value

if ReturnValue = 1 then
	Response.Redirect "Queue.asp"
else	
%>
<script language="javascript">
  alert("�۾��߿� ������ ���ܼ� Ŭ���� ������ ���߽��ϴ�.\n\n����� �ٽ� �õ��� �ּ���");
  history.back();
</script>
<% end if%>
