<% Option Explicit %>
<!--#include file="../inc/connection.asp"-->
<%
Dim code, Cmd
code = Request.Form("code")

Set Cmd = Server.CreateObject("ADODB.Command")

'서버와의 연결을 만들고, 사용할 스토어드 프로시져를 정의한다.
Cmd.ActiveConnection = strConnect
Cmd.CommandText = "usp_MakeCircle" 
Cmd.CommandType = adCmdStoredProc

Cmd.Parameters.Append  Cmd.CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)
Cmd.Parameters.Append Cmd.CreateParameter("@code",adInteger,adParamInput, ,code)

Dim ReturnValue
'스토어드 프로시져를 수행하는데, 레코드셋은 만들지 않도록 한다.
Cmd.Execute , , adExecuteNoRecords 
ReturnValue = Cmd.Parameters("RETURN_VALUE").value

if ReturnValue = 1 then
	Response.Redirect "Queue.asp"
else	
%>
<script language="javascript">
  alert("작업중에 오류가 생겨서 클럽을 만들지 못했습니다.\n\n잠시후 다시 시도해 주세요");
  history.back();
</script>
<% end if%>
