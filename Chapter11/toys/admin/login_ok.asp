<%
   id  = request("id")
   pwd = request("pwd")

if id ="admin" then  
   
   if pwd <> "1234" then
%>
<script language="javascript">
alert("비밀번호가 일치하지 않습니다.\n\n확인후 다시 시도해 주십시요");
history.back();
</script>
<% else
      session("id") = "admin"
   end if
%>
<script language="javascript">
parent.location.href = "../default.htm";
</script>
<% else %>
<script language="javascript">
alert("아이디가 일치하지 않습니다.\n\n확인후 다시 시도해 주십시요");
history.back();
</script>
<% end if %>