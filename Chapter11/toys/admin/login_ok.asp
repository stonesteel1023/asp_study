<%
   id  = request("id")
   pwd = request("pwd")

if id ="admin" then  
   
   if pwd <> "1234" then
%>
<script language="javascript">
alert("��й�ȣ�� ��ġ���� �ʽ��ϴ�.\n\nȮ���� �ٽ� �õ��� �ֽʽÿ�");
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
alert("���̵� ��ġ���� �ʽ��ϴ�.\n\nȮ���� �ٽ� �õ��� �ֽʽÿ�");
history.back();
</script>
<% end if %>