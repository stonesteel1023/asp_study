<%

for i = 1 to request("test").count

 Response.write i & " ��° ������ " & Request("test")(i) & "<br>"

next

%>