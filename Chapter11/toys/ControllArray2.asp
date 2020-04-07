<%

for i = 1 to request("test").count

 Response.write i & " 번째 정보는 " & Request("test")(i) & "<br>"

next

%>