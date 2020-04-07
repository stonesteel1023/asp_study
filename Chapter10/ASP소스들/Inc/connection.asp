<!--METADATA TYPE= "typelib" FILE="C:\Program Files\Common Files\SYSTEM\ADO\msado15.dll"  -->
<%
Dim strConnect
strConnect = "Provider=SQLOLEDB; Server=taeyo; uid=sa; pwd=; DATABASE=Community"

Function UseText(value)
	value = replace(value, "&" , "&amp;")
	value = replace(value, "<", "&lt;")
	value = replace(value, ">", "&gt;")
	UseText = value
End Function

Function UseHTML(value)
	value = replace(value, "&lt;", "<")
	value = replace(value, "&gt;", ">")
	value = replace(value, "&quot;","'")
	value = replace(value, "&amp;", "&" )
	UseHTML = value
End Function 
%>
