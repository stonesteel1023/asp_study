<!--#include file="../inc/connection.asp"-->
<% Response.Expires=0 %> 
<%
table = request("table")
%>
<html>
<head>
<title>�Խ���</title>
<link rel="stylesheet" href="../css/main.css">
<script language="JavaScript">
<!--

function toggle(iObject)
{ 
   if (iObject.style.display != "none") 
      iObject.style.display = "none" 
   else 
      iObject.style.display = "" 
} 

//board�� ����...
function goboard()
{
	document.QAbd.submit();
}

function mysearch(object)
{
   if(object.Find.value=="")
   {
     alert("�˻�� �־��ּ���");
     return;
   }
     
   object.submit();
}
-->
</script>
</head>
<%
pagesize=10
GotoPage = Request("GotoPage")

if GotoPage = "" then
 GotoPage = 1
end if 

   	Set db = Server.CreateObject("ADODB.Connection")
	db.open strconnect

	sql = "select count(name) from " & table
	set rs = db.Execute (sql)
	pagecount = Clng(rs(0) / pagesize)
	
		'�Ҽ������ϴ� �ø��ϴ� �Լ��� ��� �׳� ���巶��..
		namuzi = (rs(0) mod pagesize)/pagesize
		if namuzi > 0 and namuzi < 0.5 then
			pagecount = pagecount + 1        
		end if
	
    rs.Close
    
 					
    '������ ������������ ���ڵ常�� �����´�			
 	SQL = " SELECT TOP " & pagesize * GotoPage & " board_idx,name,title,ref,step,re_level,mail,writeday,readnum "
	sql = sql & "FROM " & table & " order by ref desc, step " 

    set rs = db.Execute(sql)
%>
<body bgcolor="#FFFFFF" vlink="#000080">
<%  if rs.BOF  then 
%>
<br>
<p align="center"><img src="images/commu.gif" WIDTH="400" HEIGHT="300"><br><br>
<a href="write.asp?table=<%=table%>"><img src="images/insert.gif" border="0" WIDTH="23" HEIGHT="22"></a> 
</p>
<%
else 
rs.Move pagesize * (GotoPage-1)  
'���ڵ��� Ŀ���� ������ �������� ���� ������� �ű��.
%> 
<div align="left">
<table border="0" cellspacing="0" width="700">
  <tr>

    <td>
    <div align="center"><div align="center"><center><table border="0" width="640">
      <tr>
        <td width="640" colspan="5" bgcolor="#FFFFFF">    
      <form action="list_search.asp" Method="POST" name="searchform1">
      <center>
        <a href="write.asp?table=<%=table%>">
        <img src="images/insert.gif" border="0" align="absmiddle" WIDTH="23" HEIGHT="22"></a>
		<a href="master_ok.asp?table=<%=table%>">
		<img src="images/admin.gif" border="0" align="absmiddle" WIDTH="23" HEIGHT="22"></a>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    
       <input type="hidden" name="table" value="<%=table%>">
     <select name="search" size="1" style="font-family: ����ü">
            <option selected value="title">����</option>
            <option value="name">�۾���</option>
            <option value="content">�� ����</option>
          </select>&nbsp; <input type="text" name="Find" size="15" style="border: 1px dashed">&nbsp;
           <a href="javascript:mysearch(document.searchform1);"><img src="images/searchnow.gif" align="middle" border="0" WIDTH="73" HEIGHT="22"></a>
       </font>
       &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
       &nbsp;<font size="2" face="����" color="navy">( <%=Gotopage%> / <%=pagecount%> page 
	)</font>
      </center>
    </form>   
    </td>
      </tr>
      <tr>
        <td width="20" style="padding-top: 3; padding-bottom: 3"></td>
        <td width="350" bgcolor="#5A69A7" style="padding-top: 3; padding-bottom: 3" align="center"><p align="center"><font face="����" size="2" color="#DBDEEE"><strong>��&nbsp; ��</strong></font></td>
        <td width="96" bgcolor="#5A69A7" style="padding-top: 3; padding-bottom: 3" align="center"><p align="center"><font face="����" size="2" color="#DBDEEE"><strong>�۾���</strong></font></td>
        <td width="123" bgcolor="#5A69A7" style="padding-top: 3; padding-bottom: 3" align="center"><p align="center"><font face="����" size="2" color="#DBDEEE"><strong>��¥</strong></font></td>
        <td width="51" bgcolor="#5A69A7" style="padding-top: 3; padding-bottom: 3" align="center"><p align="center"><font face="����" size="2" color="#DBDEEE"><strong>��ȸ</strong></font></td>
      </tr>
      
<%
i = 1
Do until rs.EOF or i > pagesize

board_idx = rs("board_idx")
name = rs("name")
mail = rs("mail")

If Len(name) > 6 Then
name = Mid(name,1,6) & ".."
End If

title = replace(rs("title"),"&quot;","'")

If Len(title) > 22 Then
title = Mid(title,1,23) & "..."
End If

'��¥�� �����ؼ� �����ִ� �κ�
yymmdd = rs("writeday")
yy= year(yymmdd)
mm = right("0" & month(yymmdd),2)
dd = right("0" & day(yymmdd),2)
h = right("0" & hour(yymmdd),2)
mi = right("0" & minute(yymmdd),2)
yymmdd = yy & "/" & mm & "/" & dd & " (" & h & ":" & mi & ")"
readnum = rs("readnum")
%>

 <tr height="25">
        <td width="20" style="padding-top: 0; padding-bottom: 0"><%if DateDiff("n",rs("writeday"),Now()) < 1440 then %>
<p><img src="images/dot.gif" width="13" height="12"><% end if %> </td>
        <td width="350" style="padding-left: 10; padding-top: 0; padding-bottom: 0" height="20" onmouseover="this.style.background='#DBDEEE'" onmouseout="this.style.background='white'"><% if rs("re_level") > 0 then %>
<% wid = 4 * rs("re_level") %>
<p><img src="images/x.gif" width="<%=wid%>" height="16"> <img src="images/re.gif" width="26" height="16"> <% end if %> 
 <a href="content.asp?id=<%=board_idx%>&amp;GotoPage=<%=GotoPage%>&amp;table=<%=table%>"><font face="����" size="2"><%=title%></font></a></td>
        <td width="96" style="padding-left: 10; padding-bottom: 0"><p align="left"><font face="����" size="2"><% if mail <>"" then %><a href="mailto:<%=mail%>" <%=alt%>><%end if%> <%=name%> <% if mail<>"" then %></a><%end if%> </font></td>
        <td width="123" style="padding-left: 5; padding-bottom: 0"><p align="left"><font face="����" size="2"><%=yymmdd%></font></td>
        <td width="51" style="padding-top: 0; padding-bottom: 0""><p align="center"><font face="����" size="2"><%=readnum%></font></td>
      </tr>
<% 
   rs.Movenext
   i = i + 1
   Loop
   
%>
    </table>
    </center></div>
    
</div><hr WIDTH="600">
</td>
  </tr>
</table></center></div>


     <table border="0" cellspacing="0" width="650">
          <tr>
            <td><p align="center"><font face="����" size="2">
           
    <% if Cint(GotoPage) <> 1 then %>
    <a href="list.asp?GoTopage=1">
    <% end if %><font color="#ffcc00">|����</font>
    <% if Cint(GotoPage) <> 1 then %></a><% end if %>


            &nbsp;          
<!---------------���� 10��, ���� 10�� ��-------------->      
<font face="����" size="2">         
   <%
		blockPage=Int((GotoPage-1)/10)*10+1

		if blockPage = 1 Then
			Response.Write "<font color=silver>[���� 10��]</font> ["
		Else%> 
			<a href="list.asp?table=<%=table%>&amp;GotoPage=<%=blockPage-10%>">[���� 10��]</a> [&nbsp;
		<%End If
		
		i=1
		Do Until i > 10 or blockPage > pagecount
			If blockPage=int(GotoPage) Then%> 
				<font face="arial" size="2" color="silver"><%=blockPage%></font>&nbsp;
			<%Else%> 
			<font face="����" size="2"><b>
				<a href="list.asp?table=<%=table%>&amp;GotoPage=<%=blockPage%>"><%=blockPage%></a></b></font>&nbsp;
			<%End If
			
			blockPage=blockPage+1
			i=i+1
		Loop
			
		if blockPage > pagecount Then
			Response.Write "] <font color=silver>[���� 10��]</font>"
		Else%> 
			] <a href="list.asp?table=<%=table%>&amp;GotoPage=<%=blockPage%>">[���� 10��]</a> 
      <%End If%>
&nbsp;      
 <% if Cint(GotoPage) <> Cint(pagecount) then %> 
<a href="list.asp?GoTopage=<%=pagecount%>">
<% end if %><font color="#ffcc00">����|</font>
<% if Cint(GotoPage) <> Cint(pagecount) then %></a><% end if %>

</font>
<!---------------���� 10��, ���� 10�� ��-------------->  
               
          </td></tr>
        </table>
<% end if '�Խ��� �ҽ��� �� %>

</body>
<% 
   rs.Close
   db.Close
   Set rs = Nothing
   Set db = Nothing
%>
</html>
