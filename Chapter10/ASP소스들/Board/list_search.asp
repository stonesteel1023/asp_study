<%@ LANGUAGE="VBSCRIPT" %>
<!--#include file="../inc/connection.asp"-->
<script LANGUAGE="VBScript" RUNAT="Server">
Function UseHTML(value)

	value = replace(value, "&lt;", "<")
	value = replace(value, "&gt;", ">")
	value = replace(value, "&quot;", "'")
	value = replace(value, "&amp;", "&" )

UseHTML = value
End Function
</script>
<%
     
PAGESIZE =10
table = request("table")
GotoPage = Request("GotoPage")

if GotoPage = "" or Cint(GotoPage) = 0 then
 GotoPage = 1
end if 


if request("Find") = "" then
    response.redirect "Noaccess.asp"

end if

	findit = trim(request("Find"))
	findit = replace(findit,"'","''")
	
	search_name = split(findit,"and")
	search_num = UBound(search_name)
		
	srch = request("search")
	
	'sub_search = request("sub")

   	Set db = Server.CreateObject("ADODB.Connection")
	db.Open strconnect

    sql = "select count(name) from " & table
    
    if search_num = 0 then
        sql = sql & " where " & srch & " LIKE '%" & findit & "%'"
    else
       sql = sql & " where 1=1 " '������ 1=1�� �ƹ� �ǹ̰� ���� �ִ�.
       for k = 0 to search_num
         sql = sql & " and " & srch & " LIKE '%" & trim(search_name(k)) & "%'"
       next  
    end if
    
    'Response.Write sql & "<br>"
         
	set rs = db.Execute (sql)
	reccount = rs(0)
	pagecount = Cint(rs(0) / pagesize)
'	Response.Write rs(0)
	
	'�Ҽ������ϴ� �ø��ϴ� �Լ��� ��� �׳� ���巶��..
	namuzi = (rs(0) mod pagesize)/pagesize
	if namuzi > 0 and namuzi < 0.5 then
      pagecount = pagecount + 1        
    end if
    
    rs.Close
	set rs = nothing
	    
  
    
	SQL = " SELECT TOP " & pagesize * GotoPage &" *, content,readnum FROM " & table
	
	
	if search_num = 0 then
        sql = sql & " where " & srch & " LIKE '%" & findit & "%'"
    else
       sql = sql & " where 1=1 " '������ 1=1�� �ƹ� �ǹ̰� ���� �ִ�.
       for k = 0 to search_num
         sql = sql & " and " & srch & " LIKE '%" & trim(search_name(k)) & "%'"
       next  
    end if
    	
	SQL = SQL & " ORDER BY ref desc, step asc "

    'Response.write sql
    set rs = db.Execute(sql)
	
%>

<html>

<head>
<title>�˻����</title>
<style type="text/css">  
 A            {text-decoration: none; color:navy}  
 A:hover      {text-decoration: none; color:blue}  
 .asp	{ color:black; background:yellow}
</style>
</head>

<body link="#000080" vlink="#000080" alink="#000080" bgcolor="#FFFFFF">
<%  if rs.BOF or rs.EOF then 
	%>

<p align="center">��</p>

<p align="center">��</p>

<p align="center">��</p>

<p align="center"><big><strong>ã���ô� �˻����� �˻������ �����ϴ�.</strong></big></p>

<p align="center">��</p>

<p align="center"><strong><font size="2" color="#808080">�Ʒ��� ��ư�� �����ø�</font><font
color="#800080" size="3">&nbsp; </font><font size="3" color="#FF8040">���� ������</font><font
color="#0080C0" size="3"> </font><font color="#808080"><font size="3">�� �̵��մϴ�.</font>.</font></strong></p>

<p align="center"><a href="javascript:history.back();"><img src="images/top.gif"
alt="�� �ø���� �̵�" border="0" width="44" height="41"></a></p>
<% else 

rs.Move pagesize * (GotoPage-1)  
'���ڵ��� Ŀ���� ������ �������� ���� ������� �ű��.
%>

<form method="POST" action="list_search.asp" id="form1" name="form1">
  <div align="center"><center><table border="0" cellpadding="0" cellspacing="0" width="500">
    
    <tr align="center">
      <td style="padding-left: 30; padding-right: 30"><div align="center"><center><table
      border="0" cellpadding="0" cellspacing="0" width="400">
        <tr>
          <td width="200" height="40" align="left" bgcolor="#91A4E6"><div align="right"><p><strong><font
          face="����" size="2" color="#333366">Board</font><font face="����" size="2"
          color="#FFFFFF"></font></strong> <select name="table" size="1"
          style="background-color: rgb(255,255,255); color: rgb(0,0,128); font-family: ����; border: 1 dashed">
     
       <option value="<%=table%>">&nbsp;���� �Խ��� &nbsp;</option>
             
             
            </select></td>
          <td width="200" height="40" align="center" bgcolor="#91A4E6"><strong><font face="����"
          size="2" color="#484893">��������</font><font face="����" size="2" color="#FFFFFF"> </font></strong><select
          name="search" size="1"
          style="background-color: rgb(255,255,255); color: rgb(0,0,128); font-family: ����; border: 1 dashed">
            <option value="title" <% if request("search") = "title" then%>selected<%end if %>>������</option>
            <option value="name" <% if request("search") = "name" then%>selected<%end if %>>�۾���</option>
            <option value="content" <% if request("search") = "content" then%>selected<%end if %>>�۳���</option>
          </select></td>
        </tr>
        <tr align="center">
          <td width="200" height="30" align="left" bgcolor="#91A4E6"><div align="right"><p><input
          type="text" name="Find" size="20" style="border: 1 dashed" value="<%=findit%>"></td>
          <td width="200" height="30" align="center" bgcolor="#91A4E6"><input type="submit"
          value="  �˻�  " name="B1"
         style="color:black;background-color:#ffcc00;border:1 solid black;height:20"></td>
        </tr>

         </table>
      </center></div><div align="center"><center><p><font face="����" size="2">�� <b><%=reccount%></b> 
      ���� ��¥������ �˻��Ǿ����ϴ�.&nbsp;
      <font size="2"
      face="arial" color="navy">( <%=Gotopage%> / <%=pagecount%> page )</font>
       &nbsp;&nbsp;&nbsp;&lt; <a href="list.asp?table=<%=table%>">�Խ�������</a> &gt;</font></p>
      </center></div><div align="center"><center><table border="0" cellpadding="0"
      cellspacing="0" width="430">
<%


i = 1
Do until rs.EOF Or i>rs.PageSize

	name = UseText(rs("name"))

	If Len(name) > 5 Then
		name = Mid(name,1,6) & "..."
	End If
	if request("search") = "writer" then  '�˻��� �۾��̶�� �۾����� �˻�� ����...
		name = replace(name,findit,"<strong><font color=red>"&findit&"</font></strong>")
	end if

	title = UseText(rs("title"))

	If Len(title) > 18 Then
		title = Mid(title,1,19) & "..."
	End If

	if request("search") = "title" then  '�˻��� �� �����̶�� ������ �˻�� ����...
		title = replace(title,findit,"<strong><font color=red>"&findit&"</font></strong>")
	end if

	wdate = rs("writeday")
	content = UseText(rs("content"))

	If Len(content) > 150 Then
		content = Mid(content,1,150) & "......"
	End If

	readnum = rs("readnum")

	if request("gotopage")<>"0" then
		s_num = (Cint(request("gotopage")-1)*10)+i
	else
		s_num = (request("gotopage")*10)+i
	end if

	if request("gotopage")="" then
		s_num = (request("gotopage")*10)+i
	end if
%>
        <tr>
          <td width="320" height="25" bgcolor="#CBD7F3" style="border: 1 dashed; padding-left: 10"><font
          face="����" size="2"><strong><%=s_num%></strong>. &nbsp;<% if rs("re_level") > 0 then %><% wid = 4 * rs("re_level") %> <img src="images/x.gif"
          width="<%=wid%>" height="16"> <img src="images/re.gif" width="26" height="16"> <% end if %> <a
          href="content.asp?id=<%=rs("board_idx")%>&found=is"><%=title%></a></font></td>
          <td width="110" height="25" bgcolor="#C1D8F0"
          style="border-right: 1 dashed; border-top: 1 dashed; border-bottom: 1 dashed"><div
          align="center"><center><p><font face="����" size="2">��ȸ�� : <%=readnum%></font></td>
        </tr>
        <tr align="center">
          <td width="430" height="20" colspan="2"
          style="border-left: 1 dashed; border-right: 1 dashed; border-bottom: 1 dashed; padding-left: 10; padding-right: 10; padding-top: 5; padding-bottom: 5"><div
          align="left"><p><font face="����" color="#000000" size="2"><%=content%></font></td>
        </tr>
        <tr align="center">
          <td width="320" height="20" style="padding-left: 10; padding-top: 5; padding-bottom: 20"><div
          align="left"><p><font face="����" size="2" color="#7093DB">��¥ : </font><font
          face="����" size="2" color="#004080"><%=wdate%></font></td>
          <td width="110" height="20" style="padding-top: 5; padding-bottom: 20"><div align="center"><center><p><font
          face="����" size="2" color="#7093DB">�۾��� :</font><font face="����" size="2"
          color="#004080"><%=name%></font></td>
        </tr>
<%
   rs.Movenext
   i = i + 1
   Loop
%>
      </table>
      </center></div></td>
    </tr>
    <tr align="center">
      <td>
      <div align="center"><center><table border="0" cellspacing="0" width="650">
      <tr><td align="center">
<!---------------���� 10��, ���� 10�� ��-------------->      
<font face="����" size="2">         
   <%
		blockPage=Int((GotoPage-1)/10)*10+1

		if blockPage = 1 Then
			Response.Write "<font color=silver>[���� 10��]</font> ["
		Else%> 
			<a href="list_search.asp?table=<%=table%>&amp;GotoPage=<%=blockPage-10%>&find=<%=findit%>&search=<%=srch%>">[���� 10��]</a> [&nbsp;
		<%End If
		
		i=1
		Do Until i > 10 or blockPage > pagecount
			If blockPage=int(GotoPage) Then%> 
				<font face="arial" size="2" color="silver"><%=blockPage%></font>&nbsp;
			<%Else%> 
			<font face="����" size="2"><b>
				<a href="list_search.asp?table=<%=table%>&amp;GotoPage=<%=blockPage%>&find=<%=findit%>&search=<%=srch%>"><%=blockPage%></a></b></font>&nbsp;
			<%End If
			
			blockPage=blockPage+1
			i=i+1
		Loop
			
		if blockPage > pagecount Then
			Response.Write "] <font color=silver>[���� 10��]</font>"
		Else%> 
			] <a href="list_search.asp?table=<%=table%>&amp;GotoPage=<%=blockPage%>&find=<%=findit%>&search=<%=srch%>">[���� 10��]</a> 
      <%End If%>
</font>
<!---------------���� 10��, ���� 10�� ��-------------->  
        </td>
      </tr>
    </table>
    </center></div>
      
      </td>
    </tr>
  </table>
  </center></div>
</form>
<% end if %>
</body>
<% 
   rs.Close
   db.Close
   Set rs = Nothing
   Set db = Nothing
%>
</html>
