<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML 3.0//EN" "html.dtd">
<html>

<head <title>
<link REL="STYLESHEET" HREF="is2style.css" TYPE="text/css">
<meta NAME="DESCRIPTION" CONTENT="Microsoft Index Server v2.0">
<meta NAME="AUTHOR" CONTENT="Index Server Team">
<meta NAME="KEYWORDS" CONTENT="query, content, hit">
<meta NAME="SUBJECT" CONTENT="sample form">
<meta NAME="MS.CATEGORY" CONTENT="Internet">
<meta NAME="MS.LOCALE" CONTENT="KO">
<style type="text/css">  
 A            {text-decoration: none; color:navy}  
 A:hover      {text-decoration: underline; color:red}  
</style>
<title></title>
</head>
<script language="javascript">
<!--
function sendform1()
{
  if(document.form1.SearchString.value=="")
  {
    alert("� ������ �˻����� �������ּ���");
    document.form1.SearchString.focus();
    return;
  }
  document.form1.submit();
}
//-->
</script>

<%

'���� �ε��������� ����� �������� �̸� �����Ѵ�.
'FormScope�� �˻��� ���丮, FreeText�� �ڿ��� ������ ��뿩��, 
'pagesize�� �� �������� ����� �˻������� ����, SiteLocale�� ����� ���
  FormScope = "/"
  FreeText = request("FreeText")  
  PageSize = 10
  SiteLocale = "KO"

'NewQuery�� ������ ������ ���� ������ ���� ���������� �ľ�
'SearchString�� ����ڰ� �Է��� �������ڸ� �����ϴ� ����
'QueryForm �������� ������ ASP �������� ��ġ��θ� ����
    NewQuery = FALSE
    UseSavedQuery = FALSE
    SearchString = ""

    QueryForm = Request.ServerVariables("PATH_INFO")

'post������� �Ѿ���� �������� �����ϴ� ��
    if Request.ServerVariables("REQUEST_METHOD") = "POST" then
        'SearchString�� ����ڰ� �Է��� ã���� �ϴ� ������ ����
        SearchString = Request.Form("SearchString")
        
        'Ư�����ϸ��� �˻��ϰ� �Ϸ��� �Ʒ��� �ּ��� ������.�׷���htm�������� �˻��� ���̴�
        'SearchString = Searchstring & " and (#filename *.htm OR #filename *.html )"
        
        'FreeText�� �ǹ��߽��� ����(�ڿ��� ����)�� ����� �������� ����
        FreeText = Request.Form("FreeText")
        
        '�� �������� �ϴܿ��� ����������,������������ Ŭ���Ͽ� ���������� ������ ���
        '�� ���� NextPageNumber�� �����ϰ� ���� �ε����������� �װ��� �����Ѵ�.
        if Request.form("pg") <> "" then
            NextPageNumber = Request.form("pg")
            NewQuery = FALSE
            UseSavedQuery = TRUE
        else
            NewQuery = true
        end if
    end if
 
  ' ���� ������,������������ ���ؼ� �غ� �Ǿ�����. 
  ' ���ڰ� ������ �� �ҽ������� post��ĸ� ����ϰԲ� �����Ͽ���
  '  if Request.ServerVariables("REQUEST_METHOD") = "GET" then
  '      SearchString = Request.QueryString("qu")
  '      'SearchString = Searchstring & " and (#filename *.htm OR #filename *.html )"
  '              FreeText = Request.QueryString("FreeText")
  '              FormScope = Request.QueryString("sc")
  '              RankBase = Request.QueryString("RankBase")
  '      if Request.QueryString("pg") <> "" then
  '          NextPageNumber = Request.QueryString("pg")
  '          NewQuery = FALSE
  '          UseSavedQuery = TRUE
  '      else
  '          NewQuery = SearchString <> ""
  '      end if
  '  end if
%>

<body bgcolor="#ffffff">

<!--���⼭���� ����ڷκ��� ������ �Է¹޴� ���̴�. -->
<form ACTION="<%=QueryForm%>" METHOD="POST" id="form1" name="form1">
  <input type="hidden" name="Action" value="�̵�"><div align="center"><center><table
  border="0" cellpadding="0" cellspacing="0" width="600">
    <tr>
      <td height="20" align="center"><img src="images/findnow1.gif" WIDTH="80" HEIGHT="20"><img
      src="images/find_s2.gif" border="0" WIDTH="80" HEIGHT="20"> &nbsp;&nbsp; &nbsp;&nbsp; <font
      size="2" face="����"><a href="search_short.asp"><img src="images/findnow01.gif"
      border="0" WIDTH="15" HEIGHT="20"><img src="images/find_s1.gif" border="0" WIDTH="80"
      HEIGHT="20"></a></font></td>
    </tr>
    <tr>
      <td height="50" align="center"><strong><font face="Century Gothic" color="#000080"
      size="5">Taeyo's </font><font color="#408080" face="Century Gothic"><big><big>Search</big></big></font></strong> 
      &nbsp; <strong>(<font face="Arial" color="#000040"><big> Index Server</big></font> )</strong></td>
    </tr>
    <tr>
      <td height="10" align="center"></td>
    </tr>
    <tr>
      <td height="10" align="center"><div align="center"><center><table border="0"
      cellpadding="0" cellspacing="0" width="450">
        <tr>
          <td height="30" colspan="2" bgcolor="#DEEAF3"><div align="center"><center><p><font
          size="2" face="����">ã�����ϴ� �ܾ �Է��ϼ���</font></td>
        </tr>
        <tr align="center">
          <td colspan="2" bgcolor="#DEEAF3" height="40"><div align="center"><center><p><font
          size="2" face="����"><input NAME="SearchString" SIZE="20" MAXLENGTH="100"
          VALUE="<%=SearchString%>"> <a href="javascript:sendform1();"><img
          src="images/searchnow.gif" border="0" WIDTH="73" HEIGHT="22"></a> </font></td>
        </tr>
        <tr align="center">
          <td ALIGN="right" width="250" bgcolor="#DEEAF3" height="40"><font size="2" face="����"><input
          NAME="FreeText" TYPE="CHECKBOX"
          <% if FreeText = "on" then
              Response.Write(" CHECKED")
          end if %>
          value="on"><a href="ixtiphlp.htm#FreeTextQueries">�ǹ� �߽� ����</a> ��� </font></td>
          <td ALIGN="right" width="250" bgcolor="#DEEAF3" height="50"><div align="left"><p><font
          size="2" face="����">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a
          HREF="ixtiphlp.htm">ã�⿡ ���� ����</a></font></td>
        </tr>
      </table>
      </center></div></td>
    </tr>
  </table>
  </center></div>
</form>
<!-- ��������� ����ڿ��� ���̴� ���̴�.-->
<div align="center"><center>

<table border="0" width="700">
  <tr>
    <td><%

 if NewQuery then
    set Session("Query") = nothing
    set Session("Recordset") = nothing

    NextRecordNumber = 1

        SrchStrLen = len(SearchString)
        
        '�̹��� �� if���� �Ѿ�� �����߿� "�� �����ϴ� ���Դϴ�.
        if left(SearchString, 1) = chr(34) then
                SrchStrLen = SrchStrLen-1
                SearchString = right(SearchString, SrchStrLen)
                SearchString = Searchstring '& " and (#filename *.htm OR #filename *.html )"
        end if

        if right(SearchString, 1) = chr(34) then
                SrchStrLen = SrchStrLen-1
                SearchString = left(SearchString, SrchStrLen)
                SearchString = Searchstring '& " and (#filename *.htm OR #filename *.html )"
        end if


    if FreeText = "on" then
      CompSearch = "$contents " & chr(34) & SearchString & chr(34)
    else
      CompSearch = SearchString
    end if

'*************************************************************

'������Ʈ�� ����Ͽ� �ε��������� ���Ǹ� �ϰ��Ѵ�.
'�� �κ��� �������� ������ �����ϴ� �κ��̴�. ����� �����ؼ� �ٶ���

    set Q = Server.CreateObject("ixsso.Query")
    set util = Server.CreateObject("ixsso.Util")

    Q.Query = CompSearch
    Q.SortBy = "rank[d],write[d]"
    Q.Columns = "DocTitle, vpath, filename, size, write, characterization, rank"
	Q.MaxRecords = 300 


        if FormScope <> "/" then
                util.AddScopeToQuery Q, FormScope, "deep"
        end if

        if SiteLocale<>"" then
                Q.LocaleID = util.ISOToLocaleID(SiteLocale)
        end if

    set RS = Q.CreateRecordSet("nonsequential")

    RS.PageSize = PageSize
    ActiveQuery = TRUE

  elseif UseSavedQuery then
    if IsObject( Session("Query") ) And IsObject( Session("RecordSet") ) then
      set Q = Session("Query")
      set RS = Session("RecordSet")

      if RS.RecordCount <> -1 and NextPageNumber <> -1 then
        RS.AbsolutePage = NextPageNumber
        NextRecordNumber = RS.AbsolutePosition
      end if

      ActiveQuery = TRUE
    else
      Response.Write "���� - ����� ������ �����ϴ�."
    end if
  end if

  if ActiveQuery then
    if not RS.EOF then
    
 %>
<hr WIDTH="80%" ALIGN="center" SIZE="2">
    <p align="center"><font size="2" face="����"><%
        LastRecordOnPage = NextRecordNumber + RS.PageSize - 1
        CurrentPage = RS.AbsolutePage
        if RS.RecordCount <> -1 AND RS.RecordCount < LastRecordOnPage then
            LastRecordOnPage = RS.RecordCount
        end if

        Response.Write "���� " & chr(34) & ""
        Response.Write "<font color=teal><b>" & SearchString & "</b></font>" & "" & chr(34) & "�� ��ġ�ϴ� "

        if RS.RecordCount <> -1 then
            Response.Write "<b>" & RS.RecordCount & "</b>" & " ���� �� "
        end if
        Response.Write "<b>" & NextRecordNumber & "</b>" & "���� " & "<b>" & LastRecordOnPage & "</b>" & "������ ����"
        
        if RS.PageCount <> -1 then
        Response.Write " &nbsp;&nbsp;&nbsp;( <font color=teal><b>" & CurrentPage & " / " & RS.PageCount  & "</b> page</font> )<p>"
        end if

 %> 
 
 
 <% Do While Not RS.EOF and NextRecordNumber <= LastRecordOnPage

	if NextRecordNumber = 1 then
		RankBase=RS("rank")
	end if

	if RankBase>1000 then
		RankBase=1000
	elseif RankBase<1 then
		RankBase=1
	end if

	NormRank = RS("rank")/RankBase

	if NormRank > 0.80 then
		stars = "rankbtn5.gif"
	elseif NormRank > 0.60 then
		stars = "rankbtn4.gif"
	elseif NormRank > 0.40 then
		stars = "rankbtn3.gif"
	elseif NormRank >.20 then
		stars = "rankbtn2.gif"
	else stars = "rankbtn1.gif"

	end if

%></font> </p>
    <div align="center"><center><table border="0" width="550">
      <tr>
        <td align="right" height="30" width="450" bgcolor="#BEC7F1"
        style="padding-left: 10; padding-right: 10; padding-top: 3; padding-bottom: 3"><p
        align="left"><b><font size="2" face="����"><%if VarType(RS("DocTitle")) = 1 or RS("DocTitle") = "" then%> <a href="<%=RS("vpath")%>"><%= Server.HTMLEncode( RS("filename") )%></a> <%else%> <a
        href="<%=RS("vpath")%>"><%= Server.HTMLEncode(RS("DocTitle"))%></a> <%end if%> </font></b></td>
        <td height="30" width="100" bgcolor="#FFFFFF" style="border: 1 dashed rgb(18,30,84)"><font
        face=" ����" size="2"><img SRC="images/<%=stars%>"></font></td>
      </tr>
      <tr>
        <td align="left" width="550"
        style="border: 1 dashed rgb(18,30,84); padding-left: 20; padding-right: 20; padding-top: 10; padding-bottom: 10"
        colspan="2"><%if VarType(RS("characterization")) = 8 and RS("characterization") <> "" then%>
<p><font face="����" size="2"><b>����: </b><%
content = Server.HTMLEncode(RS("characterization"))
content = replace(content,searchstring,"<b><font color=red>"&searchstring&"</font></b>")
%> <%=content%> <%end if%> </font></p>
        <p align="center"><font face="����" size="2"><a href="<%=RS("vpath")%>"
        class="RecordStats" style="color:blue;">http://<%=Request("server_name")%><%=RS("vpath")%></a> 
        
<br>
<%if RS("size") = "" then%>
(�� �� ���� ũ��� �ð�)
<%else%>
ũ�� <%=RS("size")%> ����Ʈ - <%=Rs("write")%>
<%end if%> </font></td>
      </tr>
    </table>
    </center></div>
<%
   RS.MoveNext
   NextRecordNumber = NextRecordNumber+1
   Loop
%>
<p align="center"><font face="����" size="2"><%
  else   ' NOT RS.EOF
      if NextRecordNumber = 1 then
          Response.Write "<center><p> </p><br> <br> <br><font size=2 color=navy>"
          Response.Write "������ ��ġ�ϴ� ������ �����ϴ�.</font></center><P>"
      else
          Response.Write "<center><p> </p><br> <br> <br><font size=2 color=navy>"
          Response.Write "������ ���� ������ �����ϴ�.</fon></center><P>"
      end if

  end if ' NOT RS.EOF


if NOT Q.OutOfDate then
' If the index is current, display the fact %> </font></p>
    <p align="center"><font face="����" size="2"><b>������ �ֱ��� ���Դϴ�.</b><br>
<%end if


  if Q.QueryIncomplete then
'    If the query was not executed because it needed to enumerate to
'    resolve the query instead of using the index, but AllowEnumeration
'    was FALSE, let the user know %>    </font></p>
    <p align="center"><font face="����" size="2"><b>������ �ʹ� ���Ƽ� �Ϸ��� �� 
    �����ϴ�.</b><br>
<%end if


  if Q.QueryTimedOut then
'    If the query took too long to execute (for example, if too much work
'    was required to resolve the query), let the user know %>    </font></p>
    <p align="center"><font face="����" size="2"><b>������ �ʹ� ���� �ɷ��� 
    �Ϸ��� �� �����ϴ�.</b><br>
<%end if%>    </font></p>
    <div align="center"><center><table>
<%
'    This is the "previous" button.
'    This retrieves the previous page of documents for the query.
%>
<%'********************************************************************
  ' �� �κ��� ������������ �̵��ϰ� �ϴ� �κ��̴�. ������ get����̾�����,
  ' ���ڴ� post������� �ٲپ���. �������� ���� �ҽ��� ���Ͽ� �� ���̸� �ݹ� �����Ͻ� ���� ���� ���̴�.  
%>
<%SaveQuery = FALSE%>
<%if CurrentPage > 1 and RS.RecordCount <> -1 then %>
      <tr>
        <td align="left"><form action="<%=QueryForm%>" method="post" id="form2" name="form2">
        <input type="hidden" name="SearchString"
          value="<%=SearchString%>"><input type="hidden" name="FreeText" value="<%=FreeText%>"><input
          type="hidden" name="sc" value="<%=FormScope%>"><input type="hidden" name="pg"
          value="<%=CurrentPage-1%>"><input type="hidden" name="RankBase" value="<%=RankBase%>"><p><font
          face="����" size="2"><input type="submit" value="���� <%=RS.PageSize%> ����"
          id="submit1" name="submit1"> </font></p>
        </form>
        </td>
<%SaveQuery = TRUE%>
<%end if%>
<%
'********************************************************************
  ' �� �κ��� ������������ �̵��ϰ� �ϴ� �κ��̴�. ������ get����̾�����,
  ' ���ڴ� post������� �ٲپ���. �������� ���� �ҽ��� ���Ͽ� �� ���̸� �ݹ� �����Ͻ� ���� ���� ���̴�.

  if Not RS.EOF then%>
        <td align="right"><form action="<%=QueryForm%>" method="post" id="form3" name="form3">
        <input type="hidden" name="SearchString"
          value="<%=SearchString%>"><input type="hidden" name="FreeText" value="<%=FreeText%>"><input
          type="hidden" name="sc" value="<%=FormScope%>"><input type="hidden" name="RankBase"
          value="<%=RankBase%>"><input type="hidden" name="pg" value="<%=CurrentPage+1%>"><% NextString = "���� "
               if RS.RecordCount <> -1 then
                   NextSet = (RS.RecordCount - NextRecordNumber) + 1
                   if NextSet > RS.PageSize then
                       NextSet = RS.PageSize
                   end if
                   NextString = NextString & NextSet & " ����"
               else
                   NextString = "������ " & NextString & " ������"
               end if
             %>
<p><font
          face="����" size="2"><input type="submit" value="<%=NextString%>" id="submit2"
          name="submit2"> </font></p>
        </form>
        </td>
<%SaveQuery = TRUE%>
<%end if%>
      </tr>
    </table>
    </center></div>
<%

'����� ��ü���� �����Ѵ�.

    if SaveQuery then
        set Session("Query") = Q
        set Session("RecordSet") = RS
    else
        RS.close
        Set RS = Nothing
        Set Q = Nothing
        set Session("Query") = Nothing
        set Session("RecordSet") = Nothing
    end if
 %>
<% end if %>
</td>
  </tr>
</table>
</center></div>
</body>
</html>
