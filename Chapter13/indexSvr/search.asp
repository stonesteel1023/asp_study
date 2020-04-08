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
    alert("어떤 것으로 검색할지 기입해주세요");
    document.form1.SearchString.focus();
    return;
  }
  document.form1.submit();
}
//-->
</script>

<%

'차후 인덱스서버에 사용할 변수들을 미리 지정한다.
'FormScope는 검색할 디렉토리, FreeText는 자연어 질의의 사용여부, 
'pagesize는 한 페이지에 출력할 검색문서의 갯수, SiteLocale은 사용할 언어
  FormScope = "/"
  FreeText = request("FreeText")  
  PageSize = 10
  SiteLocale = "KO"

'NewQuery는 현재의 쿼리가 이전 쿼리와 같은 쿼리인지를 파악
'SearchString은 사용자가 입력한 쿼리문자를 저장하는 변수
'QueryForm 변수에는 현재의 ASP 페이지의 위치경로를 저장
    NewQuery = FALSE
    UseSavedQuery = FALSE
    SearchString = ""

    QueryForm = Request.ServerVariables("PATH_INFO")

'post방식으로 넘어오는 변수들을 저장하는 곳
    if Request.ServerVariables("REQUEST_METHOD") = "POST" then
        'SearchString에 사용자가 입력한 찾고자 하는 쿼리를 저장
        SearchString = Request.Form("SearchString")
        
        '특정파일만을 검색하게 하려면 아래의 주석을 지우자.그러면htm문서만이 검색될 것이다
        'SearchString = Searchstring & " and (#filename *.htm OR #filename *.html )"
        
        'FreeText는 의미중심의 쿼리(자연어 질의)를 사용할 것인지를 저장
        FreeText = Request.Form("FreeText")
        
        '이 페이지의 하단에서 이전페이지,다음페이지를 클릭하여 페이지값을 전송할 경우
        '그 값을 NextPageNumber에 저장하고 차후 인덱스서버에게 그값을 전달한다.
        if Request.form("pg") <> "" then
            NextPageNumber = Request.form("pg")
            NewQuery = FALSE
            UseSavedQuery = TRUE
        else
            NewQuery = true
        end if
    end if
 
  ' 이전 페이지,다음페이지를 위해서 준비가 되어졌다. 
  ' 필자가 편집한 이 소스에서는 post방식만 사용하게끔 수정하였다
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

<!--여기서부터 사용자로부터 쿼리를 입력받는 폼이다. -->
<form ACTION="<%=QueryForm%>" METHOD="POST" id="form1" name="form1">
  <input type="hidden" name="Action" value="이동"><div align="center"><center><table
  border="0" cellpadding="0" cellspacing="0" width="600">
    <tr>
      <td height="20" align="center"><img src="images/findnow1.gif" WIDTH="80" HEIGHT="20"><img
      src="images/find_s2.gif" border="0" WIDTH="80" HEIGHT="20"> &nbsp;&nbsp; &nbsp;&nbsp; <font
      size="2" face="돋움"><a href="search_short.asp"><img src="images/findnow01.gif"
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
          size="2" face="돋움">찾고자하는 단어를 입력하세요</font></td>
        </tr>
        <tr align="center">
          <td colspan="2" bgcolor="#DEEAF3" height="40"><div align="center"><center><p><font
          size="2" face="돋움"><input NAME="SearchString" SIZE="20" MAXLENGTH="100"
          VALUE="<%=SearchString%>"> <a href="javascript:sendform1();"><img
          src="images/searchnow.gif" border="0" WIDTH="73" HEIGHT="22"></a> </font></td>
        </tr>
        <tr align="center">
          <td ALIGN="right" width="250" bgcolor="#DEEAF3" height="40"><font size="2" face="돋움"><input
          NAME="FreeText" TYPE="CHECKBOX"
          <% if FreeText = "on" then
              Response.Write(" CHECKED")
          end if %>
          value="on"><a href="ixtiphlp.htm#FreeTextQueries">의미 중심 쿼리</a> 사용 </font></td>
          <td ALIGN="right" width="250" bgcolor="#DEEAF3" height="50"><div align="left"><p><font
          size="2" face="돋움">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a
          HREF="ixtiphlp.htm">찾기에 관한 정보</a></font></td>
        </tr>
      </table>
      </center></div></td>
    </tr>
  </table>
  </center></div>
</form>
<!-- 여기까지가 사용자에게 보이는 폼이다.-->
<div align="center"><center>

<table border="0" width="700">
  <tr>
    <td><%

 if NewQuery then
    set Session("Query") = nothing
    set Session("Recordset") = nothing

    NextRecordNumber = 1

        SrchStrLen = len(SearchString)
        
        '이밑의 두 if문은 넘어온 쿼리중에 "를 제거하는 것입니다.
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

'컴포넌트를 사용하여 인덱스서버에 질의를 하게한다.
'이 부분이 실질적인 쿼리를 수행하는 부분이다. 상당히 주의해서 바라보자

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
      Response.Write "오류 - 저장된 쿼리가 없습니다."
    end if
  end if

  if ActiveQuery then
    if not RS.EOF then
    
 %>
<hr WIDTH="80%" ALIGN="center" SIZE="2">
    <p align="center"><font size="2" face="돋움"><%
        LastRecordOnPage = NextRecordNumber + RS.PageSize - 1
        CurrentPage = RS.AbsolutePage
        if RS.RecordCount <> -1 AND RS.RecordCount < LastRecordOnPage then
            LastRecordOnPage = RS.RecordCount
        end if

        Response.Write "쿼리 " & chr(34) & ""
        Response.Write "<font color=teal><b>" & SearchString & "</b></font>" & "" & chr(34) & "에 일치하는 "

        if RS.RecordCount <> -1 then
            Response.Write "<b>" & RS.RecordCount & "</b>" & " 문서 중 "
        end if
        Response.Write "<b>" & NextRecordNumber & "</b>" & "에서 " & "<b>" & LastRecordOnPage & "</b>" & "까지의 문서"
        
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
        align="left"><b><font size="2" face="돋움"><%if VarType(RS("DocTitle")) = 1 or RS("DocTitle") = "" then%> <a href="<%=RS("vpath")%>"><%= Server.HTMLEncode( RS("filename") )%></a> <%else%> <a
        href="<%=RS("vpath")%>"><%= Server.HTMLEncode(RS("DocTitle"))%></a> <%end if%> </font></b></td>
        <td height="30" width="100" bgcolor="#FFFFFF" style="border: 1 dashed rgb(18,30,84)"><font
        face=" 돋움" size="2"><img SRC="images/<%=stars%>"></font></td>
      </tr>
      <tr>
        <td align="left" width="550"
        style="border: 1 dashed rgb(18,30,84); padding-left: 20; padding-right: 20; padding-top: 10; padding-bottom: 10"
        colspan="2"><%if VarType(RS("characterization")) = 8 and RS("characterization") <> "" then%>
<p><font face="돋움" size="2"><b>설명: </b><%
content = Server.HTMLEncode(RS("characterization"))
content = replace(content,searchstring,"<b><font color=red>"&searchstring&"</font></b>")
%> <%=content%> <%end if%> </font></p>
        <p align="center"><font face="돋움" size="2"><a href="<%=RS("vpath")%>"
        class="RecordStats" style="color:blue;">http://<%=Request("server_name")%><%=RS("vpath")%></a> 
        
<br>
<%if RS("size") = "" then%>
(알 수 없는 크기와 시간)
<%else%>
크기 <%=RS("size")%> 바이트 - <%=Rs("write")%>
<%end if%> </font></td>
      </tr>
    </table>
    </center></div>
<%
   RS.MoveNext
   NextRecordNumber = NextRecordNumber+1
   Loop
%>
<p align="center"><font face="돋움" size="2"><%
  else   ' NOT RS.EOF
      if NextRecordNumber = 1 then
          Response.Write "<center><p> </p><br> <br> <br><font size=2 color=navy>"
          Response.Write "쿼리에 일치하는 문서가 없습니다.</font></center><P>"
      else
          Response.Write "<center><p> </p><br> <br> <br><font size=2 color=navy>"
          Response.Write "쿼리에 대한 문서가 없습니다.</fon></center><P>"
      end if

  end if ' NOT RS.EOF


if NOT Q.OutOfDate then
' If the index is current, display the fact %> </font></p>
    <p align="center"><font face="돋움" size="2"><b>색인은 최근의 것입니다.</b><br>
<%end if


  if Q.QueryIncomplete then
'    If the query was not executed because it needed to enumerate to
'    resolve the query instead of using the index, but AllowEnumeration
'    was FALSE, let the user know %>    </font></p>
    <p align="center"><font face="돋움" size="2"><b>쿼리가 너무 많아서 완료할 수 
    없습니다.</b><br>
<%end if


  if Q.QueryTimedOut then
'    If the query took too long to execute (for example, if too much work
'    was required to resolve the query), let the user know %>    </font></p>
    <p align="center"><font face="돋움" size="2"><b>쿼리가 너무 오래 걸려서 
    완료할 수 없습니다.</b><br>
<%end if%>    </font></p>
    <div align="center"><center><table>
<%
'    This is the "previous" button.
'    This retrieves the previous page of documents for the query.
%>
<%'********************************************************************
  ' 이 부분은 이전페이지로 이동하게 하는 부분이다. 원래는 get방식이었으나,
  ' 필자는 post방식으로 바꾸었다. 여러분은 원본 소스와 비교하여 그 차이를 금방 이해하실 수가 있을 것이다.  
%>
<%SaveQuery = FALSE%>
<%if CurrentPage > 1 and RS.RecordCount <> -1 then %>
      <tr>
        <td align="left"><form action="<%=QueryForm%>" method="post" id="form2" name="form2">
        <input type="hidden" name="SearchString"
          value="<%=SearchString%>"><input type="hidden" name="FreeText" value="<%=FreeText%>"><input
          type="hidden" name="sc" value="<%=FormScope%>"><input type="hidden" name="pg"
          value="<%=CurrentPage-1%>"><input type="hidden" name="RankBase" value="<%=RankBase%>"><p><font
          face="돋움" size="2"><input type="submit" value="이전 <%=RS.PageSize%> 문서"
          id="submit1" name="submit1"> </font></p>
        </form>
        </td>
<%SaveQuery = TRUE%>
<%end if%>
<%
'********************************************************************
  ' 이 부분은 다음페이지로 이동하게 하는 부분이다. 원래는 get방식이었으나,
  ' 필자는 post방식으로 바꾸었다. 여러분은 원본 소스와 비교하여 그 차이를 금방 이해하실 수가 있을 것이다.

  if Not RS.EOF then%>
        <td align="right"><form action="<%=QueryForm%>" method="post" id="form3" name="form3">
        <input type="hidden" name="SearchString"
          value="<%=SearchString%>"><input type="hidden" name="FreeText" value="<%=FreeText%>"><input
          type="hidden" name="sc" value="<%=FormScope%>"><input type="hidden" name="RankBase"
          value="<%=RankBase%>"><input type="hidden" name="pg" value="<%=CurrentPage+1%>"><% NextString = "다음 "
               if RS.RecordCount <> -1 then
                   NextSet = (RS.RecordCount - NextRecordNumber) + 1
                   if NextSet > RS.PageSize then
                       NextSet = RS.PageSize
                   end if
                   NextString = NextString & NextSet & " 문서"
               else
                   NextString = "문서의 " & NextString & " 페이지"
               end if
             %>
<p><font
          face="돋움" size="2"><input type="submit" value="<%=NextString%>" id="submit2"
          name="submit2"> </font></p>
        </form>
        </td>
<%SaveQuery = TRUE%>
<%end if%>
      </tr>
    </table>
    </center></div>
<%

'사용한 객체들을 정리한다.

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
