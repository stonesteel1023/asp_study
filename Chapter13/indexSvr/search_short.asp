<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML 3.0//EN" "html.dtd">
<html>

<head><%
' ********** INSTRUCTIONS FOR QUICK CUSTOMIZATION **********
'
' This form is set up for easy customization. It allows you to modify the
' page logo, the page background, the page title and simple query
' parameters by modifying a few files and form variables. The procedures
' to do this are explained below.
'
'
' *** Modifying the Form Logo:

' The logo for the form is named is2logo.gif. To change the page logo, simply
' name your logo is2logo.gif and place in the same directory as this form. If
' your logo is not a GIF file, or you don't want to copy it, change the following
' line so that the logo variable contains the URL to your logo.

        FormLogo = "is2logo.gif"

'
' *** Modifying the Form's background pattern.

' You can use either a background pattern or a background color for your
' form. If you want to use a background pattern, store the file with the name
' is2bkgnd.gif in the same directory as this file and remove the remark character
' the single quote character) from the line below. Then put the remark character on
' the second line below.
'
' If you want to use a different background color than white, simply edit the
' bgcolor line below, replacing white with your color choice.

'       FormBG = "background = " & chr(34) & "is2bkgnd.gif" & chr(34)
        FormBG = "bgcolor = " & chr(34) & "#FFFFFF" & chr(34)


' *** Modifying the Form's Title Text.

' The Form's title text is set on the following line.
%>

<title>taeyo's Search</title>
<%
'
' *** Modifying the Form's Search Scope.
'
' The form will search from the root of your web server's namespace and below
' (deep from "/" ). To search a subset of your server, for example, maybe just
' a PressReleases directory, modify the scope variable below to list the virtual path to
' search. The search will start at the directory you specify and include all sub-
' directories.

      nd = request("nd")
      if nd ="asp" then
        FormScope = "/taeo/ASP/"
      elseif nd ="mfc" then
        FormScope = "/taeo/mfc/"
      end if

     FreeText = request("FreeText")
     '

' *** Modifying the Number of Returned Query Results.
'
' You can set the number of query results returned on a single page
' using the variable below.

        PageSize = 10

'
' *** Setting the Locale.
'
' The following line sets the locale used for queries. In most cases, this
' should match the locale of the server. You can set the locale below.

        SiteLocale = "KO"

' ********** END QUICK CUSTOMIZATION SECTIONS ***********
%>
<link REL="STYLESHEET" HREF="is2style.css" TYPE="text/css">
<meta NAME="DESCRIPTION" CONTENT="Microsoft Index Server v2.0용 ASP 쿼리 예제 양식">
<meta NAME="AUTHOR" CONTENT="Index Server Team">
<meta NAME="KEYWORDS" CONTENT="query, content, hit">
<meta NAME="SUBJECT" CONTENT="sample form">
<meta NAME="MS.CATEGORY" CONTENT="Internet">
<meta NAME="MS.LOCALE" CONTENT="KO">
<%
DIM strMessageDate, tmpAMPM, tmpTime, MessageDate
' Set Initial Conditions
    NewQuery = FALSE
    UseSavedQuery = FALSE
    SearchString = ""

    QueryForm = Request.ServerVariables("PATH_INFO")

' Did the user press a SUBMIT button to execute the form? If so get the form variables.
    if Request.ServerVariables("REQUEST_METHOD") = "POST" then
        'SearchString 이 찾고자 하는 쿼리
        SearchString = Request.Form("SearchString")
        'FreeText는 체크박스(의미중심의 쿼리)
        FreeText = Request.Form("FreeText")
        ' NOTE: this will be true only if the button is actually pushed.
        if Request.Form("Action") = "이동" then
            NewQuery = TRUE
			RankBase=1000
        end if
    end if
    
    if Request.ServerVariables("REQUEST_METHOD") = "GET" then
        SearchString = Request.QueryString("qu")
        
                FreeText = Request.QueryString("FreeText")
                FormScope = Request.QueryString("sc")
				RankBase = Request.QueryString("RankBase")
        if Request.QueryString("pg") <> "" then
            NextPageNumber = Request.QueryString("pg")
            NewQuery = FALSE
            UseSavedQuery = TRUE
        else
            NewQuery = SearchString <> ""
        end if
    end if
%>
<style type="text/css">  
 A            {text-decoration: none; color:navy}  
 A:hover      {text-decoration: underline; color:red}  
</style>
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


<body <%=FormBG%> Onload="document.form1.SearchString.focus();">

<form ACTION="<%=QueryForm%>" METHOD="POST" id="form1" name="form1">
  <input type="hidden" name="Action" value="이동"><div align="left"><table border="0" cellpadding="0" cellspacing="0" width="600" style="margin-left: 50">
     <tr>
      <td height="20" align="center"><img src="../images/findnow1.gif" WIDTH="80" HEIGHT="20"><img src="../images/find_s1.gif" border="0" WIDTH="80" HEIGHT="20">  &nbsp;&nbsp;
&nbsp;&nbsp;      <font size="2" face="돋움"><a href="search.asp"><img src="../images/findnow01.gif" border="0" WIDTH="15" HEIGHT="20"> <img src="../images/find_s2.gif" border="0" WIDTH="80" HEIGHT="20"></a></font></td>
    </tr>
   <tr>
      <td height="50" align="center"><strong><font face="Century Gothic" color="#000080" size="5">Taeyo's </font><font color="#408080" face="Century Gothic"><big><big>Search</big></big></font></strong> 
      &nbsp; <strong>(<font face="Arial" color="#000040"><big> Index Server</big></font> )</strong></td>
    </tr>
    <tr>
      <td height="10" align="center"></td>
    </tr>
    <tr>
      <td height="10" align="center"><div align="center"><div align="center"><center><table border="0" cellpadding="0" cellspacing="0" width="450">
         <tr>
          <td height="30" colspan="2" bgcolor="#DEEAF3"><div align="center"><center><p><font face="돋움" size="2">찾고자하는 단어를 입력하세요</font></td>
        </tr>
        <tr align="center">
          <td height="30" colspan="2" bgcolor="#DEEAF3"><font face="돋움" size="2"><div align="center"><center><p><font color="#000040"><input id="1" type="radio" value="asp" <% if request("nd")="asp" or request("nd")="" then %>checked<%end if%> name="nd"></font><font color="#004080">
          <label for="1">
          <strong>ASP</strong></font><font color="#000040"> 공부방</font>
          </label>
          &nbsp;&nbsp;&nbsp; <input id="2" type="radio" name="nd" value="mfc" <% if request("nd")="mfc" then %>checked<%end if%>> 
          <label for="2">
          <strong><font color="#408080">MFC</font></strong><font color="#000040"> 공부방 </font>
          </label>
          </font></td>
        </tr>
        <tr align="center">
          <td colspan="2" bgcolor="#DEEAF3" height="40"><div align="center"><center><p><font face="돋움" size="2"><input NAME="SearchString" SIZE="20" MAXLENGTH="100" VALUE="<%=SearchString%>"> <a href="javascript:sendform1();"><img src="../images/searchnow.gif" border="0" WIDTH="73" HEIGHT="22"></a> </font></td>
        </tr>
        <tr align="center">
          <td ALIGN="right" width="250" bgcolor="#DEEAF3" height="40"><font face="돋움" size="2"><input NAME="FreeText" TYPE="CHECKBOX" <% if FreeText = "on" then
              Response.Write(" CHECKED")
          end if %> value="on"><a href="ixtiphlp.htm#FreeTextQueries">의미 중심 쿼리</a> 사용 </font></td>
          <td ALIGN="right" width="250" bgcolor="#DEEAF3" height="50"><div align="left"><p><font face="돋움" size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a HREF="ixtiphlp.htm">찾기에 관한 정보</a></font></td>
        </tr>
      </table>
      </center></div>
   
      </div></td>
    </tr>
  </table>
  </div>
</form>
<div align="left">
 
<table border="0" width="700">
  <tr>
    <td>

<%
  if NewQuery then
    set Session("Query") = nothing
    set Session("Recordset") = nothing
    NextRecordNumber = 1

' Remove any leading and ending quotes from SearchString

        SrchStrLen = len(SearchString)
        
        '이밑의 두 if문은 넘어온 쿼리중에 "를 제거하는 것입니다.
        if left(SearchString, 1) = chr(34) then
                SrchStrLen = SrchStrLen-1
                SearchString = right(SearchString, SrchStrLen)
                SearchString = Searchstring & " and (#filename *.htm OR #filename *.html )"
        end if

        if right(SearchString, 1) = chr(34) then
                SrchStrLen = SrchStrLen-1
                SearchString = left(SearchString, SrchStrLen)
                SearchString = Searchstring & " and (#filename *.htm OR #filename *.html )"
        end if


    if FreeText = "on" then
      CompSearch = "$contents " & chr(34) & SearchString & chr(34)
    else
      CompSearch = SearchString
    end if

'Response.Write "aaa" & CompSearch

'*************************************************************

    set Q = Server.CreateObject("ixsso.Query")
    set util = Server.CreateObject("ixsso.Util")

    Q.Query = CompSearch
    Q.SortBy = "rank[d]"
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

<p align="center"><font face="돋움" size="2"><%
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

 %> <% if Not RS.EOF and NextRecordNumber <= LastRecordOnPage then%> </font></p>
<% end if %> <% Do While Not RS.EOF and NextRecordNumber <= LastRecordOnPage

        ' This is the detail portion for Title, Abstract, URL, Size, and
    ' Modification Date.

    ' If there is a title, display it, otherwise display the filename.
%> <% ' Graphically indicate rank of document with list of stars (*'s).

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

%> 
<div align="center"><center>

<table border="0" width="550">
  <tr>
    <td align="right" height="25" width="450" bgcolor="#E3E7F9" style="padding-left: 10; padding-right: 10; padding-top: 3; padding-bottom: 3"><p align="left"><font face="돋움" size="2"><%if VarType(RS("DocTitle")) = 1 or RS("DocTitle") = "" then%> <a href="<%=RS("vpath")%>"><%= Server.HTMLEncode( RS("filename") )%></a> <%else%> <a href="<%=RS("vpath")%>"><%= Server.HTMLEncode(RS("DocTitle"))%></a> <%end if%> </font></td>
    
    <td height="25" width="100" bgcolor="#ffffff" style="border: 1 dashed rgb(227,231,249);"><font face=" 돋움" size="2"><img SRC="../images/<%=stars%>"></font></td>
  </tr>
  <% if false then %>
  <tr>
    <td align="left" width="550" style="border: 1 dashed rgb(227,231,249); padding-left: 20; padding-right: 20; padding-top: 10; padding-bottom: 10" colspan="2">
    <font face="돋움" size="2"> <%
' ==== begin convert to local time format

			strMessageDate = ( year(RS("write"))&"/"&month(RS("write")) & "/" & day(RS("write")) )
			if Hour(RS("write")) < 12 then
				if Hour(RS("write")) = 0 then
					tHour= 12
				else
					tHour = Hour(RS("write"))
				end if
					AmPm = "오전"
			else	 	if Hour(RS("write")) <> 12 then
						tHour = Hour(RS("write")) - 12
					else
						tHour = Hour(RS("write"))	
					end if
				AmPm = "오후"	
			end if
			locTime = AmPm & " " & tHour & ":" & Minute(RS("write")) & ":" & Second(RS("write"))
			messagedate= strMessageDate & " " & locTime
' ==== end  convert to local time format
%> 
<%=strmessagedate%></font></td>
  </tr>
  <%end if%>
</table>
</center></div>

<%
          RS.MoveNext
          NextRecordNumber = NextRecordNumber+1
      Loop
 %></p>

<p align="center"><font face="돋움" size="2"><%
  else   ' NOT RS.EOF
      if NextRecordNumber = 1 then
          Response.Write "쿼리에 일치하는 문서가 없습니다.<P>"
      else
          Response.Write "쿼리에 대한 문서가 없습니다.<P>"
      end if

  end if ' NOT RS.EOF


if NOT Q.OutOfDate then
' If the index is current, display the fact %> </font></p>

<p align="center"><font face="돋움" size="2"><b>색인은 최근의 것입니다.</b><br>
<%end if


  if Q.QueryIncomplete then
'    If the query was not executed because it needed to enumerate to
'    resolve the query instead of using the index, but AllowEnumeration
'    was FALSE, let the user know %></font></p>

<p align="center"><font face="돋움" size="2"><b>쿼리가 너무 많아서 완료할 수 
없습니다.</b><br>
<%end if


  if Q.QueryTimedOut then
'    If the query took too long to execute (for example, if too much work
'    was required to resolve the query), let the user know %></font></p>

<p align="center"><font face="돋움" size="2"><b>쿼리가 너무 오래 걸려서 
완료할 수 없습니다.</b><br>
<%end if%></font></p>
<div align="center"><center>

<table>
<%
'    This is the "previous" button.
'    This retrieves the previous page of documents for the query.
%>
<%SaveQuery = FALSE%>
<%if CurrentPage > 1 and RS.RecordCount <> -1 then %>
  <tr>
    <td align="left"><form action="<%=QueryForm%>" method="get" id="form2" name="form2">
      <input type="hidden" name="nd" value="<%=nd%>">
      <input type="hidden" name="qu" value="<%=SearchString%>"><input type="hidden" name="FreeText" value="<%=FreeText%>"><input type="hidden" name="sc" value="<%=FormScope%>"><input type="hidden" name="pg" value="<%=CurrentPage-1%>"><input type="hidden" name="RankBase" value="<%=RankBase%>"><p><font face="돋움" size="2"><input type="submit" value="이전 <%=RS.PageSize%> 문서" id="submit1" name="submit1"> </font></p>
    </form>
    </td>
<%SaveQuery = TRUE%>
<%end if%>
<%
'    This is the "next" button for unsorted queries.
'    This retrieves the next page of documents for the query.

  if Not RS.EOF then%>
    <td align="right"><form action="<%=QueryForm%>" method="get" id="form3" name="form3">
    <input type="hidden" name="nd" value="<%=nd%>">
      <input type="hidden" name="qu" value="<%=SearchString%>"><input type="hidden" name="FreeText" value="<%=FreeText%>"><input type="hidden" name="sc" value="<%=FormScope%>"><input type="hidden" name="RankBase" value="<%=RankBase%>"><input type="hidden" name="pg" value="<%=CurrentPage+1%>"><% NextString = "다음 "
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
<p><font face="돋움" size="2"><input type="submit" value="<%=NextString%>" id="submit2" name="submit2"> </font></p>
    </form>
    </td>
<%SaveQuery = TRUE%>
<%end if%>
  </tr>
</table>
</center></div><% ' Display the page number %>
<%
    ' If either of the previous or back buttons were displayed, save the query
    ' and the recordset in session variables.
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
</td></tr></table></div>
</body>
</html>
