<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML 3.0//EN" "html.dtd">
<html>
<head>

<%
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

    <title>ASP 찾기 예제 양식</title>

<%
'
' *** Modifying the Form's Search Scope.
'
' The form will search from the root of your web server's namespace and below
' (deep from "/" ). To search a subset of your server, for example, maybe just
' a PressReleases directory, modify the scope variable below to list the virtual path to
' search. The search will start at the directory you specify and include all sub-
' directories.

        FormScope = "/"

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
        SearchString = Request.Form("SearchString")
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

</head>

<body <%=FormBG%>>

<table>
    <tr><td><a HREF="http://www.microsoft.com/ntserver/search" target="_top"><img SRC="<%=FormLogo%>" VALIGN="MIDDLE" ALIGN="LEFT" border="0"></a></td></tr>
    <tr><td ALIGN="RIGHT"><h3>ASP 찾기 예제 양식</h3></td></tr>
</table>

<p>

        <form ACTION="<%=QueryForm%>" METHOD="POST">
                <table WIDTH="500">
                        <tr>
                                <td>아래에 쿼리를 입력하십시오.</td>
                        </tr>
                        <tr>
                                <td><input TYPE="TEXT" NAME="SearchString" SIZE="65" MAXLENGTH="100" VALUE="<%=SearchString%>"></td>
                                <td><input TYPE="SUBMIT" NAME="Action" VALUE="이동"></td>
                        </tr>
                                <tr>
                                <td ALIGN="RIGHT"><a HREF="ixtiphlp.htm">찾기에 관한 정보</a></td>
                        </tr>
                        <tr>
                        </tr>
                        <tr>
                                <td>
                                        <input NAME="FreeText" TYPE="CHECKBOX" <% if FreeText = "on" then
                                                                                Response.Write(" CHECKED")
                                                                end if %>><a href="ixtiphlp.htm#FreeTextQueries">의미 중심 쿼리</a> 사용
                                </td>
                        </tr>
                </table>
        </form>

<br>

<%
  if NewQuery then
    set Session("Query") = nothing
    set Session("Recordset") = nothing
    NextRecordNumber = 1

' Remove any leading and ending quotes from SearchString

        SrchStrLen = len(SearchString)

        if left(SearchString, 1) = chr(34) then
                SrchStrLen = SrchStrLen-1
                SearchString = right(SearchString, SrchStrLen)
        end if

        if right(SearchString, 1) = chr(34) then
                SrchStrLen = SrchStrLen-1
                SearchString = left(SearchString, SrchStrLen)
        end if

    if FreeText = "on" then
      CompSearch = "$contents " & chr(34) & SearchString & chr(34)
    else
      CompSearch = SearchString
    end if

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

<p>
<hr WIDTH="80%" ALIGN="center" SIZE="3">
<p>

<%
        LastRecordOnPage = NextRecordNumber + RS.PageSize - 1
        CurrentPage = RS.AbsolutePage
        if RS.RecordCount <> -1 AND RS.RecordCount < LastRecordOnPage then
            LastRecordOnPage = RS.RecordCount
        end if

        Response.Write "쿼리 " & chr(34) & "<I>"
        Response.Write SearchString & "</I>" & chr(34) & "에 일치하는 "

        if RS.RecordCount <> -1 then
            Response.Write RS.RecordCount & " 문서 중 "
        end if
        Response.Write NextRecordNumber & "에서 " & LastRecordOnPage & "까지의 문서<p>"

 %>

<% if Not RS.EOF and NextRecordNumber <= LastRecordOnPage then%>
		<table border="0">
		<colgroup width="105">
<% end if %>

<% Do While Not RS.EOF and NextRecordNumber <= LastRecordOnPage

        ' This is the detail portion for Title, Abstract, URL, Size, and
    ' Modification Date.

    ' If there is a title, display it, otherwise display the filename.
%>
    <p>

<% ' Graphically indicate rank of document with list of stars (*'s).

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

	<tr class="RecordTitle">
		<td align="right" valign="top" class="RecordTitle">
			<%= NextRecordNumber%>.
		</td>
		<td><b class="RecordTitle"> 
			<%if VarType(RS("DocTitle")) = 1 or RS("DocTitle") = "" then%>
				<a href="<%=RS("vpath")%>" class="RecordTitle"><%= Server.HTMLEncode( RS("filename") )%></a>
			<%else%>
				<a href="<%=RS("vpath")%>" class="RecordTitle"><%= Server.HTMLEncode(RS("DocTitle"))%></a>
			<%end if%>
		</b></td>
	</tr>

	<tr>
		<td valign="top" align="left">
			<img SRC="<%=stars%>">
			<br>
<%
    ' Construct the URL for hit highlighting
			WebHitsQuery = "CiWebHitsFile=" & Server.URLEncode( RS("vpath") )
			WebHitsQuery = WebHitsQuery & "&CiRestriction=" & Server.URLEncode( Q.Query )
			WebHitsQuery = WebHitsQuery & "&CiBeginHilite=" & Server.URLEncode( "<strong class=Hit>" )
			WebHitsQuery = WebHitsQuery & "&CiEndHilite=" & Server.URLEncode( "</strong>" )
			WebHitsQuery = WebHitsQuery & "&CiUserParam3=" & QueryForm
'	        WebHitsQuery = WebHitsQuery & "&CiLocale=" & Q.LocaleID
 %>
            <a href="oop/qsumrhit.htw?<%= WebHitsQuery %>"><img src="hilight.gif" align="left" alt="요약 모드를 사용하여 문서의 일치 용어 강조" WIDTH="22" HEIGHT="22"> 요약</a>
			<br>
			<a href="oop/qfullhit.htw?<%= WebHitsQuery %>&amp;CiHiliteType=Full"><img src="hilight.gif" align="left" alt="문서의 일치 용어 강조" WIDTH="22" HEIGHT="22"> 전체</a>
		</td>
		<td valign="top">
			<%if VarType(RS("characterization")) = 8 and RS("characterization") <> "" then%>
				<b><i>설명:  </i></b><%= Server.HTMLEncode(RS("characterization"))%>
            <%end if%>
			<p>
			<i class="RecordStats"><a href="<%=RS("vpath")%>" class="RecordStats" style="color:blue;">http://<%=Request("server_name")%><%=RS("vpath")%></a>
<%
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
            <br><%if RS("size") = "" then%>(알 수 없는 크기와 시간)<%else%>크기 <%=RS("size")%> 바이트 - <%=messagedate%> GMT<%end if%></i>
		</td>
	</tr>
	<tr>
	</tr>

<%
          RS.MoveNext
          NextRecordNumber = NextRecordNumber+1
      Loop
 %>

</table>

<p><br>

<%
  else   ' NOT RS.EOF
      if NextRecordNumber = 1 then
          Response.Write "쿼리에 일치하는 문서가 없습니다.<P>"
      else
          Response.Write "쿼리에 대한 문서가 없습니다.<P>"
      end if

  end if ' NOT RS.EOF


if NOT Q.OutOfDate then
' If the index is current, display the fact %>
<p>
    <i><b>색인은 최근의 것입니다.</b></i><br>
<%end if


  if Q.QueryIncomplete then
'    If the query was not executed because it needed to enumerate to
'    resolve the query instead of using the index, but AllowEnumeration
'    was FALSE, let the user know %>

    <p>
    <i><b>쿼리가 너무 많아서 완료할 수 없습니다.</b></i><br>
<%end if


  if Q.QueryTimedOut then
'    If the query took too long to execute (for example, if too much work
'    was required to resolve the query), let the user know %>
    <p>
    <i><b>쿼리가 너무 오래 걸려서 완료할 수 없습니다.</b></i><br>
<%end if%>


<table>

<%
'    This is the "previous" button.
'    This retrieves the previous page of documents for the query.
%>

<%SaveQuery = FALSE%>
<%if CurrentPage > 1 and RS.RecordCount <> -1 then %>
    <td align="left">
        <form action="<%=QueryForm%>" method="get">
            <input TYPE="HIDDEN" NAME="qu" VALUE="<%=SearchString%>">
                        <input TYPE="HIDDEN" NAME="FreeText" VALUE="<%=FreeText%>">
            <input TYPE="HIDDEN" NAME="sc" VALUE="<%=FormScope%>">
            <input TYPE="HIDDEN" name="pg" VALUE="<%=CurrentPage-1%>">
			<input TYPE="HIDDEN" NAME="RankBase" VALUE="<%=RankBase%>">
            <input type="submit" value="이전 <%=RS.PageSize%> 문서">
        </form>
    </td>
        <%SaveQuery = TRUE%>
<%end if%>


<%
'    This is the "next" button for unsorted queries.
'    This retrieves the next page of documents for the query.

  if Not RS.EOF then%>
    <td align="right">
        <form action="<%=QueryForm%>" method="get">
            <input TYPE="HIDDEN" NAME="qu" VALUE="<%=SearchString%>">
                        <input TYPE="HIDDEN" NAME="FreeText" VALUE="<%=FreeText%>">
            <input TYPE="HIDDEN" NAME="sc" VALUE="<%=FormScope%>">
            <input TYPE="HIDDEN" NAME="RankBase" VALUE="<%=RankBase%>">
			<input TYPE="HIDDEN" name="pg" VALUE="<%=CurrentPage+1%>">

            <% NextString = "다음 "
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
            <input type="submit" value="<%=NextString%>">
        </form>
    </td>
    <%SaveQuery = TRUE%>
<%end if%>

</table>

<% ' Display the page number %>

<%if RS.PageCount <> -1 then
        Response.Write RS.PageCount & " 페이지 중 "
  end if %>
<%=CurrentPage%> 페이지 

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

<!--#include file ="is2foot.inc"-->

</body>
</html>

