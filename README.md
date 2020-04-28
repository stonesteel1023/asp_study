# ASP 인코딩문제 해결
```
<%@Language="VBScript" CODEPAGE="65001" %>
<%
Response.CharSet="utf-8"
Session.codepage="65001"
Response.codepage="65001"
Response.ContentType="text/html;charset=utf-8"
%>
```

# ASP 파일을 include 할 때

+ `<!--#include file = "../........."` -->  뭐 요런식으로 including 할 때,

    `Include 파일은 부모 디렉터리를 표시하기 위해 '..'를 사용할 수 없습니다.`

    위처럼 부모 디렉터리를 표시하기 위해 땡땡을 사용할 수 없다는 에러메세지가 나올 때는

    IIS 에서 웹사이트 속성 -> 홈디렉토리 -> 구성 -> 옵션 탭에 가서

    '부모 경로 사용' 체크를 해주면 해결이 된다.

+ 비단, include file 에서만 나오는 것이 아니라, 아래와 같이 스크립트 소스를 참조 하는 경우에도 나오는 에러메세지 이다.

    ```
    <Script src="../function.vbs" language="VBScipt" runat="Server"></script>
    ```
 
    요런식의 스크립트 참조에서도 제목과 같은 에러가 발생 할 수 있다.

# ASP 기본문법

|선언 및 생성|------|
|---|------|
|Dim|변수를 선언한다|
|Set|객체를 생성한다|

|논리 구문(조건문, 반복문)|------|
|---|------|
|IF ~ TEHN|만약 ~라면(조건이 적은 경우에 사용)|
|SELECT CASE|만약 ~ 라면(조건이 많은 경우에 사용)|
|FOR ~ NEXT|반복하며 실행(반복 횟수를 정확히 알고 있을 때 사용)|
|DO WHILE|반복하며 실행(반복 횟수를 정확하게 알 수 없을 때 사용)|


# 선언 및 생성

# 1. Dim(Dimension: 차원)

변수를 선언할 때 사용한다. 변수는 '변하는 수'로, 그 안에 들어갈 값은 상황에 따라 변하게 된다.

의미 없는 변수는 프로그램을 분석하는 데 상당한 장애 요소가 될 수 있기때문에 의미없는 변수를 선언하지 않는 것이 좋다.

대부분의 다른 프로그램(C, C++,JAVA)에서는 변수를 선언하지 않고 사용하면 오류가 발생하지만, ASP는 기본적으로 변수 선언을 하지 않고 변수를 사용해도 오류를 발생시키지 않는다.(옵션 수정으로 오류가 발생하도록 설정 가능하다고 한다.)

```
Dim a 
Dim b
Dim c,d,e,f
Dim strName, intNumer
```

# 2. Set

'객체(Object)의 인스턴스(Instance)'를 생성할 때 사용하는 명령어이다.

객체는 추상적인 의미이며, 직접 사용할 수 없는 존재이고, 인스턴스 구체적이며 직접 사용할 수 있는 존재이다.


객체가 기본적으로 가지고 있는 항목은 '속성(Property)'이며 '객체의 특정한 면들을 설명'하는 용도로 사용한다.


객체가 '수행할 수 있는 작업'은 메서드(Method)라고 한다.

```
Dim myRay   '변수 선언
Set myRay = Server.CreateObject("Vehicle.car")  'car 객체의 인스턴스를 생성,myRay에 저장.

myRay.color = "ivory"  '속성에 값을 할당.
myRay.accel = "good"

myRay.drive("slowly")  '메서드 실행.
myRay.off()

Ser myRay = Nothing  '인스턴스인 myRay를 메모리에서 해제.
```

# 논리 구문

## 조건문

## 1. IF ~ THEN

몇가지 조건이 있을 때, 각각의 조건에 따라 다른 과정을 거칙게 하고 싶을 때 사용하는 구문이다.

```
    <%
        '조건문 : if-elseif문
        dim num
        num=35
        if num=1 THEN
            Response.Write "1입니다"
        ELSEIF num=2 THEN
            Response.Write "2"
        ELSE
            Response.Write "1,2가 아닌 다른수"
        END IF
        %>
        <HR>
        <%if Hour(Now) < 12 THEN%>
            지금은 오전 <%=HOUR(NOW)%>시
        <%ELSE%>
            지금은 오후 <%=HOUR(NOW)-12%>시
        <% END IF %>
```

## 2. SELECT CASE

같은 조건 구문이지만, '조건이 많을 때' 주로 사용하며 간결하고 사용이 편리하다.

CASE 2,3,4,5 처럼 조건이 중복되면 콤마(,)를 이용하면 된다.

```
        <% 
        '조건문 : select-case문
        dim intNum
        intNum = 40
        SELECT CASE intNum
        CASE 1
            Response.Write "1라는 수"
        CASE 2,3,4,5
            Response.Write "2~5사이 수"
        CASE ELSE
            Response.Write "1,2~5와는 다른 수"
        END SELECT
        %>
        <HR>
```

## 반복문

## 1. FOR ~ NEXT

동일한 형식의 작업을 반복하려 할 때 주로 사용하는 구문이다.

일반적으로 시작 부분에 작은수, 끝 부분에 큰 수를 넣고, 다음 1씩 증가하면서 실행하는 방식이다.

FOR ~ NEXT는 반복 실행 횟수를 정확하게 알고 있을 때 주로 사용한다.

```
        <%
        '반복문 : FOR~NEXT
        dim intLoop
        For intLoop=1 to 30 step 1 'step생략 가능(1씩 증가)
            Response.Write intLoop&" "
        Next
        %>
```

## 2. DO WHILE

FOR ~ NEXT문과 마찬가지로 동일한 형식의 반복 작업시 사용한다. 하지만 이는 실행 횟수를 정확하게 모르고 있을 때, 또는 순환할 때마다 조건을 검사해서 계속 순환할 지 여부를 체크해야 할 때 사용한다.

이 반복문을 이용할 땐, '무한 루프(무한하게 계속되는 순환문)'에 빠지지 않도록 주의해야한다.

```
        <%
        '반복문 : DO-WHILE
        dim intLp
        intLp = 1
        Do While intLp <=40
            Response.Write intLp &" "
            intLp = intLp + 1  '1씩 증가한다.
        Loop
        %>
```

|데이터 타입|------|
|---|------|
|숫자 서브 타입|정수나 분수, 또는 부동 소수점 등 5개의 타입이 있다|
|문자열 서브 타입|텍스트 정보를 보관한다|
|날짜 서브 타입|날짜와 시간을 보관한다. 미리 정해진 타입이다|
|불린 서브 타입|참/거짓을 의미하는 TRUE/FALSE 값 중 하나를 가진다|
|그 외의 서브 타입|Empty, Null 등이 있다|

# 1. 숫자와 서브 타입

- 숫자가 들어가는 변수이다. 숫자의 크기에 따라 여섯 가지 타입으로 나눌 수 있다.

### 정수(Integer)

숫자 타입 중에서 가장 일반적으로 쓰이는 타입이다. `-32,768 ~ 32,767(2의 15승 - 1)`사이의 정수.

### 바이트(Byte)

이 타입은 기본적인 숫자의 연산에 사용되는 타입이다. `0 ~ 255(2의 8승 - 1)`까지의 정수. 

### 긴 정수(Long)

Integer보다 훨씬 더 큰 수를 지원하는 타입이다. `-2,147,483,648~2,147,483,647(2의 31승 - 1)`사이의 정수.

### 싱글(Single)

`-3.402823E38 ~ -1.401298E-45`까지의 음수, 그리고 `1.401298E-45 ~ 3.402823E38`까지의 양수 단정도 부동소수점(單精度 浮動小數點)을 지원하는 타입이다.

### 더블(Double)

음수에 대해서는 `-1.79769313486232E308 ~ -4.94065645841247E-324`,

양수에 대해서는 `4.94065645841247E-324 ~ 1.79769313486232E308`까지의 배정도 부동소수점(倍精度 浮動小數點)을 지원합니다.

### 통화(Currency)

'화폐 단위' 를 의미하는 데이터 타입이다. 이 타입은 소숫점 네자리까지의 숫자를 지원한다. 지원하는 범위는 `-922,337,203,685,477.5808 ~ 922,337,203,685,477.5807` 사이 이다.

# 2. 문자열 서브 타입

- '문자열'을 보관하는 변수 타입이다. 문자열 서브 타입을 정의하기 위해서는 문자열 앞뒤에 큰 따옴표(")를 붙이면 된다. 아래의 코드에서 확인할 수 있듯, 큰 따옴표 안에 숫자가 존재하기 때문에 '문자열 서브 타입'이 되어 두 문자열이 합쳐져 "191"이 아닌 "13952"가 되는 것이다.

```
        <%
        '문자열 서브 타입
        Dim strFirst, strSecond, strSum
        strFirst = "139"
        strSecond = "52"
        strSum = strFirst + strSecond

        Response.Write "strFirst > " &strFirst &"<BR>"
        Response.Write "strSecond > " &strSecond &"<BR>"
        Response.Write "strSum > " &strSum &"<BR>"
        %>
```

# 3. 날짜 서브 타입

- 표현하고자 하는 날짜의 앞뒤에  '#' 기호로 둘러싸면 '날짜 정보'를 나타내는 타입으로 이용할 수 있다.

```
        <%
        '날짜 서브 타입
        Response.Write "'#'을 이용한 경우 > " &#7/9/2019# &"<BR>"
        Response.Write "'#'을 이용하지 않은 경우 > " & 7/9/2019 &"<BR>"
        Response.Write "'큰 따옴표'을 이용한 경우 > " &"7/9/2019" &"<BR>"
        %>
```

# 4. 부울린 서브 타입

- 참(TRUE), 거짓(FALSE) 둘 중 하나의 값만 가질수 있는 타입이다.이 값을 ASP 상에서 정수형으로 변환할 수 있다. `TRUE(-1), FALSE(0)`의 값을 가지게 된다. - 이는 IF ~ THEN구문에서 자주 사용된다.

```
        <%
        '부울린 서브 타입
        Response.Write "TRUE: " &TRUE &"<BR>"
        Response.Write "FALSE: " &FALSE &"<BR>"
        %>
```

# 5. 그 외의 서브 타입

- Empty

'값을 전혀 가지고 있지 않다'는 의미로 사용된다. ( 0이 아니다!!)

- NULL

'데이터를 가지고 있지 않은 것, 아무것도 아니다.'는 의미로 사용된다.

- Object

- Errror

# ASP 에러모음

- ASP 0100 메모리 부족
- ASP 0101 예기치 않은 오류
- ASP 0102 문자열 입력 필요
- ASP 0103 정수 입력 필요
- ASP 0104 허용되지 않는 작업
- ASP 0105 범위를 벗어난 인덱스
- ASP 0106 형식 불일치
- ASP 0107 스택 오버플로
- ASP 0108 개체 만들기 실패
- ASP 0109 구성원 없음
- ASP 0110 알 수 없는 이름
- ASP 0111 알 수 없는 인터페이스
- ASP 0112 매개 변수 없음
- ASP 0113 스크립트 시간 초과
- ASP 0114 자유 스레드 개체가 아님
- ASP 0115 예기치 않은 오류
- ASP 0116 스크립트 닫기 구분 기호 없음
- ASP 0117 스크립트 닫기 태그 없음
- ASP 0118 개체 닫기 태그 없음
- ASP 0119 Classid 또는 Progid 특성 없음
- ASP 0120 잘못된 Runat 특성
- ASP 0121 개체 태그의 잘못된 범위
- ASP 0122 개체 태그의 잘못된 범위
- ASP 0123 ID 특성 없음
- ASP 0124 언어 특성 없음
- ASP 0125 특성 닫기 구분 기호 없음
- ASP 0126 Include 파일 없음
- ASP 0127 HTML 설명 닫기 구분 기호 없음
- ASP 0128 File 또는 Virtual 특성 없음
- ASP 0129 알 수 없는 스크립트 언어
- ASP 0130 잘못된 파일 특성
- ASP 0131 허용되지 않는 부모 경로
- ASP 0132 컴파일 오류
- ASP 0133 잘못된 ClassID 특성
- ASP 0134 잘못된 ProgID 특성
- ASP 0135 순환적 포함
- ASP 0136 잘못된 개체 인스턴스 이름
- ASP 0137 잘못된 글로벌 스크립트
- ASP 0138 중첩된 스크립트 블록
- ASP 0139 중첩된 개체
- ASP 0140 잘못된 페이지 명령
- ASP 0141 반복된 페이지 명령
- ASP 0142 스레드 토큰 오류
- ASP 0143 잘못된 응용 프로그램 이름
- ASP 0144 초기화 오류
- ASP 0145 새 응용 프로그램 실패
- ASP 0146 새 세션 실패
- ASP 0147 500 Server Error
- ASP 0148 Server Too Busy
- ASP 0149 Application Restarting
- ASP 0150 응용 프로그램 디렉터리 오류
- ASP 0151 변경 알림 오류
- ASP 0152 Security Error
- ASP 0153 스레드 오류
- ASP 0154 HTTP 헤더 쓰기 오류
- ASP 0155 페이지 콘텐트 쓰기 오류
- ASP 0156 헤더 오류
- ASP 0157 버퍼링 설정
- ASP 0158 URL 없음
- ASP 0159 버퍼링 해제
- ASP 0160 로깅 실패
- ASP 0161 데이터 형식 오류
- ASP 0162 쿠키를 수정할 수 없음
- ASP 0163 잘못된 쉼표 사용
- ASP 0164 잘못된 시간 제한 값
- ASP 0165 SessionID 오류
- ASP 0166 초기화되지 않은 개체
- ASP 0167 세션 초기화 오류
- ASP 0168 허용되지 않는 개체 사용
- ASP 0169 개체 정보 없음
- ASP 0170 세션 삭제 오류
- ASP 0171 경로 없음
- ASP 0172 잘못된 경로
- ASP 0173 잘못된 경로 문자
- ASP 0174 잘못된 경로 문자
- ASP 0175 허용되지 않는 경로 문자
- ASP 0176 경로를 찾을 수 없음
- ASP 0177 Server.CreateObject 실패
- ASP 0178 Server.CreateObject 액세스 오류
- ASP 0179 응용 프로그램 초기화 오류
- ASP 0180 허용되지 않는 개체 사용
- ASP 0181 잘못된 스레딩 모델
- ASP 0182 개체 정보 없음
- ASP 0183 빈 쿠키 키
- ASP 0184 쿠키 이름 없음
- ASP 0185 기본 속성 없음
- ASP 0186 인증서 구문 분석 오류
- ASP 0187 개체 추가 충돌
- ASP 0188 허용되지 않는 개체 사용
- ASP 0189 허용되지 않은 개체 사용
- ASP 0190 예기치 않은 오류
- ASP 0191 예기치 않은 오류
- ASP 0192 예기치 않은 오류
- ASP 0193 OnStartPage 실패
- ASP 0194 OnEndPage 실패
- ASP 0195 잘못된 서버 메서드 호출
- ASP 0196 Out of Process 구성 요소를 시작할 수 없음
- ASP 0197 허용되지 않는 개체 사용
- ASP 0198 Server shutting down
- ASP 0199 허용되지 않는 개체 사용
- ASP 0200 범위를 벗어난 'Expires' 특성
- ASP 0201 잘못된 기본 스크립트 언어
- ASP 0202 코드 페이지 없음
- ASP 0203 잘못된 코드 페이지
- ASP 0204 잘못된 CodePage 값
- ASP 0205 변경 알림
- ASP 0206 BinaryRead를 호출할 수 없음
- ASP 0207 Request.Form을 사용할 수 없음
- ASP 0208 일반 Request 컬렉션을 사용할 수 없음
- ASP 0209 TRANSACTION 속성의 잘못된 값
- ASP 0210 구현되지 않은 메서드
- ASP 0211 범위를 벗어난 개체
- ASP 0212 버퍼를 지울 수 없음
- ASP 0214 잘못된 Path 매개 변수
- ASP 0215 ENABLESESSIONSTATE 속성의 잘못된 값
- ASP 0216 MSDTC 서비스를 실행하고 있지 않음
- ASP 0217 개체 태그의 잘못된 범위
- ASP 0218 LCID가 없음
- ASP 0219 잘못된 LCID
- ASP 0220 Requests for GLOBAL.ASA Not Allowed
- ASP 0221 잘못된 @ 명령 지시어
- ASP 0222 잘못된 TypeLib 지정
- ASP 0223 TypeLib를 찾을 수 없음
- ASP 0224 TypeLib를 로드할 수 없음
- ASP 0225 TypeLib를 래핑할 수 없음
- ASP 0226 StaticObjects를 수정할 수 없음
- ASP 0227 Server.Execute 실패
- ASP 0228 Server.Execute 오류
- ASP 0229 Server.Transfer 실패
- ASP 0230 Server.Transfer 오류
- ASP 0231 Server.Execute 오류
- ASP 0232 잘못된 쿠키 지정
- ASP 0233 쿠키 스크립트 소스를 로드할 수 없음
- ASP 0234 잘못된 include 지시어
- ASP 0235 Server.Transfer 오류
- ASP 0236 잘못된 쿠키 지정
- ASP 0237 잘못된 쿠키 지정
- ASP 0238 특성 값 없음
- ASP 0239 파일을 처리할 수 없음
- ASP 0240 스크립트 엔진 예외
- ASP 0241 CreateObject 예외
- ASP 0242 OnStartPage 인터페이스 쿼리 예외
- ASP 0243 Global.asa에 있는 잘못된 METADATA 태그
- ASP 0244 세션 정보를 사용할 수 없음
- ASP 0245 코드 페이지 값 혼용
- ASP 0246 Too many concurrent users. Please try again later.
- ASP 0247 BinaryRead의 잘못된 인수
- ASP 0248 스크립트가 트랜잭션 처리되지 않음. ObjectContext 개체를 사용하려면 이 ASP 파일을 트랜잭션 처리해야 합니다.
- ASP 0249 Request에서 IStream을 사용할 수 없음. Request.Form 컬렉션 또는 Request.BinaryRead를 사용한 다음 Request 개체에서 IStream을 사용할 수 없습니다.
- ASP 0250 잘못된 기본 코드 페이지. 이 응용 프로그램에 지정된 기본 코드 페이지가 잘못되었습니다.
- ASP 0251 Response 버퍼 제한 초과됨. ASP 페이지를 실행하여 Response 버퍼의 구성된 제한이 초과되었습니다
