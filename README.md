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
