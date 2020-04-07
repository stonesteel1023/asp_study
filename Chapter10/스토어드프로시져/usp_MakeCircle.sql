CREATE  PROC usp_MakeCircle 
         @code  int 
AS 
  Begin 
      BEGIN TRAN 
  
         /* Queue 테이블에 있는 임시 데이터들을 가져와서 담을 변수들*/ 
	--가능한 변수는 생성과 동시에 초기화 명시 
         Declare @boss  varchar(10)		-- 클럽 신청자를 담을 변수
	 Set @boss = '' 
         Declare @title  varchar(50) 		-- 클럽 제목을 담을 변수
	 Set @title = '' 
         Declare @Member  varchar(500) 		-- 클럽 참여회원을 담을 변수
	 Set @Member = '' 
         Declare @memcount smallint	 	-- 클럽 참여회원 수를 담을 변수
	 Set @memcount = 0 
         Declare @description varchar(500) 	-- 클럽 소개를 담을 변수
	 Set @description = '' 

         /* Queue 테이블에 있는 특정 데이터를 가져와서 각각의 변수에 담는다*/ 
         select @boss = Que_boss , @title = Que_title , @memcount = Que_memcount, 
          @Member = Que_Member, @description = Que_description  from Queue 
         where Que_code= @code 
  
         /* Circle별 게시판 이름으로 table 변수에 "BD"라는 문자를 덧붙여 지정한다.*/ 
         Declare @BoardName varchar(25) 	-- 게시판 이름을 저장할 변수
	 Set @BoardName = '' 
         Set @BoardName = 'Board_' + @boss + convert(varchar(10),@code) 

         /* 가져온 데이터를 Circle 테이블에 옮긴다.*/ 
         Insert into Circle 
         ( Cc_boss, Cc_title, Cc_description, Cc_memcount, Cc_BoardName) values 
         ( @boss, @title, @description, @memcount, @BoardName ) 
  
         /* 루프를 돌면서 각각의 아이디를 CircleMember 테이블에 Insert 한다*/ 
         Declare @StartCharIdx	smallint 
	 Set @StartCharIdx = '' 
         Declare @EndCharIdx	smallint 
	 Set @EndCharIdx = 0 
         Declare @UserID	Varchar(10) 
	 Set @UserID = '' 
         Set   @StartCharIdx = 1 
  
        While 1 = 1 
        Begin 
           SELECT @EndCharIdx = CHARINDEX(',', @member, @StartCharIdx) 
           if (@EndCharIdx = 0) Break 
              select @UserID = SubString(@member, @StartCharIdx, @EndCharIdx - @StartCharIdx) 
              Insert into CircleMember (Cc_BoardName, mem_id) values (@BoardName, @UserID) 
  
              Set @StartCharIdx = @EndCharIdx+1 
         End 
  
         /* Queue 테이블에 있는 특정 데이터를 삭제한다*/ 
         Delete from Queue where Que_code = @code 
  
         /* 동적으로 테이블을 생성하기 위해서 그 Create 할 문자열을 만든다.*/ 
         Declare @CreateBoard varchar(1000) 
  
         Set @CreateBoard  = '' 
         Set @CreateBoard = @CreateBoard + ' Create table ' + @BoardName 
         Set @CreateBoard = @CreateBoard + ' ( ' 
         Set @CreateBoard = @CreateBoard + '       board_idx int IDENTITY (1, 1) NOT NULL primary key,  ' 
         Set @CreateBoard = @CreateBoard + '       name varchar(20) NOT NULL, ' 
         Set @CreateBoard = @CreateBoard + '       title varchar(40) NOT NULL, ' 
         Set @CreateBoard = @CreateBoard + '       mail varchar(50) NULL, ' 
         Set @CreateBoard = @CreateBoard + '       url varchar(50) NULL, ' 
         Set @CreateBoard = @CreateBoard + '       writeday varchar(30) NOT NULL, ' 
         Set @CreateBoard = @CreateBoard + '       pwd varchar(20) NOT NULL, ' 
         Set @CreateBoard = @CreateBoard + '       ref smallint NOT NULL, ' 
         Set @CreateBoard = @CreateBoard + '       step smallint NOT NULL, ' 
         Set @CreateBoard = @CreateBoard + '       re_level smallint NOT NULL, ' 
         Set @CreateBoard = @CreateBoard + '       readnum smallint NOT NULL, ' 
         Set @CreateBoard = @CreateBoard + '       content text NOT NULL, ' 
         Set @CreateBoard = @CreateBoard + '       part varchar(10) NULL  ' 
         Set @CreateBoard = @CreateBoard + ' ) ' 
  
         Exec(@CreateBoard) 
  
     COMMIT TRAN 
  
     RETURN(1) 
  End 


/*
Drop proc usp_MakeCircle


declare @c  int
exec @c = usp_MakeCircle 3
select @c

*/
