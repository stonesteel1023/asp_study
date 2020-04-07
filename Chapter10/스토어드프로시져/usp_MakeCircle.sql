CREATE  PROC usp_MakeCircle 
         @code  int 
AS 
  Begin 
      BEGIN TRAN 
  
         /* Queue ���̺� �ִ� �ӽ� �����͵��� �����ͼ� ���� ������*/ 
	--������ ������ ������ ���ÿ� �ʱ�ȭ ��� 
         Declare @boss  varchar(10)		-- Ŭ�� ��û�ڸ� ���� ����
	 Set @boss = '' 
         Declare @title  varchar(50) 		-- Ŭ�� ������ ���� ����
	 Set @title = '' 
         Declare @Member  varchar(500) 		-- Ŭ�� ����ȸ���� ���� ����
	 Set @Member = '' 
         Declare @memcount smallint	 	-- Ŭ�� ����ȸ�� ���� ���� ����
	 Set @memcount = 0 
         Declare @description varchar(500) 	-- Ŭ�� �Ұ��� ���� ����
	 Set @description = '' 

         /* Queue ���̺� �ִ� Ư�� �����͸� �����ͼ� ������ ������ ��´�*/ 
         select @boss = Que_boss , @title = Que_title , @memcount = Que_memcount, 
          @Member = Que_Member, @description = Que_description  from Queue 
         where Que_code= @code 
  
         /* Circle�� �Խ��� �̸����� table ������ "BD"��� ���ڸ� ���ٿ� �����Ѵ�.*/ 
         Declare @BoardName varchar(25) 	-- �Խ��� �̸��� ������ ����
	 Set @BoardName = '' 
         Set @BoardName = 'Board_' + @boss + convert(varchar(10),@code) 

         /* ������ �����͸� Circle ���̺� �ű��.*/ 
         Insert into Circle 
         ( Cc_boss, Cc_title, Cc_description, Cc_memcount, Cc_BoardName) values 
         ( @boss, @title, @description, @memcount, @BoardName ) 
  
         /* ������ ���鼭 ������ ���̵� CircleMember ���̺� Insert �Ѵ�*/ 
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
  
         /* Queue ���̺� �ִ� Ư�� �����͸� �����Ѵ�*/ 
         Delete from Queue where Que_code = @code 
  
         /* �������� ���̺��� �����ϱ� ���ؼ� �� Create �� ���ڿ��� �����.*/ 
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
