/* Microsoft SQL Server - Scripting			*/
/* Server: Taeyo				*/
/* Database: Community					*/
/* Creation Date 2000-12-21 오후 11:37:33 			*/
use community

if exists (select * from sysobjects where id = object_id(N'[dbo].[Queue]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Queue]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[CircleMember]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CircleMember]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[Circle]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Circle]
GO


CREATE TABLE [dbo].[Queue] (
	[Que_code] [int] IDENTITY (1, 1) NOT NULL ,
	[Que_boss] [varchar] (10) NOT NULL ,
	[Que_title] [varchar] (50) NOT NULL ,
	[Que_date] [datetime] NOT NULL ,
	[Que_member] [varchar] (500) NOT NULL ,
	[Que_memcount] [smallint] NOT NULL ,
	[Que_description] [varchar] (500) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Circle] (
	[Cc_code] [int] IDENTITY (1, 1) NOT NULL ,
	[Cc_boss] [varchar] (10) NOT NULL ,
	[Cc_title] [varchar] (50) NOT NULL ,
	[Cc_date] [datetime] NOT NULL ,
	[Cc_memcount] [smallint] NOT NULL ,
	[Cc_BoardName] [varchar] (50) NOT NULL UNIQUE ,
	[Cc_description] [varchar] (500) NULL 
)
GO

CREATE TABLE [dbo].[CircleMember] (
	[Seq_id] [int] IDENTITY (1, 1) NOT NULL ,
	[Cc_BoardName] [varchar] (50) NOT NULL ,
	[mem_id] [varchar] (10) NOT NULL 
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[Circle] WITH NOCHECK ADD 
	CONSTRAINT [PK_Circle] PRIMARY KEY  CLUSTERED 
	(
		[Cc_code]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Queue] WITH NOCHECK ADD 
	CONSTRAINT [DF_Queue_Que_date] DEFAULT (getdate()) FOR [Que_date],
	CONSTRAINT [DF_Queue_Que_memcount] DEFAULT (1) FOR [Que_memcount],
	CONSTRAINT [PK_Queue] PRIMARY KEY  CLUSTERED 
	(
		[Que_code]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Circle] WITH NOCHECK ADD 
	CONSTRAINT [DF_Circle_Cc_date] DEFAULT (getdate()) FOR [Cc_date]
GO

ALTER TABLE [dbo].[CircleMember] WITH NOCHECK ADD 
	CONSTRAINT [PK_CircleMember] PRIMARY KEY  CLUSTERED 
	(
		[Seq_id]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[CircleMember] WITH NOCHECK ADD 
	CONSTRAINT [FK_CircleMember_Circle] FOREIGN KEY (Cc_BoardName)
	REFERENCES Circle (Cc_BoardName)
GO

