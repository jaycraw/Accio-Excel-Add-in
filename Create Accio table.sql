SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Accio](
	--Primary keyed such that no user can use the same nickname twice. However, user can give the same workbook multiple nicknames
	[id] [int] IDENTITY(1,1) NOT NULL,
	[Username] [varchar](25) NOT NULL,
	[CommonName] [varchar](50) NOT NULL,
	[FilePath] [varchar](max) NULL,
PRIMARY KEY CLUSTERED 
(
	[Username] ASC,
	[CommonName] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
