USE [Purchasing]
GO
/****** Object:  Table [dbo].[tblOrders]    Script Date: 9/24/2020 5:29:19 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tblOrders](
	[OrdID] [int] IDENTITY(1,1) NOT NULL,
	[Vendor] [varchar](35) NULL,
	[Add1] [varchar](35) NULL,
	[Add2] [varchar](35) NULL,
	[City] [varchar](20) NULL,
	[State] [varchar](2) NULL,
	[Zip] [varchar](15) NULL,
	[Contact] [varchar](35) NULL,
	[Phone] [varchar](25) NULL,
	[Fax] [varchar](15) NULL,
	[AcctNum] [varchar](20) NULL,
	[RefNum] [varchar](35) NULL,
	[Dept] [varchar](35) NULL,
	[Ext] [varchar](10) NULL,
	[OrdDate] [datetime] NULL,
	[DelTo] [varchar](35) NULL,
	[Terms] [varchar](35) NULL,
	[Comments] [varchar](100) NULL,
	[Owner] [varchar](20) NULL,
 CONSTRAINT [PK_tblOrders] PRIMARY KEY CLUSTERED 
(
	[OrdID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[tblVendors]    Script Date: 9/24/2020 5:29:19 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tblVendors](
	[VenID] [int] IDENTITY(1,1) NOT NULL,
	[Vendor] [varchar](35) NULL,
	[Add1] [varchar](35) NULL,
	[Add2] [varchar](35) NULL,
	[City] [varchar](25) NULL,
	[State] [varchar](2) NULL,
	[Zip] [varchar](15) NULL,
	[Contact] [varchar](35) NULL,
	[Phone] [varchar](25) NULL,
	[Fax] [varchar](15) NULL,
	[AcctNum] [varchar](15) NULL,
	[Owner] [varchar](20) NULL,
 CONSTRAINT [PK_tblVendors] PRIMARY KEY CLUSTERED 
(
	[VenID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
) ON [PRIMARY]

GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnHidden', @value=0 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders', @level2type=N'COLUMN',@level2name=N'OrdID'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnOrder', @value=1 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders', @level2type=N'COLUMN',@level2name=N'OrdID'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnWidth', @value=675 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders', @level2type=N'COLUMN',@level2name=N'OrdID'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnHidden', @value=0 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders', @level2type=N'COLUMN',@level2name=N'Vendor'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnOrder', @value=0 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders', @level2type=N'COLUMN',@level2name=N'Vendor'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnWidth', @value=5775 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders', @level2type=N'COLUMN',@level2name=N'Vendor'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnHidden', @value=0 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders', @level2type=N'COLUMN',@level2name=N'Add1'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnOrder', @value=0 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders', @level2type=N'COLUMN',@level2name=N'Add1'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnWidth', @value=3480 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders', @level2type=N'COLUMN',@level2name=N'Add1'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnHidden', @value=0 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders', @level2type=N'COLUMN',@level2name=N'Add2'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnOrder', @value=0 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders', @level2type=N'COLUMN',@level2name=N'Add2'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnWidth', @value=1470 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders', @level2type=N'COLUMN',@level2name=N'Add2'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnHidden', @value=0 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders', @level2type=N'COLUMN',@level2name=N'City'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnOrder', @value=0 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders', @level2type=N'COLUMN',@level2name=N'City'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnWidth', @value=-1 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders', @level2type=N'COLUMN',@level2name=N'City'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnHidden', @value=0 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders', @level2type=N'COLUMN',@level2name=N'State'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnOrder', @value=0 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders', @level2type=N'COLUMN',@level2name=N'State'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnWidth', @value=645 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders', @level2type=N'COLUMN',@level2name=N'State'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnHidden', @value=0 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders', @level2type=N'COLUMN',@level2name=N'Zip'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnOrder', @value=0 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders', @level2type=N'COLUMN',@level2name=N'Zip'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnWidth', @value=1185 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders', @level2type=N'COLUMN',@level2name=N'Zip'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnHidden', @value=0 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders', @level2type=N'COLUMN',@level2name=N'Contact'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnOrder', @value=0 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders', @level2type=N'COLUMN',@level2name=N'Contact'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnWidth', @value=2385 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders', @level2type=N'COLUMN',@level2name=N'Contact'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnHidden', @value=0 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders', @level2type=N'COLUMN',@level2name=N'Phone'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnOrder', @value=0 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders', @level2type=N'COLUMN',@level2name=N'Phone'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnWidth', @value=2040 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders', @level2type=N'COLUMN',@level2name=N'Phone'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnHidden', @value=0 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders', @level2type=N'COLUMN',@level2name=N'Fax'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnOrder', @value=0 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders', @level2type=N'COLUMN',@level2name=N'Fax'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnWidth', @value=-1 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders', @level2type=N'COLUMN',@level2name=N'Fax'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnHidden', @value=0 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders', @level2type=N'COLUMN',@level2name=N'AcctNum'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnOrder', @value=0 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders', @level2type=N'COLUMN',@level2name=N'AcctNum'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnWidth', @value=1590 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders', @level2type=N'COLUMN',@level2name=N'AcctNum'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnHidden', @value=0 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders', @level2type=N'COLUMN',@level2name=N'RefNum'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnOrder', @value=0 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders', @level2type=N'COLUMN',@level2name=N'RefNum'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnWidth', @value=1770 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders', @level2type=N'COLUMN',@level2name=N'RefNum'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnHidden', @value=0 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders', @level2type=N'COLUMN',@level2name=N'Dept'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnOrder', @value=0 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders', @level2type=N'COLUMN',@level2name=N'Dept'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnWidth', @value=1965 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders', @level2type=N'COLUMN',@level2name=N'Dept'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnHidden', @value=0 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders', @level2type=N'COLUMN',@level2name=N'Ext'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnOrder', @value=0 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders', @level2type=N'COLUMN',@level2name=N'Ext'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnWidth', @value=600 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders', @level2type=N'COLUMN',@level2name=N'Ext'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnHidden', @value=0 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders', @level2type=N'COLUMN',@level2name=N'OrdDate'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnOrder', @value=0 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders', @level2type=N'COLUMN',@level2name=N'OrdDate'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnWidth', @value=900 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders', @level2type=N'COLUMN',@level2name=N'OrdDate'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_Format', @value=N'm/d/yy' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders', @level2type=N'COLUMN',@level2name=N'OrdDate'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnHidden', @value=0 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders', @level2type=N'COLUMN',@level2name=N'DelTo'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnOrder', @value=0 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders', @level2type=N'COLUMN',@level2name=N'DelTo'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnWidth', @value=1425 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders', @level2type=N'COLUMN',@level2name=N'DelTo'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnHidden', @value=0 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders', @level2type=N'COLUMN',@level2name=N'Terms'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnOrder', @value=0 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders', @level2type=N'COLUMN',@level2name=N'Terms'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnWidth', @value=-1 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders', @level2type=N'COLUMN',@level2name=N'Terms'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnHidden', @value=0 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders', @level2type=N'COLUMN',@level2name=N'Comments'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnOrder', @value=0 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders', @level2type=N'COLUMN',@level2name=N'Comments'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnWidth', @value=2070 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders', @level2type=N'COLUMN',@level2name=N'Comments'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnHidden', @value=0 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders', @level2type=N'COLUMN',@level2name=N'Owner'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnOrder', @value=0 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders', @level2type=N'COLUMN',@level2name=N'Owner'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnWidth', @value=-1 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders', @level2type=N'COLUMN',@level2name=N'Owner'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DefaultView', @value=0x02 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_Filter', @value=NULL , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_OrderBy', @value=N'tblOrders.Vendor' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_OrderByOn', @value=1 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_Orientation', @value=0x00 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_TableMaxRecords', @value=100000 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblOrders'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnHidden', @value=0 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblVendors', @level2type=N'COLUMN',@level2name=N'VenID'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnOrder', @value=1 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblVendors', @level2type=N'COLUMN',@level2name=N'VenID'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnWidth', @value=705 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblVendors', @level2type=N'COLUMN',@level2name=N'VenID'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnHidden', @value=0 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblVendors', @level2type=N'COLUMN',@level2name=N'Vendor'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnOrder', @value=2 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblVendors', @level2type=N'COLUMN',@level2name=N'Vendor'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnWidth', @value=4215 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblVendors', @level2type=N'COLUMN',@level2name=N'Vendor'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnHidden', @value=0 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblVendors', @level2type=N'COLUMN',@level2name=N'Add1'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnOrder', @value=4 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblVendors', @level2type=N'COLUMN',@level2name=N'Add1'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnWidth', @value=3510 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblVendors', @level2type=N'COLUMN',@level2name=N'Add1'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnHidden', @value=0 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblVendors', @level2type=N'COLUMN',@level2name=N'Add2'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnOrder', @value=5 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblVendors', @level2type=N'COLUMN',@level2name=N'Add2'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnWidth', @value=2610 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblVendors', @level2type=N'COLUMN',@level2name=N'Add2'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnHidden', @value=0 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblVendors', @level2type=N'COLUMN',@level2name=N'City'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnOrder', @value=6 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblVendors', @level2type=N'COLUMN',@level2name=N'City'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnWidth', @value=1860 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblVendors', @level2type=N'COLUMN',@level2name=N'City'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnHidden', @value=0 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblVendors', @level2type=N'COLUMN',@level2name=N'State'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnOrder', @value=7 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblVendors', @level2type=N'COLUMN',@level2name=N'State'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnWidth', @value=645 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblVendors', @level2type=N'COLUMN',@level2name=N'State'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnHidden', @value=0 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblVendors', @level2type=N'COLUMN',@level2name=N'Zip'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnOrder', @value=8 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblVendors', @level2type=N'COLUMN',@level2name=N'Zip'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnWidth', @value=1185 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblVendors', @level2type=N'COLUMN',@level2name=N'Zip'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnHidden', @value=0 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblVendors', @level2type=N'COLUMN',@level2name=N'Contact'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnOrder', @value=9 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblVendors', @level2type=N'COLUMN',@level2name=N'Contact'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnWidth', @value=3450 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblVendors', @level2type=N'COLUMN',@level2name=N'Contact'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnHidden', @value=0 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblVendors', @level2type=N'COLUMN',@level2name=N'Phone'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnOrder', @value=10 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblVendors', @level2type=N'COLUMN',@level2name=N'Phone'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnWidth', @value=2040 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblVendors', @level2type=N'COLUMN',@level2name=N'Phone'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnHidden', @value=0 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblVendors', @level2type=N'COLUMN',@level2name=N'Fax'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnOrder', @value=11 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblVendors', @level2type=N'COLUMN',@level2name=N'Fax'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnWidth', @value=1350 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblVendors', @level2type=N'COLUMN',@level2name=N'Fax'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnHidden', @value=0 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblVendors', @level2type=N'COLUMN',@level2name=N'AcctNum'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnOrder', @value=12 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblVendors', @level2type=N'COLUMN',@level2name=N'AcctNum'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnWidth', @value=-1 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblVendors', @level2type=N'COLUMN',@level2name=N'AcctNum'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnHidden', @value=0 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblVendors', @level2type=N'COLUMN',@level2name=N'Owner'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnOrder', @value=3 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblVendors', @level2type=N'COLUMN',@level2name=N'Owner'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_ColumnWidth', @value=915 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblVendors', @level2type=N'COLUMN',@level2name=N'Owner'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DefaultView', @value=0x02 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblVendors'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_Filter', @value=NULL , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblVendors'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_OrderBy', @value=N'tblVendors.Vendor' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblVendors'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_OrderByOn', @value=1 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblVendors'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_Orientation', @value=0x00 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblVendors'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_TableMaxRecords', @value=10000 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'tblVendors'
GO
