USE [Purchasing]
GO

/****** Object:  StoredProcedure [dbo].[Proc_Select_Order_Item]    Script Date: 9/24/2020 5:07:31 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE [dbo].[Proc_Select_Order_Item]
	-- Add the parameters for the stored procedure here
	@pvintOrderID      INT,
	@pvintItemRecordID INT = null

AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

    -- Insert statements for procedure here
SELECT [RecID]
      ,[OrdID]
      ,[Qty]
      ,[Units]
      ,[Description]
      ,[UnitPrice]
      ,[ChargeTo]
      ,[Purpose]
      ,[Class]
      ,[Taxable]
      ,[Owner]
  FROM [Purchasing].[dbo].[tblOrderDetails]
  WHERE OrdID = @pvintOrderID
  AND (@pvintItemRecordID = 0 or @pvintItemRecordID is null or @pvintItemRecordID = RecID)
END

GO

