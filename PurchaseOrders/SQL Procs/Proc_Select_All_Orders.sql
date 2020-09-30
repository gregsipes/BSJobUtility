USE [Purchasing]
GO

/****** Object:  StoredProcedure [dbo].[Proc_Select_All_Orders]    Script Date: 9/24/2020 5:06:34 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE [dbo].[Proc_Select_All_Orders] 
	-- Add the parameters for the stored procedure here
	@pvintOrderID INT = null,
	@pvintLookbackInYears INT = 5

AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

    -- Insert statements for procedure here
SELECT tblOrders.[OrdID]
      ,tblOrders.[Owner]
      ,tblOrders.[OrdDate]
      ,tblOrders.[Vendor]
      ,tblOrderDetails.[Qty]
      ,tblOrderDetails.[Description]
      ,tblOrderDetails.[UnitPrice]
	  ,tblOrderDetails.Qty * tblOrderDetails.UnitPrice AS TotalPrice
      ,tblOrderDetails.[Taxable]
      ,tblOrders.[Add1] as AddrLine1
      ,tblOrders.[Add2] as AddrLine2
      ,tblOrders.[City]
      ,tblOrders.[State]
      ,tblOrders.[Zip]
      ,tblOrders.[Contact]
      ,tblOrders.[Phone]
      ,tblOrders.[Fax]
      ,tblOrders.[AcctNum]
      ,tblOrders.[RefNum]
      ,tblOrders.[Dept]
      ,tblOrders.[Ext]
      ,tblOrders.[DelTo]
      ,tblOrders.[Terms]
      ,tblOrders.[Comments]
	  ,tblOrderDetails.[RecID] as ItemRecordID
      ,tblOrderDetails.[Qty]
      ,tblOrderDetails.[Units]
      ,tblOrderDetails.[ChargeTo]
      ,tblOrderDetails.[Purpose]
      ,tblOrderDetails.[Class]
  FROM [Purchasing].[dbo].[tblOrders]
  INNER JOIN tblOrderDetails on tblOrderDetails.OrdID = tblOrders.OrdID
  WHERE (@pvintOrderID is null or @pvintOrderID = tblOrders.OrdID)
  AND (tblOrders.OrdDate >= DATEADD(year, -@pvintLookbackInYears, GETDATE()))
  ORDER BY OrdDate DESC
END

--proc_select_all_orders

GO

