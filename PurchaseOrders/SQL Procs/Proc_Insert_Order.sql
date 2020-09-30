USE [Purchasing]
GO

/****** Object:  StoredProcedure [dbo].[Proc_Insert_Order]    Script Date: 9/24/2020 5:05:09 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE [dbo].[Proc_Insert_Order]
	-- Add the parameters for the stored procedure here
       @pvchrVendor   VARCHAR = null,
       @pvchrAdd1     VARCHAR = null,
       @pvchrAdd2     VARCHAR = null,
       @pvchrCity     VARCHAR = null,
       @pvchrState    VARCHAR = null,
       @pvchrZip      VARCHAR = null,
       @pvchrContact  VARCHAR = null,
       @pvchrPhone    VARCHAR = null,
       @pvchrFax	  VARCHAR = null,
       @pvchrAcctNum  VARCHAR = null,
       @pvhcrRefNum   VARCHAR = null,
       @pvchrDept     VARCHAR = null,
       @pvchrExt      VARCHAR = null,
       @pvdatOrdDate  DATETIME = null,
       @pvchrDelTo	  VARCHAR = null,
       @pvchrTerms	  VARCHAR = null,
       @pvchrComments VARCHAR = null,
       @pvchrOwner    VARCHAR = null

AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

    -- Insert statements for procedure here
	INSERT INTO tblOrders (
       [Vendor]
      ,[Add1]
      ,[Add2]
      ,[City]
      ,[State]
      ,[Zip]
      ,[Contact]
      ,[Phone]
      ,[Fax]
      ,[AcctNum]
      ,[RefNum]
      ,[Dept]
      ,[Ext]
      ,[OrdDate]
      ,[DelTo]
      ,[Terms]
      ,[Comments]
      ,[Owner]) 
	  VALUES (
       @pvchrVendor,
       @pvchrAdd1,
       @pvchrAdd2,
       @pvchrCity,
       @pvchrState,
       @pvchrZip,
       @pvchrContact,
       @pvchrPhone,
       @pvchrFax,
       @pvchrAcctNum,
       @pvhcrRefNum,
       @pvchrDept,
       @pvchrExt,
       @pvdatOrdDate,
       @pvchrDelTo,
       @pvchrTerms,
       @pvchrComments,
       @pvchrOwner
		)

END

-- Return the new record ID
SELECT SCOPE_IDENTITY() AS OrdID

GO

