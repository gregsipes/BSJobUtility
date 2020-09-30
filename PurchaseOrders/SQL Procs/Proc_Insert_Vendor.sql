USE [Purchasing]
GO

/****** Object:  StoredProcedure [dbo].[Proc_Insert_Vendor]    Script Date: 9/24/2020 5:06:16 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE [dbo].[Proc_Insert_Vendor]
	-- Add the parameters for the stored procedure here
       @pvchrVendor   VARCHAR(255) = null,
       @pvchrAdd1     VARCHAR(255) = null,
       @pvchrAdd2     VARCHAR(255) = null,
       @pvchrCity     VARCHAR(255) = null,
       @pvchrState    VARCHAR(255) = null,
       @pvchrZip      VARCHAR(255) = null,
       @pvchrContact  VARCHAR(255) = null,
       @pvchrPhone    VARCHAR(255) = null,
       @pvchrFax	  VARCHAR(255) = null,
       @pvchrAcctNum  VARCHAR(255) = null,
       @pvchrOwner    VARCHAR(255) = null

AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

    -- Insert statements for procedure here
	INSERT INTO tblVendors (
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
       @pvchrOwner
		)

END

-- Return the new record ID
SELECT SCOPE_IDENTITY() AS VendorID

--EXEC Proc_Insert_Vendor
GO

