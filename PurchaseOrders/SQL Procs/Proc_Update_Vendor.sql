USE [Purchasing]
GO

/****** Object:  StoredProcedure [dbo].[Proc_Update_Vendor]    Script Date: 9/24/2020 5:08:29 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE [dbo].[Proc_Update_Vendor]
	-- Add the parameters for the stored procedure here
	   @pvintVenID    INT,
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
	UPDATE tblVendors
	SET 
       [Vendor] = @pvchrVendor
      ,[Add1] = @pvchrAdd1
      ,[Add2] = @pvchrAdd2
      ,[City] = @pvchrCity
      ,[State] = @pvchrState
      ,[Zip] = @pvchrZip
      ,[Contact] = @pvchrContact
      ,[Phone] = @pvchrPhone
      ,[Fax] = @pvchrFax
      ,[AcctNum] = @pvchrAcctNum
      ,[Owner] = @pvchrOwner
   WHERE [VenID] = @pvintVenID
END


GO

