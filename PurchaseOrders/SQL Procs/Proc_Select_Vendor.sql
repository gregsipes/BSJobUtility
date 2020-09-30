USE [Purchasing]
GO

/****** Object:  StoredProcedure [dbo].[Proc_Select_Vendor]    Script Date: 9/24/2020 5:07:45 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE [dbo].[Proc_Select_Vendor]
	-- Add the parameters for the stored procedure here
	@pvchrVendorName  VARCHAR(255) = null
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

    -- Insert statements for procedure here
SELECT [VenID]
      ,[Vendor]
      ,[Add1]
      ,[Add2]
      ,[City]
      ,[State]
      ,[Zip]
      ,[Contact]
      ,[Phone]
      ,[Fax]
      ,[AcctNum]
      ,[Owner]
  FROM [Purchasing].[dbo].[tblVendors]
  WHERE (@pvchrVendorName is null or @pvchrVendorName = Vendor)
  AND (Vendor <> '' AND Vendor is not NULL)
  ORDER BY Vendor
END

GO

