USE [Purchasing]
GO

/****** Object:  StoredProcedure [dbo].[Proc_Update_Order_Item]    Script Date: 9/24/2020 5:08:16 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE [dbo].[Proc_Update_Order_Item]
	-- Add the parameters for the stored procedure here
	@pvintOrdID        INT,   
    @pvintRecID        INT = null,   
	@pvintQty          INT = null,
	@pvchrUnits        VARCHAR(20) = null,
	@pvchrDescription  VARCHAR(150) = null,
	@pvcurUnitPrice    MONEY = null,
	@pvchrChargeTo     VARCHAR(35) = null,
	@pvchrPurpose	   VARCHAR(50) = null,
	@pvchrClass	       VARCHAR(35) = null,
	@pvchrTaxable	   VARCHAR(10) = null,
	@pvchrOwner        VARCHAR(20) = null

AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

    -- Insert statements for procedure here
	IF (@pvintRecID = 0 or @pvintRecID is null)
	BEGIN
		INSERT INTO tblOrderDetails ([OrdID] ,[Qty] ,[Units] ,[Description] ,[UnitPrice] 
		,[ChargeTo] ,[Purpose] ,[Class] ,[Taxable] ,[Owner])
		VALUES (@pvintOrdID, @pvintQty, @pvchrUnits, @pvchrDescription, @pvcurUnitPrice,
		@pvchrChargeTo, @pvchrPurpose, @pvchrClass, @pvchrTaxable, @pvchrOwner)

		SELECT SCOPE_IDENTITY() AS RecID
	END
	ELSE
	BEGIN
	    UPDATE tblOrderDetails
		SET [OrdID] = @pvintOrdID
		,[Qty] = @pvintQty
        ,[Units] = @pvchrUnits
        ,[Description] = @pvchrDescription
        ,[UnitPrice] = @pvcurUnitPrice
        ,[ChargeTo] = @pvchrChargeTo
        ,[Purpose] = @pvchrPurpose
        ,[Class] = @pvchrClass
        ,[Taxable] = @pvchrTaxable
        ,[Owner] = @pvchrOwner
	    WHERE RecID = @pvintRecID

		SELECT @pvintRecID AS RecID
	END
END

GO

