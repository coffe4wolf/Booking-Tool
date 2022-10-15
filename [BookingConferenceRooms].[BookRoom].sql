USE [SPD_MRP]
GO

/****** Object:  StoredProcedure [BookingConferenceRooms].[BookRoom]    Script Date: 15.10.2022 16:42:11 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO



CREATE PROCEDURE [BookingConferenceRooms].[BookRoom]
	 @RoomID		bigint
	,@DatetimeStart datetime
	,@DatetimeEnd	datetime
	,@Note			nvarchar(255)
AS 

BEGIN TRY
	BEGIN TRAN

	INSERT INTO [BookingConferenceRooms].[Bookings] (
		 [ID Room]
		,[Datetime start]
		,[Datetime end]
		,[Note]
	)
	SELECT
		 @RoomID
		,@DatetimeStart
		,@DatetimeEnd
		,@Note;


	COMMIT TRAN
END TRY

BEGIN CATCH

        IF @@TRANCOUNT > 0
            ROLLBACK TRANSACTION;

        DECLARE 
			 @SQLErrorID bigint
			,@SQLErrorMessage nvarchar(2048);

        SELECT
			 @SQLErrorID = ERROR_NUMBER()
			,@SQLErrorMessage = ERROR_MESSAGE();

		SELECT
			 @SQLErrorID
			,@SQLErrorMessage;

END CATCH
GO