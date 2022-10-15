USE [SPD_MRP]
GO

/****** Object:  StoredProcedure [BookingConferenceRooms].[unbookRoom]    Script Date: 15.10.2022 16:52:24 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


CREATE PROCEDURE [BookingConferenceRooms].[unbookRoom]
	 @RoomID bigint
	,@DatetimeStart datetime
	,@DatetimeEnd	datetime
AS

SET NOCOUNT ON;

BEGIN TRY
	BEGIN TRAN

		IF OBJECT_ID('tempdb..#bookingsToDelete') IS NOT NULL  DROP TABLE #bookingsToDelete;

		CREATE TABLE #bookingsToDelete (
			 [RowID] bigint IDENTITY(1,1)
			,[id] bigint
		);

		INSERT INTO #bookingsToDelete ([id])
		SELECT
			b.ID
		FROM
			[BookingConferenceRooms].[Bookings] b
		WHERE
			[ID Room] = @RoomID
			AND
			(@DatetimeStart BETWEEN b.[Datetime start] AND b.[Datetime end]
			OR @DatetimeEnd BETWEEN b.[Datetime start] AND b.[Datetime end]
			OR b.[Datetime start] BETWEEN @DatetimeStart AND @DatetimeEnd);

		DELETE [BookingConferenceRooms].[Bookings]
		WHERE [id] IN (SELECT [id] FROM #bookingsToDelete);

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