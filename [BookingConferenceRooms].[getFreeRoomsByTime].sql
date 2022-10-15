USE [SPD_MRP]
GO

/****** Object:  StoredProcedure [BookingConferenceRooms].[getFreeRoomsByTime]    Script Date: 15.10.2022 16:51:46 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


CREATE PROCEDURE [BookingConferenceRooms].[getFreeRoomsByTime]
	 @DatetimeStart datetime
	,@DatetimeEnd	datetime
AS

SET NOCOUNT ON;

SELECT
	[Name]
FROM (
	SELECT
		[Name]
	FROM
		[BookingConferenceRooms].Rooms

	EXCEPT

	SELECT DISTINCT 
		r.[Name]
	FROM
		[BookingConferenceRooms].Bookings b
		LEFT JOIN [BookingConferenceRooms].Rooms r
			ON b.[ID Room] = r.ID
	WHERE
		(@DatetimeStart BETWEEN b.[Datetime start] AND b.[Datetime end]
		OR @DatetimeEnd BETWEEN b.[Datetime start] AND b.[Datetime end]
		OR b.[Datetime start] BETWEEN @DatetimeStart AND @DatetimeEnd)
	) a
ORDER BY 
	[Name] ASC
GO