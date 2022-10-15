USE [SPD_MRP]
GO

/****** Object:  StoredProcedure [BookingConferenceRooms].[GetBookingsByDate]    Script Date: 15.10.2022 16:51:18 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO



CREATE PROCEDURE [BookingConferenceRooms].[GetBookingsByDate]
	 @RoomID bigint
	,@Date	date
AS

SET NOCOUNT ON;

IF OBJECT_ID('tempdb..#schedule') IS NOT NULL  DROP TABLE #schedule;

DECLARE 
	 @startTime time = '06:30:00'
	,@endTime	time = '21:00:00';

DECLARE
	 @startOfDay	datetime = DATEADD(DAY, DATEDIFF(DAY, '19000101', @Date), CAST(@startTime as nvarchar(10)))
	,@endOfDay		datetime = DATEADD(DAY, DATEDIFF(DAY, '19000101', @Date), CAST(@endTime as nvarchar(10)));


CREATE TABLE #schedule (
	 [ID] int IDENTITY(1,1)
	,[Time] time
);

WITH
calendar ([date])
AS (
	SELECT
		CAST(@Date as datetime)

	UNION ALL

	SELECT
		DATEADD(mi, 30, [date])
	FROM
		calendar
	WHERE
		[date] < @endOfDay
)
INSERT INTO #schedule (
	[Time]
)
SELECT
	CAST([date] as time)
FROM 
	calendar
WHERE
	CAST([date] as time) > @startTime AND CAST([date] as time) < @endTime
OPTION 
	(maxrecursion 0);


WITH bookingsOfDay
AS (
	SELECT
		 b.ID
		,b.[Datetime start]
		,b.[Datetime end]
		,DATEDIFF(minute, b.[Datetime start], b.[Datetime end]) as [Duration]
		,b.[ID Room]
		,r.[Name]
		,b.Note
	FROM
		[BookingConferenceRooms].[Bookings] b
		LEFT JOIN [BookingConferenceRooms].[Rooms] r
			ON b.[ID Room] = r.ID
		WHERE
			b.[Datetime start] BETWEEN @startOfDay AND @endOfDay
			AND b.[ID Room] = @RoomID
),
bookingSchedule
AS (
	SELECT
		 s.[Time]
		,b.Note
	FROM
		#schedule s
		LEFT JOIN bookingsOfDay b
			ON s.[Time] BETWEEN CAST(b.[Datetime start] as time) AND CAST(b.[Datetime end] as time)
)
SELECT
	[Note]
FROM 
	bookingSchedule
ORDER BY
	[Time] ASC;

GO