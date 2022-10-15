USE [SPD_MRP]
GO

/****** Object:  StoredProcedure [BookingConferenceRooms].[CreateBookingsCalendarPivot]    Script Date: 15.10.2022 16:49:12 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


CREATE PROCEDURE [BookingConferenceRooms].[CreateBookingsCalendarPivot]
AS

DECLARE 
	 @startTime time = '06:30:00'
	,@endTime	time = '21:00:00';

DECLARE
	 @startOfDay	datetime = DATEADD(DAY, DATEDIFF(DAY, '19000101', '20220801'), CAST(@startTime as nvarchar(10)))
	,@endOfDay		datetime = DATEADD(DAY, DATEDIFF(DAY, '19000101', '20220801'), CAST(@endTime as nvarchar(10)));

DECLARE 
	@email nvarchar(120) = N'mailto:spd-moscow@spd.ru';

IF OBJECT_ID('tempdb..#schedule') IS NOT NULL  DROP TABLE #schedule;
CREATE TABLE #schedule (
	 [ID] int IDENTITY(1,1)
	,[Datetime] datetime
	,[Date] date
	,[Time] time
);

WITH
calendar ([date])
AS (
	SELECT
		CAST('20220801' as datetime)

	UNION ALL

	SELECT
		DATEADD(mi, 30, [date])
	FROM
		calendar
	WHERE
		[date] < '20221231'
)
INSERT #schedule (
	 [Datetime]
	,[Date]
	,[Time]
)
SELECT
	 CAST([date] as datetime) as [Datetime]
	,CAST([date] as date) as [Date]
	,CAST([date] as time) as [Time]
FROM 
	calendar
WHERE
	CAST([date] as time) > @startTime AND CAST([date] as time) < @endTime
OPTION 
	(maxrecursion 0);

WITH schedule 
AS (
	SELECT 
		 [Datetime]
		,[Date]
		,[Time]
		,r.ID		as [Room ID]
		,r.[Name]	as [Room name]
	FROM
		#schedule s
		CROSS JOIN BookingConferenceRooms.Rooms r
	),
result 
AS (
	SELECT 
		 [Datetime]
		,[Date]
		,[Time]
		,[Room ID]
		,[Room name]
		,CASE WHEN b.Note IS NOT NULL THEN N'Занято' ELSE @email + N'?cc=;&subject=Бронирование переговорной "' + [Room name] + N'"&body=Прошу забронировать переговорную "' + [Room name] + '" ' + FORMAT([Date], 'dd.MM') + ' c ' + CAST(FORMAT([Datetime], 'HH:mm') as nvarchar(11)) + N' по ЗАПОЛНИТЬ' end as [Note]
	FROM
		schedule s
		LEFT JOIN BookingConferenceRooms.Bookings b
			ON s.[Datetime] BETWEEN b.[Datetime start] AND b.[Datetime end]
			AND s.[Room ID] = b.[ID Room]
	)
SELECT 
	 [Datetime]
	,MAX(CASE WHEN [Room name] = N'Нева' THEN Note ELSE NULL END)	[Нева]
	,MAX(CASE WHEN [Room name] = N'Ока' THEN Note ELSE NULL END)	[Ока]
	,MAX(CASE WHEN [Room name] = N'Тура' THEN Note ELSE NULL END)	[Тура]
	,MAX(CASE WHEN [Room name] = N'Ишим' THEN Note ELSE NULL END)	[Ишим]
	,MAX(CASE WHEN [Room name] = N'Кама' THEN Note ELSE NULL END)	[Кама]
	,MAX(CASE WHEN [Room name] = N'Ачимов' THEN Note ELSE NULL END) [Ачимов]
	,MAX(CASE WHEN [Room name] = N'Печора' THEN Note ELSE NULL END) [Печора]
	,MAX(CASE WHEN [Room name] = N'Урал' THEN Note ELSE NULL END)	[Урал]
	,MAX(CASE WHEN [Room name] = N'Лена' THEN Note ELSE NULL END)	[Лена]
	,MAX(CASE WHEN [Room name] = N'Енисей' THEN Note ELSE NULL END) [Енисей]
	,MAX(CASE WHEN [Room name] = N'Тобол' THEN Note ELSE NULL END)	[Тобол]
	,MAX(CASE WHEN [Room name] = N'Обь' THEN Note ELSE NULL END)	[Обь]
	,MAX(CASE WHEN [Room name] = N'Москва' THEN Note ELSE NULL END) [Москва]

FROM
	 result
GROUP BY
	[Datetime];
GO


