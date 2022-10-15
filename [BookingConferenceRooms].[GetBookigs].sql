USE [SPD_MRP]
GO

/****** Object:  StoredProcedure [BookingConferenceRooms].[GetBookigs]    Script Date: 15.10.2022 16:50:48 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


CREATE PROCEDURE [BookingConferenceRooms].[GetBookigs]
AS

SET NOCOUNT ON;

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
GO