WITH 
[DanhMuc] AS (
    SELECT [MaCK] [Ticker]
    FROM [DWH-CoSo].[dbo].[DanhMucChoVayMargin]
    WHERE [Ngay] = (SELECT MAX([Ngay]) FROM [DWH-CoSo].[dbo].[DanhMucChoVayMargin])
),
[SinceDate] AS (
    SELECT MIN([Date]) [Date]
    FROM (
        SELECT DISTINCT TOP 66 [Date]
        FROM [DWH-ThiTruong].[dbo].[DuLieuGiaoDichNgay] 
        WHERE [Date] <= <dataDate>
        ORDER BY [Date] DESC
    ) [x]
),
[AvgVolume] AS (
    SELECT [Ticker], AVG([Volume]) [AvgVolume]
    FROM [DWH-ThiTruong].[dbo].[DuLieuGiaoDichNgay] 
    WHERE [Date] >= (SELECT [Date] FROM [SinceDate])
        AND [Date] <= <dataDate>
    GROUP BY [Ticker]
),
[LastDate] AS (
    SELECT [Ticker],MAX([Date]) [Date]
    FROM [DWH-ThiTruong].[dbo].[DuLieuGiaoDichNgay]
    GROUP BY [Ticker]
),
[LastVolume] AS (
    SELECT [Ticker], [Volume] [LastVolume]
    FROM [DWH-ThiTruong].[dbo].[DuLieuGiaoDichNgay] 
    WHERE EXISTS (
		SELECT [Date] FROM [LastDate] 
		WHERE [LastDate].[Ticker] = [DuLieuGiaoDichNgay].[Ticker] 
			AND [LastDate].[Date] = [DuLieuGiaoDichNgay].[Date]
    )
),
[ResultMarket] AS (
SELECT
    [DanhMuc].[Ticker],
    [AvgVolume].[AvgVolume],
    [LastVolume].[LastVolume]
FROM [DanhMuc]
LEFT JOIN [AvgVolume] ON [AvgVolume].[Ticker] = [DanhMuc].[Ticker]
LEFT JOIN [LastVolume] ON [LastVolume].[Ticker] = [DanhMuc].[Ticker]
),
[RawResultVPR0108] AS (
	SELECT
		CONCAT(
			[vpr0108].[ticker],
			'_',
			SUBSTRING(
				[vpr0108].[room_name],
				CHARINDEX('(',[vpr0108].[room_name])+1,
				CHARINDEX(')',[vpr0108].[room_name])-CHARINDEX('(',[vpr0108].[room_name])-1)
			) [Key],
		[vpr0108].[ticker] [Ticker],
        [vpr0108].[room_code] [MaRoom],
		[relationship].[account_code] [TaiKhoan],
		[230006_ThongTinChung].[NgayHieuLuc] [SetUpDate],
        CASE [relationship].[branch_id]
            WHEN '0001' THEN N'Headquarter'
            WHEN '0101' THEN N'Dist.03'
            WHEN '0102' THEN N'Phu My Hung'
            WHEN '0104' THEN N'Dist.07'
            WHEN '0105' THEN N'Tan Binh'
            WHEN '0111' THEN N'Institutional Business'
            WHEN '0113' THEN N'Internet Broker'
            WHEN '0117' THEN N'Dist.01'
            WHEN '0118' THEN N'AMD-03'
            WHEN '0119' THEN N'Institutional Business 02'
            WHEN '0201' THEN N'Ha Noi'
            WHEN '0202' THEN N'Thanh Xuan'
            WHEN '0203' THEN N'Cau Giay'
            WHEN '0301' THEN N'Hai Phong'
            WHEN '0116' THEN N'AMD01'
            ELSE [branch].[branch_name]
        END [ChiNhanh],
		[vpr0108].[total_volume] [SetUp],
		[vpr0108].[used_volume] [UsedRoom]
	FROM [vpr0108]
	LEFT JOIN [230006_ThongTinChung]
		ON [230006_ThongTinChung].[TenRoom] = [vpr0108].[room_name]
			AND [230006_ThongTinChung].[Ngay] = [vpr0108].[date]
	LEFT JOIN [230006_DanhSachTieuKhoan]
		ON [230006_DanhSachTieuKhoan].[MaHieuRoom] = [230006_ThongTinChung].[MaHieuRoom]
			AND [230006_DanhSachTieuKhoan].[Ngay] = [230006_ThongTinChung].[Ngay]
	LEFT JOIN [relationship] 
		ON [relationship].[sub_account] = [230006_DanhSachTieuKhoan].[TieuKhoan]
			AND [relationship].[date] = [230006_DanhSachTieuKhoan].[Ngay]
	LEFT JOIN [branch] ON [branch].[branch_id] = [relationship].[branch_id]
	WHERE [vpr0108].[date] = <dataDate> AND [room_name] LIKE '%(%)%'
),
[ResultVPR0108] AS (
	SELECT
		[Key],
		[Ticker],
        MAX([MaRoom]) [MaRoom],
		MAX([ChiNhanh]) [ChiNhanh],
		STRING_AGG(SUBSTRING([TaiKhoan],PATINDEX('%[^0]%',SUBSTRING([TaiKhoan],5,6))+4,LEN([TaiKhoan])-PATINDEX('%[^0]%',SUBSTRING([TaiKhoan],5,6))),', ') [TaiKhoan],
		MAX([RawResultVPR0108].[SetUpDate]) [SetUpDate],
		MAX([SetUp]) [SetUp],
		MAX([UsedRoom]) [UsedRoom]
	FROM [RawResultVPR0108]
	WHERE LEN([Key]) - LEN([Ticker]) - 1 <= 4
	GROUP BY [Key],[Ticker]
),
[RawResultRoRieng] AS (
	SELECT
		[ticker_code] [Ticker],
		CONCAT([ticker_code],'_',SUBSTRING([room_name],CHARINDEX('(',[room_name])+1,CHARINDEX(')',[room_name])-CHARINDEX('(',[room_name])-1)) [Key],
        [margin_ratio] [MRRatioRieng],
        [margin_max_price] [MaxPriceRieng]
    FROM [vpr0109]
    WHERE CHARINDEX([ticker_code],[room_name]) > 0
		AND [date] = <dataDate>
        AND [room_code] LIKE 'CL%'
        AND [margin_ratio] > 0
),
[ResultRoRieng] AS (
	SELECT * 
	FROM [RawResultRoRieng]
	WHERE LEN([Key]) - LEN([Ticker]) - 1 <= 4
),
[ResultRoChung] AS (
	SELECT
		[t].[ticker_code] [Ticker],
		[t].[margin_ratio] [MRRatioChung],
		[t].[collateral_ratio] [DPRatioChung],
		[t].[margin_max_price] [MaxPriceChung]
	FROM [vpr0109] [t]
	WHERE [t].[date] = <dataDate>
		AND [t].[room_code] LIKE 'TC01%'
		AND [t].[margin_ratio] > 0
),
[RawResult] AS (
SELECT 
	[ResultVPR0108].[Key],
	[ResultVPR0108].[Ticker],
	[ResultVPR0108].[MaRoom],
	CONCAT([ResultVPR0108].[MaRoom],[ResultVPR0108].[Ticker]) [Code],
	[ResultVPR0108].[ChiNhanh],
	[ResultVPR0108].[TaiKhoan],
	[ResultVPR0108].[SetUpDate],
	[ResultVPR0108].[SetUp],
	[ResultVPR0108].[UsedRoom],
	[ResultRoRieng].[MRRatioRieng],
	[ResultRoRieng].[MaxPriceRieng],
	[ResultRoChung].[MRRatioChung],
	[ResultRoChung].[DPRatioChung],
    [ResultRoChung].[MaxPriceChung],
	CASE
	    WHEN [ResultRoRieng].[MRRatioRieng] IS NULL
	        THEN [ResultRoChung].[MRRatioChung]
	    ELSE [ResultRoRieng].[MRRatioRieng]
	END [MRRatio],
    CASE
	    WHEN [ResultRoRieng].[MaxPriceRieng] IS NULL
	        THEN [ResultRoChung].[MaxPriceChung]
	    ELSE [ResultRoRieng].[MaxPriceRieng]
	END [MaxPrice],
	[ResultMarket].[AvgVolume],
	[ResultMarket].[LastVolume]
FROM [ResultVPR0108]
LEFT JOIN [ResultRoRieng]
	ON [ResultRoRieng].[Key] = [ResultVPR0108].[Key]
LEFT JOIN [ResultRoChung]
	ON [ResultRoChung].[Ticker] = [ResultVPR0108].[Ticker]
LEFT JOIN [ResultMarket]
	ON [ResultMarket].[Ticker] = [ResultVPR0108].[Ticker]
WHERE [ResultVPR0108].[SetUp] > 0
)
SELECT
    *,
    CASE
        WHEN [MRRatio] > [DPRatioChung]
            THEN [MaxPrice] * [MRRatio] / 100 * [SetUp]
        ELSE [MaxPrice] * [DPRatioChung] / 100 * [SetUp]
    END [Outstanding]
FROM [RawResult]
ORDER BY [ChiNhanh],[TaiKhoan],[Ticker]