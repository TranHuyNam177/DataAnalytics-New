WITH 
[SinceDate] AS (
    SELECT
        MIN([Date]) [Date]
    FROM (
        SELECT DISTINCT TOP 66 [t].[Date]
        FROM [DWH-ThiTruong].[dbo].[DuLieuGiaoDichNgay] [t]
        WHERE [t].[Date] <= <dataDate>
        ORDER BY [t].[Date] DESC
    ) [x]
),
[AvgVolume] AS (
    SELECT
        [DWH-ThiTruong].[dbo].[DuLieuGiaoDichNgay].[Ticker],
        AVG([DWH-ThiTruong].[dbo].[DuLieuGiaoDichNgay].[Volume]) [AvgVolume]
    FROM [DWH-ThiTruong].[dbo].[DuLieuGiaoDichNgay]
    WHERE [DuLieuGiaoDichNgay].[Date] >= (SELECT [Date] FROM [SinceDate])
    AND [DuLieuGiaoDichNgay].[Date] <= <dataDate>
    GROUP BY [DWH-ThiTruong].[dbo].[DuLieuGiaoDichNgay].[Ticker]
),
[Market] AS (
    SELECT
        [Ticker],
        [Volume],
        [Ref] * 1000 [Ref],
        [Close] * 1000 [ClosePrice],
        [High]
    FROM [DWH-ThiTruong].[dbo].[DuLieuGiaoDichNgay]
    WHERE [Date] = <dataDate>
),
[Info] AS (
    SELECT
        [r].[sub_account] [TieuKhoan],
        [r].[account_code] [TaiKhoan],
        [r].[branch_id] [MaChiNhanh],
        CASE [r].[branch_id]
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
        END [Location]
    FROM [relationship] [r]
    LEFT JOIN [branch] ON [branch].[branch_id] = [r].[branch_id]
    WHERE [r].[date] = <dataDate>
        AND EXISTS (
            SELECT [t].[sub_account] 
            FROM [vcf0051] [t] 
            WHERE [t].[sub_account] = [r].[sub_account] 
                AND [t].[date] = [r].[date]
                AND [t].[contract_type] LIKE N'MR%'
        )
),
[Asset] AS (
    SELECT
        [rmr0015].[sub_account] [TieuKhoan],
        SUM([rmr0015].[market_value]) / 1000000 [TotalAsset]
    FROM [rmr0015]
    WHERE [rmr0015].[date] = <dataDate>
    GROUP BY [rmr0015].[sub_account]
),
[SCRDL] AS (
    SELECT
        [account_code] [TaiKhoan],
        [stock] [Stock],
        [SCR],
        [DL]
    FROM (SELECT * FROM [high_risk_account] WHERE [high_risk_account].[date] = <dataDate>) [t1]
    PIVOT (MAX([t1].[value]) FOR [type] IN (SCR, DL)) [t2]
),
[MR] AS (
    SELECT
        [VPR0109CL01].[ticker_code] [Ticker],
        [VPR0109CL01].[margin_ratio] [MRRatio],
        [VPR0109TC01].[DPRatio],
        [VPR0109CL01].[margin_max_price] [MaxPrice]
    FROM [vpr0109] [VPR0109CL01]
    INNER JOIN (
        SELECT
            [vpr0109].[ticker_code] [Ticker],
            [vpr0109].[margin_ratio] [DPRatio]
        FROM [vpr0109]
        WHERE [vpr0109].[room_code] = 'TC01_PHS'
            AND [vpr0109].[date] = <dataDate>
    ) [VPR0109TC01]
    ON [VPR0109TC01].[Ticker] = [VPR0109CL01].[ticker_code]
    WHERE [VPR0109CL01].[date] = <dataDate>
        AND [VPR0109CL01].[room_code] = 'CL01_PHS'
),
[GeneralRoom] AS (
    SELECT
        [230007].[ticker] [Ticker],
        [230007].[system_total_room] [GeneralRoom]
    FROM [230007]
    WHERE [230007].[date] = <dataDate>
),
[SpecialRoom] AS (
    SELECT 
        [vpr0108].[room_code] [MaRoom],
        [Info].[TaiKhoan],
        [vpr0108].[ticker] [MaCK],
        [vpr0108].[total_volume] [SpecialRoom],
        [230006_ThongTinChung].[NgayHieuLuc]
    FROM [vpr0108]
    INNER JOIN [230006_ThongTinChung] 
        ON [230006_ThongTinChung].[TenRoom] = [vpr0108].[room_name]
        AND [vpr0108].[date] = [230006_ThongTinChung].[Ngay] AND [vpr0108].[date] = <dataDate>
    INNER JOIN [230006_DanhSachTieuKhoan] 
        ON [230006_ThongTinChung].[MaHieuRoom] = [230006_DanhSachTieuKhoan].[MaHieuRoom]
        AND [230006_ThongTinChung].[Ngay] = [230006_DanhSachTieuKhoan].[Ngay]
    LEFT JOIN [Info] ON [Info].[TieuKhoan] = [230006_DanhSachTieuKhoan].[TieuKhoan]
    WHERE EXISTS (
        SELECT [TenRoom] 
        FROM [230006_ThongTinChung], [230006_DanhSachTieuKhoan]
        WHERE [230006_ThongTinChung].[TenRoom] = [vpr0108].[room_name]
            AND [230006_ThongTinChung].[Ngay] = [vpr0108].[date] AND [vpr0108].[date] = <dataDate>
            AND [230006_ThongTinChung].[MaHieuRoom] = [230006_DanhSachTieuKhoan].[MaHieuRoom]
            AND [230006_ThongTinChung].[Ngay] = [230006_DanhSachTieuKhoan].[Ngay]
        )
),
[HighRisk] AS (
    SELECT
        [account_code] [Account],
        [stock] [Stock],
        [quantity] [Quantity],
        [price] [Price],
        [cash] [Cash],
        ([total_outstanding]-[cash])/1000000 [TotalLoan],
        [market_value]/1000000 [MarginValue]
    FROM [high_risk_account]
    WHERE [date] = <dataDate> AND [type] = 'SCR'
),
[RawResult] AS (
    SELECT
        CASE
            WHEN [HighRisk].[Stock] IN <VN30> THEN 'VN30'
            WHEN [HighRisk].[Stock] IN <HNX30> THEN 'HNX30'
            ELSE ''
        END [Index],
        [SpecialRoom].[MaRoom],
        [HighRisk].[Account],
        [Info].[Location],
        [HighRisk].[Stock],
        [HighRisk].[Quantity],
        [HighRisk].[Price],
        [Asset].[TotalAsset],
        [HighRisk].[Cash],
        [HighRisk].[TotalLoan],
        [HighRisk].[MarginValue],
        [SCRDL].[SCR],
        [SCRDL].[DL],
        [MR].[MRRatio],
        [MR].[DPRatio],
        CASE
            WHEN [MR].[MRRatio] > [MR].[DPRatio]
                THEN [MR].[MRRatio]
            ELSE [MR].[DPRatio]
        END [MaxRatio],
        [MR].[MaxPrice],
        [GeneralRoom].[GeneralRoom],
        ISNULL([SpecialRoom].[SpecialRoom],0) [SpecialRoom],
        [SpecialRoom].[NgayHieuLuc],
        ([HighRisk].[TotalLoan]-[Asset].[TotalAsset]+[HighRisk].[MarginValue])*1000000/[HighRisk].[Quantity] [BreakevenPrice],
        [AvgVolume].[AvgVolume] [AvgVolume3M],
        [Market].[Volume],
        [Market].[ClosePrice],
        CASE 
            WHEN [HighRisk].[Price] > [MR].[MaxPrice] THEN [MR].[MaxPrice]
            ELSE [HighRisk].[Price]
        END [MinPrice],
        CASE
            WHEN [HighRisk].[Price] > [MR].[MaxPrice]
                THEN (ISNULL([GeneralRoom].[GeneralRoom],0)+ISNULL([SpecialRoom].[SpecialRoom],0))*[MR].[MaxPrice]*[MR].[MRRatio]/100/1000000000
            ELSE (ISNULL([GeneralRoom].[GeneralRoom],0)+ISNULL([SpecialRoom].[SpecialRoom],0))*[HighRisk].[Price]*[MR].[MRRatio]/100/1000000000
        END [TotalPotentialOutstanding]
    FROM [HighRisk]
    LEFT JOIN [Info] ON [Info].[TaiKhoan] = [HighRisk].[Account] 
    LEFT JOIN [SCRDL] ON [SCRDL].[TaiKhoan] = [HighRisk].[Account]
        AND [SCRDL].[Stock] = [HighRisk].[Stock]
    LEFT JOIN [MR] ON [MR].[Ticker] = [HighRisk].[Stock]
    LEFT JOIN [GeneralRoom] ON [GeneralRoom].[Ticker] = [HighRisk].[Stock]
    LEFT JOIN [Market] ON [Market].[Ticker] = [HighRisk].[Stock]
    LEFT JOIN [Asset] ON [Asset].[TieuKhoan] = [InFo].[TieuKhoan]
    LEFT JOIN [AvgVolume] ON [AvgVolume].[Ticker] = [HighRisk].[Stock]
    LEFT JOIN [SpecialRoom] 
        ON [SpecialRoom].[MaCK] = [HighRisk].[Stock]
        AND [SpecialRoom].[TaiKhoan] = [HighRisk].[Account]
)
SELECT
    [RawResult].*,
    CASE 
        WHEN [RawResult].[MaxPrice]*[RawResult].[MaxRatio] = 0
            THEN 0
        ELSE ([RawResult].[BreakevenPrice]-[RawResult].[MinPrice]*[RawResult].[MaxRatio]/100)/([RawResult].[MaxPrice]*[RawResult].[MaxRatio]/100) 
    END [PctBreakevenPriceMaxPrice],
    1 - [RawResult].[BreakevenPrice] / [RawResult].[ClosePrice] [PctBreakevenPriceMarketPrice],
    [RawResult].[Volume] / [RawResult].[AvgVolume3M] - 1 [VolumeOnAvg3M]
FROM [RawResult]
ORDER BY [RawResult].[Account], [RawResult].[Index], [RawResult].[Stock]
