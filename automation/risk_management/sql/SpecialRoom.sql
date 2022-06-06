WITH
[Info] AS (
    SELECT
        [r].[sub_account] [TieuKhoan],
        [r].[account_code] [TaiKhoan],
        [broker].[broker_name] [TenMoiGioi],
        CASE [r].[branch_id]
            WHEN '0001' THEN N'Headquarter'
            WHEN '0101' THEN N'Dist.03'
            WHEN '0102' THEN N'PMH T.F'
            WHEN '0104' THEN N'Dist.07'
            WHEN '0105' THEN N'Tan Binh'
            WHEN '0111' THEN N'Institutional Business'
            WHEN '0113' THEN N'IB'
            WHEN '0117' THEN N'Dist.01'
            WHEN '0118' THEN N'AMD-03'
            WHEN '0119' THEN N'Institutional Business 02'
            WHEN '0201' THEN N'Ha Noi'
            WHEN '0202' THEN N'Thanh Xuan'
            WHEN '0203' THEN N'Cau Giay'
            WHEN '0301' THEN N'Hai Phong'
            ELSE [branch].[branch_name]
        END [Location]
    FROM [relationship] [r]
    LEFT JOIN [branch] ON [branch].[branch_id] = [r].[branch_id]
    LEFT JOIN [broker] ON [broker].[broker_id] = [r].[broker_id]
    WHERE [r].[date] = <dataDate>
        AND EXISTS (
            SELECT [t].[sub_account]
            FROM [vcf0051] [t]
            WHERE [t].[sub_account] = [r].[sub_account]
                AND [t].[date] = [r].[date]
                AND [t].[contract_type] LIKE N'MR%'
        )
)
SELECT
    [Info].[TaiKhoan] + [vpr0108].[ticker] [TKStock],
    [vpr0108].[room_code] + [vpr0108].[ticker] [Code],
    [vpr0108].[room_code] [MaRoom],
    [Info].[TaiKhoan],
    [vpr0108].[ticker] [MaCK],
    [vpr0108].[total_volume] [SpecialRoom],
    '' [GroupDeal]
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
    AND [vpr0108].[used_volume] <> 0
ORDER BY [Info].[TaiKhoan], [vpr0108].[ticker]
