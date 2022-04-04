WITH 
[i] AS (
	SELECT 
		[relationship].[sub_account],
		[relationship].[account_code],
		[account].[customer_name],
		[branch].[branch_name],
		'New' [approve_date],
		[vcf0051].[contract_type],
		[broker].[broker_name]
	FROM
		[relationship]
	LEFT JOIN
		[vcf0051]
		ON [vcf0051].[sub_account] = [relationship].[sub_account]
		AND [vcf0051].[date] = [relationship].[date]
	LEFT JOIN
		[account]
		ON [relationship].[account_code] = [account].[account_code]
	LEFT JOIN
		[branch]
		ON [relationship].[branch_id] = [branch].[branch_id]
	LEFT JOIN
		[broker]
		ON [relationship].[broker_id] = [broker].[broker_id]
	WHERE [relationship].[date] = '__input__[1]'
		AND [vcf0051].[status] IN ('A','B')
		AND [vcf0051].[contract_type] NOT LIKE '%GOLD%'
		AND [vcf0051].[contract_type] NOT LIKE '%SILV%'
		AND [vcf0051].[contract_type] NOT LIKE '%VIP%'
		AND [vcf0051].[contract_type] NOT LIKE '%GOLD%'
		AND [vcf0051].[contract_type] LIKE '%MR%'
),
[t] AS (
	SELECT 
		[sum].[account_code],
		SUM([sum].[fee])/ROUND((DATEDIFF(day,'__input__[0]','__input__[1]')+1)/30,0) [fee]
	FROM (
		SELECT 
			[i].[account_code],
			SUM([trading_record].[fee]) [fee]
		FROM
			[trading_record]
		RIGHT JOIN [i] ON [i].[sub_account] = [trading_record].[sub_account]
		WHERE [trading_record].[date] BETWEEN '__input__[0]' AND '__input__[1]'
		GROUP BY
			[trading_record].[date],
			[i].[account_code]
	) [sum]
	GROUP BY 
		[sum].[account_code]
), 
[p] AS (
	SELECT
		[sum].[account_code],
		SUM([sum].[fee])/ROUND((DATEDIFF(day,'__input__[0]','__input__[1]')+1)/30,0) [fee]
	FROM (
		SELECT
			[i].[account_code],
			SUM([payment_in_advance].[total_fee]) [fee]
		FROM
			[payment_in_advance]
		RIGHT JOIN [i] ON [i].[sub_account] = [payment_in_advance].[sub_account]
		WHERE [payment_in_advance].[date] BETWEEN '__input__[0]' AND '__input__[1]'
		GROUP BY
			[payment_in_advance].[date],
			[i].[account_code]
	) [sum]
	GROUP BY 
		[sum].[account_code]
),
[r] AS (
	SELECT
		[sum].[account_code],
		SUM([sum].[int])/ROUND((DATEDIFF(day,'__input__[0]','__input__[1]')+1)/30,0) [int]
	FROM (
		SELECT
			[i].[account_code],
			SUM([rln0019].[interest]) [int]
		FROM
			[rln0019]
		RIGHT JOIN [i] ON [i].[sub_account] = [rln0019].[sub_account]
		WHERE [rln0019].[date] BETWEEN '__input__[0]' AND '__input__[1]'
		GROUP BY
			[rln0019].[date], 
			[i].[account_code]
	) [sum]
	GROUP BY 
		[sum].[account_code]
), 
[nor] AS (
	SELECT [full].* FROM (
		SELECT
			[i].[account_code],
			[i].[customer_name],
			[i].[branch_name],
			[i].[broker_name],
			[i].[approve_date],
			[i].[contract_type],
			ISNULL([t].[fee],0) + 0.3*(ISNULL([p].[fee],0)+ISNULL([r].[int],0)) [fee]
		FROM [i]
		FULL JOIN [t] ON [t].[account_code] = [i].[account_code]
		FULL JOIN [p] ON [p].[account_code] = [i].[account_code]
		FULL JOIN [r] ON [r].[account_code] = [i].[account_code]
	) [full]
	WHERE [full].[fee] >= 20000000
), 
[anav] AS (
	SELECT 
		[sum].[account_code],
		AVG([sum].[nav]) [nav]
	FROM (
		SELECT 
			[n].[date],
			[i].[account_code],
			SUM([n].[nav]) [nav]
		FROM [i]
		LEFT JOIN (
			SELECT [nav].[date], [nav].[sub_account], [nav].[nav]
			FROM [nav]
			WHERE [nav].[date] BETWEEN '__input__[0]' AND '__input__[1]'
		) [n]
		ON [n].[sub_account] = [i].[sub_account]
		GROUP BY 
			[n].[date], 
			[i].[account_code]
	) [sum]
	GROUP BY
		[sum].[account_code]
)
SELECT 
	[all].[account_code],
	[all].[customer_name],
	[all].[branch_name],
	[all].[approve_date],
	CASE 
		WHEN [all].[fee] < 40000000 THEN 20000000
	ELSE
		40000000
	END [criteria_fee],
	[all].[fee],
	CASE 
		WHEN [all].[fee] < 40000000 THEN [all].[fee] / 20000000
	ELSE
		[all].[fee] / 40000000
	END [pct_fee],
	[all].[nav],
	'Nor Margin' [level],
	CASE 
		WHEN [all].[fee] < 40000000 THEN 'Promote SILV PHS'
	    ELSE 'Promote GOLD PHS'
	END [after_review],
	REPLACE([all].[contract_type],' ','') [rate],
	[all].[broker_name]
FROM (
	SELECT
		[nor].*,
		[anav].[nav]
	FROM [nor]
	LEFT JOIN [anav] ON [anav].[account_code] = [nor].[account_code]
	WHERE [anav].[account_code] IN (SELECT [nor].[account_code] FROM [nor]) 
	-- WHERE is used for optimization on nested loop
) [all]


