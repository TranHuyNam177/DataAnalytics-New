from request import *


def listen_batch_job(
    batch_type: str,
    run_time=None,
    wait_time=15,
):
    """
    This function listen daily batch job

    :param run_time: accept datetime object, None or 'now' to run now
    :param batch_type: accept either 'mid' (for mid-day batch) or 'end' (for end-day batch)
    :param wait_time: number of seconds between retries
    """

    if run_time=='now' or run_time is None:
        t = dt.datetime.now()
    else:
        t = run_time

    day = t.day
    month = t.month
    year = t.year

    while True:

        batch_status = pd.read_sql(
            f"""
            SELECT * 
            FROM 
                [batch]
            WHERE
                [batch].[date] = '{year}-{month}-{day}'
            """,
            connect_DWH_CoSo
        )

        if batch_type=='mid':
            n = 1
        elif batch_type=='end':
            n = 2
        else:
            raise ValueError("batch_type must be either 'mid' or 'end'")

        if batch_status.shape[0]>=n:
            break
        else:
            time.sleep(wait_time)


def get_bank_name(x):

    """
    Tìm tên ngân hàng liên kết cuối cùng

    :param x: là account_code hoặc sub_account
    """

    if x is None:
        return 'Không tìm thấy NH liên kết trên VCF0051'

    # xuất hiện chữ cái trong x -> là account_code, không thì sub_account
    if any([x[i] in 'QWERTYUIOPASDFGHJKLZXCVBNM' for i in range(len(x))]):
        result = pd.read_sql(
            f"""
            SELECT TOP 1 [vcf0051].[bank_name] 
            FROM [vcf0051] 
            WHERE [vcf0051].[sub_account] IN (
                SELECT [sub_account] FROM [sub_account]
                WHERE [sub_account].[account_code] = '{x}'
            )
            AND [bank_name] <> N'---' 
            ORDER BY [date] DESC
            """,
            connect_DWH_CoSo
        )
    else:
        result = pd.read_sql(
            f"""
            SELECT TOP 1 [vcf0051].[bank_name] 
            FROM [vcf0051] 
            WHERE [vcf0051].[sub_account] = '{x}'
            AND [bank_name] <> N'---' 
            ORDER BY [date] DESC
            """,
            connect_DWH_CoSo
        )

    if result.empty:
        result = 'Không tìm thấy NH liên kết trên VCF0051'
    else:
        result = result.squeeze()

    return result


def convert_int(x: int):
    x = str(x)
    if len(x)==1:
        x = '0'+x
    return x