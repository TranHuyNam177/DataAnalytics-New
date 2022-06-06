from implementation import TaskMonitor

@TaskMonitor
def DWH_NotifySyncStatusToday(db):
    from datawarehouse import NOTIFYSYNCSTATUSTODAY
    NOTIFYSYNCSTATUSTODAY(db)

@TaskMonitor
def DWH_NotifySyncStatusBackDate(db):
    from datawarehouse import NOTIFYSYNCSTATUSBACKDATE
    NOTIFYSYNCSTATUSBACKDATE(db)
