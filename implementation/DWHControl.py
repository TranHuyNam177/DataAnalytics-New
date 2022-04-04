from implementation import TaskMonitor

@TaskMonitor
def DWH_CoSoCheckTodayUpdate():
    from datawarehouse.DWH_CoSo import CheckSyncUpdate
    CheckSyncUpdate.CheckToday('DWH-CoSo').run()

@TaskMonitor
def DWH_PhaiSinhCheckTodayUpdate():
    from datawarehouse.DWH_CoSo import CheckSyncUpdate
    CheckSyncUpdate.CheckToday('DWH-PhaiSinh').run()

@TaskMonitor
def DWH_CoSoCheckBackDateUpdate():
    from datawarehouse.DWH_CoSo import CheckSyncUpdate
    CheckSyncUpdate.CheckBackDate('DWH-CoSo').run()

@TaskMonitor
def DWH_PhaiSinhCheckBackDateUpdate():
    from datawarehouse.DWH_CoSo import CheckSyncUpdate
    CheckSyncUpdate.CheckBackDate('DWH-PhaiSinh').run()
