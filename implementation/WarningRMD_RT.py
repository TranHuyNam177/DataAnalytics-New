from implementation import TaskMonitor
from warning import warning_RMD_RT

@TaskMonitor
def WarningRMD_RT():
    warning_RMD_RT.run(True)
