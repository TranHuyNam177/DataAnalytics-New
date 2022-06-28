from implementation import TaskMonitor


@TaskMonitor
def PR_EmailHoTroMoiGioi():
    from automation.product import EmailHoTroMoiGioi
    EmailHoTroMoiGioi.run(send_mail=True)

