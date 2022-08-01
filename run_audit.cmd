@echo off
@cd /d U:\scripts

IF EXIST ksc_vmware.xlsx (
    echo "ksc_vmware.xlsx ok"
) ELSE (
    echo "WARNING ksc_vmware.xlsx missing"
	pause
	Exit /B 5
)

IF EXIST ksc_hyperv.xlsx (
    @echo "ksc_hyperv.xlsx ok"
) ELSE (
    @echo "WARNING ksc_hyperv.xlsx missing"
	pause
	Exit /B 5
)

IF EXIST vmware.csv (
    echo "vmware.csv ok"
) ELSE (
    echo "WARNING vmware.csv missing"
	pause
	Exit /B 5
)

IF EXIST vmware_mgmt.csv (
    echo "vmware_mgmt.csv ok"
) ELSE (
    echo "WARNING vmware_mgmt.csv missing"
	pause
	Exit /B 5
)

IF EXIST vmmreport.csv (
    echo "vmmreport.csv ok"
) ELSE (
    echo "WARNING vmmreport.csv missing"
	pause
	Exit /B 5
)




audit_vm.py

echo "Audit is finished"
pause
