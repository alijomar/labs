# Script to retrieve tools version for virtual machine
# Created by Ali Omar
# Modified Date:
# Version 1.0

get-vm | get-vmguest | select vmname,toolsversion|ft -AutoSize