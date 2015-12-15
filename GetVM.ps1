# Get information on a specific virtual machine.
# Must know the hostname. Entering the IP address of the VM will not work!!!
# Modified Date:
# Version 1.0

Get-VMGuest -VM (read-host "Name") | FT State,IPAddress,OSFullName -AutoSize
