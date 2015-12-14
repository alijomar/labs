# Get VM Name with its IP Address
# Created by Ali Omar
# Modified Date: 
# Version 1.0

Get-VM | Select Name, @{N="IP Address";E={@($_.guest.IPAddress[0])}}
