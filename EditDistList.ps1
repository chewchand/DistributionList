# Import AD Module
Import-Module ActiveDirectory
Import-Module ExchangeOnlineManagement

# Connect to Exchange Online
Connect-ExchangeOnline

# Import the data from CSV file and assign it to variable
$List = Import-Csv "C:\Dory_Scripts\DistributionList\DistrList.csv" # Change the filename to the correct job role


foreach ($Req in $List) {
    # Retrieve UserSamAccountName and Password
    $Action = $Req.Action
    #Write-Host "Action = " $Action
    $Username = $Req.Username
    #Write-Host "Username = " $Username
    $Group = $Req.DistGrp
    #Write-Host "Group = " $Group

     # Add member to distribution group
    if ($Action -eq "ADD") {
        Write-Host "ADD Action"
        try {
            Add-DistributionGroupMember -Identity $Group -Member $Username -ErrorAction Stop
            Write-Host $Username " is added to " $Group -ForegroundColor Green
        }
        catch {
            Write-Host $_.Exception.Message -ForegroundColor Red
        }
    }
    elseif ($Action -eq "DEL") {
        Write-Host "DEL Action"
        try {
            Remove-DistributionGroupMember -Identity $Group -Member $Username -Confirm:$False
            Write-Host $Username " is removed from " $Group -ForegroundColor Green
        }
        catch {
            Write-Host $_.Exception.Message -ForegroundColor Red
        }
    }
}


# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false