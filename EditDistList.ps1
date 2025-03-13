#
# Code created by Dory-Chew-Chan 16/6/23
#
# Adds users to distribution lists. For external users, you can put the whole email e.g. robertsr@thedenbighalliance.org.uk
#
# Import AD Module
Import-Module ActiveDirectory
Import-Module ExchangeOnlineManagement

# Connect to Exchange Online
try {
    Connect-ExchangeOnline
}
catch {
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit
}

#test github

# Import the data from CSV file and assign it to variable
$List = Import-Csv "C:\Users\chewchand\OneDrive - Denbigh School\Ndrive\Dory_Scripts\DistributionList\DistrList.csv" # Change the filename and path to suit


foreach ($Req in $List) {
    # Retrieve UserSamAccountName and Password
    $Action = $Req.Action
    #Write-Host "Action = " $Action
    $Username = $Req.Username
    Write-Host "Username = " $Username
    $Group = $Req.DistGrp
    Write-Host "Group = " $Group

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
try {
    Disconnect-ExchangeOnline -Confirm:$false
}
catch {
    Write-Host $_.Exception.Message -ForegroundColor Red
}