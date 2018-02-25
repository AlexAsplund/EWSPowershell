#EWS module for powershell
## Usage
```powershell
# Start with creating a ExchangeService session
$Credential = Get-Credential
Invoke-ExchangeService -Credential $Credential -Domain "contoso.com" -AutoDiscoverUrl "user@contoso.com"

# And start running the commands
$AllRooms = Get-AllRooms


$Subject = "Test"
$Body = "Body"
$Resource = "ConferenceRoom@contoso.com"
$StartDate = Get-Date
$EndDate = (Get-Date).AddHours(2)


$AppointmentSplat = @{
    Subject = $Subject
    Body = $Body
    Resource = $Resource
    StartDate = $StartDate
    EndDate = $EndDate
    AllRooms = $AllRooms
}

Add-Appointment @AppointmentSplat
```