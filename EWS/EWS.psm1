Import-Module -Name "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"

<#
.SYNOPSIS
    Creates a new appointment
.DESCRIPTION
    Creates a new appointment with the parameters specified. Can use Impersonation if specified beforehand with Set-UserImpersonation
.EXAMPLE
    $AppointmentSplat = @{
    Subject = "Important meeting"
    Body = "Content for meeting"
    Resource = "Room to use or resource"
    StartDate = [DateTime]$StartDate
    EndDate = [DateTime]$EndDate
    AllRooms = $AllRooms
}

$Appointment = Add-Appointment @AppointmentSplat
.NOTES
    $ExchangeService
#>
function Add-EWSAppointment {
    [CmdletBinding()]
    param (
        # Subject
        [Parameter(Mandatory = $true)]
        [string]
        $Subject,

        # Body
        [Parameter(Mandatory = $true)]
        [string]
        $Body,

        # Resource
        [Parameter(Mandatory = $false)]
        [ValidatePattern(".*\@.*")]
        [string]
        $Resource,

        # Start
        [Parameter(Mandatory = $true)]
        [datetime]
        $StartDate,

        # End
        [Parameter(Mandatory = $true)]
        [DateTime]
        $EndDate,

        # All rooms from $ExchangeService.GetRoomLists() | foreach {$ExchangeService.GetRooms($_)} | 
        [parameter(Mandatory = $true)]
        [Array]
        $AllRooms

    )
    
    begin {
        if ($ExchangeService -eq $null) {
            Write-Error "Exchange session not invoked. Create an exchange-session with Invoke-ExchangeSession first!"
            Break
        }
    }
    
    process {
        
        try {
            $Appointment = [Microsoft.Exchange.WebServices.Data.Appointment]::new($ExchangeService)
            $Appointment.Subject = $Subject
            $Appointment.Body = $Body
            $Appointment.Start = $StartDate
            $Appointment.End = $EndDate
            $Appointment.Location = ($AllRooms | Where-Object {$_.Mail -eq $Resource}).RoomString
            
            # If resource parameter is set: Add the resource
            if (![string]::IsNullOrEmpty($Resource)) {

                $Appointment.Resources.Add($Resource)

            }

            $Appointment.Save([Microsoft.Exchange.WebServices.Data.SendInvitationsMode]::SendToAllAndSaveCopy)
            $Success = $true
            $Exception = ""
        }
        catch {
            $Success = $false
            $Exception = $_.exception
            Write-Error $_.exception
        }
    }
    
    end {
        $AppointmentData = @{
            Subject   = $Subject
            Body      = $Body
            StartDate = (Get-Date $StartDate -Format "yyyy-MM-dd HH:mm:ss")
            EndDate   = (Get-Date $EndDate -Format "yyyy-MM-dd HH:mm:ss")
            Location  = $Resource
            Success   = $Success
            Exception = $Exception
        }
        Return $AppointmentData
    }
}

function Set-EWSUserImpersonation ($UserID) {
   
    $ImpersonateID = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $UserID)
    $Global:ExchangeService.ImpersonatedUserId = $ImpersonateID
    
}

function Invoke-EWSExchangeService {
    param(
        [parameter(Mandatory = $True)]
        [PSCredential]
        $Credential,
        [parameter(Mandatory = $True)]
        [string]
        $Domain,
        [parameter(Mandatory = $True)]
        [string]
        $AutoDiscoverUrl
        
    )

    # Set the domain in networkcredential - don't know if this is the way to do it but it works.
    $Credential.GetNetworkCredential().Domain = $Domain

    # Create ExchangeCredential from regular [PSCredential]
    $ExchangeCredential = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials($Credential.Username, $Credential.GetNetworkCredential().Password, $Credential.GetNetworkCredential().Domain)
    
    # Create the exchange service object
    $Global:ExchangeService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService 
    
    # Not sure if this is needed
    $Global:ExchangeService.UseDefaultCredentials = $true 
    
    # Add the ExchangeCredential to the ExchangeService
    $Global:ExchangeService.Credentials = $ExchangeCredential
    
    # {$true} is needed if it's office 365 with a redirect on the autodiscovery. I think there's a lot more secure ways to check if
    # a redirection is correct.
    $Global:ExchangeService.AutodiscoverUrl($AutoDiscoverUrl, {$true})

}

# Simple function to get the room data that i want to create appointments.
Function Get-EWSAllRooms {
    $AllRooms = $ExchangeService.GetRoomLists() | ForEach-Object {$ExchangeService.GetRooms($_)} | Select-Object @{Name = "Mail"; Expression = {$_.Address}}, @{Name = "RoomString"; Expression = {$_.Name}}
    Return $AllRooms
}

<#
.SYNOPSIS
    Gets all appointments within $StartDate and $EndDate
.DESCRIPTION
    Gets all appointments within $StartDate and $EndDate
.EXAMPLE
    PS C:\> Get-EWSAppointment -StartDate (Get-Date) -EndDate (Get-Date).AddDays(14)
#>
function Get-EWSAppointment {
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $True, Position = 0)]
        [DateTime]
        $StartDate,
        [parameter(Mandatory = $True, Position = 1)]
        [DateTime]
        $EndDate
    )
    
    begin {
        $CalendarFolder = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar
        $CalendarBind = [Microsoft.Exchange.WebServices.Data.CalendarFolder]::Bind($ExchangeService, $CalendarFolder, [Microsoft.Exchange.WebServices.Data.PropertySet]::new())
        $CalendarView = [Microsoft.Exchange.WebServices.Data.CalendarView]::new($StartDate, $EndDate)
        $CalendarView.PropertySet = [Microsoft.Exchange.WebServices.Data.PropertySet]::new([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Subject, [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Start, [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::End, [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Duration)

    }
    
    process {
        $SearchResult = $CalendarBind.FindAppointments($CalendarView)
    }
    
    end {
        Return $SearchResult
    }
}

function Remove-EWSAppointment {
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [parameter(Mandatory = $True)]
        [System.object]
        $AppointmentItem,
        
        [ValidateSet('CancelMeeting', 'Decline', 'Delete')]
        [parameter(Mandatory = $True)]
        [string]
        $Action,
        
        [string]
        $CancellationMessageText = ''
    )
    
    begin {
        
    }
    
    process {
        Foreach ($Appointment in $AppointmentItem) {
            If ($pscmdlet.ShouldProcess($Appointment.Subject, $Action)) {
                
                
                if ($Action -Eq 'CancelMeeting') {
                    $Appointment.CancelMeeting($CancellationMessageText)
                }

                if ($Action -Eq 'Decline') {
                    $Appointment.Decline($True)
                }

                if ($Action -Eq 'Delete') {
                    $Appointment.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::MoveToDeletedItems)
                }
                
            }
        
        }
    }
    end {
    }
}

function Send-EWSMailMessage {
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [string]$Subject,
        [string]$Body,
        [Array]$To,
        [Array]$CC,
        [Array]$BCC
    )
    
    begin {
        $Message = [Microsoft.Exchange.WebServices.Data.EmailMessage]::new($ExchangeService)
        $Message.Subject = $Subject
        $Message.Body = $Body
        
        $To | ForEach-Object {
            $Message.ToRecipients.Add($_) | Out-Null
        }
        if ($CC -ne $null) {
            $CC | ForEach-Object {
                $Message.ToRecipients.Add($_) | Out-Null
            }
        }
        if ($BCC -ne $null) {
            $BCC | ForEach-Object {
                $Message.ToRecipients.Add($_) | Out-Null
            }
        }
        
    }
    
    process {
        If ($pscmdlet.ShouldProcess($Subject, "Sending mailmessage")) {
            $Message.SendAndSaveCopy()
        }
    }
    
    end {
        Return $true
    }
}

# Todo
# Send-EWSMailMessage
# Get-EWSMailMessage
# Set-EWSOutOfOffice
# 