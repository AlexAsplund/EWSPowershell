@{
    # If authoring a script module, the RootModule is the name of your .psm1 file
    RootModule = 'EWS.psm1'

    Author = 'Alex Asplund <https://github.com/AlexAsplund>'

    CompanyName = 'https://github.com/AlexAsplund'

    ModuleVersion = '0.1'

    # Use the New-Guid command to generate a GUID, and copy/paste into the next line
    GUID = '1c2dbc52-104f-476f-a244-14eecb30363a'

    Copyright = '2017 Alex Asplund'

    Description = 'Connects to Exchange via the EWS API and can create appointments, impersonation etc'

    # Minimum PowerShell version supported by this module (optional, recommended)
    # PowerShellVersion = '4.0'

    # Which PowerShell Editions does this module work with? (Core, Desktop)
    CompatiblePSEditions = @('Desktop', 'Core')

    # Which PowerShell functions are exported from your module? (eg. Get-CoolObject)
    FunctionsToExport = @(  'Add-EWSAppointment',
                            'Get-EWSAllRooms',
                            'Get-EWSAppointment',
                            'Invoke-EWSExchangeService',
                            'Set-EWSUserImpersonation',
                            'Search-EWSAppointment',
                            'Remove-EWSAppointment',
                            'Remove-EWSAppointment'
                            )

    # PowerShell Gallery: Define your module's metadata
    PrivateData = @{
        PSData = @{
            # What keywords represent your PowerShell module? (eg. cloud, tools, framework, vendor)
            Tags = @('EWS', 'Exchange','Impersonation','Office365')

            # What software license is your code being released under? (see https://opensource.org/licenses)
            LicenseUri = 'https://opensource.org/licenses/MIT'

            # What is the URL to your project's website?
            ProjectUri = 'https://github.com/AlexAsplund'

            # What new features, bug fixes, or deprecated features, are part of this release?
            ReleaseNotes = @'
            Created the module
'@
        }
    }
}