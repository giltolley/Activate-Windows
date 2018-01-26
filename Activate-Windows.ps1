#requires -version 4
<#
.SYNOPSIS
  Short script to automate the activation of Windows using the OA3xOriginalProductKey WMI entry.
.DESCRIPTION
  Short and sweet, I just needed something to take the place of manually typing out the commands,
  or copying and pasting them everytime.

  I might expand it in the future to allow for remote calls, but I just want to remove a few steps
  for right now.
.INPUTS
  None
.OUTPUTS
  None
.NOTES
  Version:        1.0
  Author:         Gil Tolley
  Creation Date:  9/28/2017
  Purpose/Change: Initial script development
.EXAMPLE
  Basic command
  
  .\Activate-Windows.ps1
#>

#---------------------------------------------------------[Script Parameters]------------------------------------------------------

#Param (
  #Script parameters go here
#)

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

#$ErrorActionPreference = 'SilentlyContinue'

#Import Modules & Snap-ins

#----------------------------------------------------------[Declarations]----------------------------------------------------------

$LicensingService =  Get-WmiObject -query 'select * from SoftwareLicensingService'
#$WindowsKey = Get-WmiObject -Class SoftwareLicensingService -ComputerName $env:COMPUTERNAME -Property OA3xOriginalProductKey

#-----------------------------------------------------------[Functions]------------------------------------------------------------

<#
Function <FunctionName> {
  Param ()
  Begin {
    Write-Host '<description of what is going on>...'
  }
  Process {
    Try {
      <code goes here>
    }
    Catch {
      Write-Host -BackgroundColor Red "Error: $($_.Exception)"
      Break
    }
  }
  End {
    If ($?) {
      Write-Host 'Completed Successfully.'
      Write-Host ' '
    }
  }
}
#>

#-----------------------------------------------------------[Execution]------------------------------------------------------------
$Key = $LicensingService.OA3xOriginalProductKey

$LicensingService.InstallProductKey($Key)
$LicensingService.RefreshLicenseStatus()