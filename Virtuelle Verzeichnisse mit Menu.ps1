function Menu
{
     param (
           [string]$Title = 'Virtuelle Verzeichnise'
     )
     cls
     Write-Host "$Title"
     Write-Host " "
     Write-Host "1: Anzeigen."
     Write-Host "2: Bearbeiten."
     Write-Host "3: Script beenden."
     Write-Host " "
}
do
{
     Menu
     $input = Read-Host "Auswahl"
     switch ($input)
     {
           '1' {
                cls
                'Anzeigen'
                Get-ECPVirtualDirectory | Format-List Name,InternalURL,ExternalURL

           } '2' {
                cls
                'Bearbeiten'
           }  '3' {
                return
           }
     }
     pause
}
until ($input -eq '3')
