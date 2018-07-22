Add-PSSnapin *exchange* -ErrorAction Ignore

function VirtuelleVerzeichnisseAnzeigen() {
    Clear
        
    Write-Host "Anzeigen`n"

    Get-ECPVirtualDirectory | Format-List Name,InternalURL,ExternalURL

    Read-Host "Enter für weiter" | Out-Null
}

function VirtuelleVerzeichnisseBearbeiten() {
    Clear
        
    Write-Host "Anzeigen`n"

    Get-ECPVirtualDirectory | Format-List Name,InternalURL,ExternalURL

    Read-Host "Enter für weiter" | Out-Null
}


function Menu
{
    do
    {
        Clear

        Write-Host "Virtuelle Verzeichnisse"
        Write-Host " "
        Write-Host "1: Anzeigen."
        Write-Host "2: Bearbeiten."
        Write-Host "3: Skript beenden."
        Write-Host ""
    
        $input = Read-Host -Prompt "Auswahl"

        switch ($input)
        {
            '1' { VirtuelleVerzeichnisseAnzeigen } 
            
            '2' { VirtuelleVerzeichnisseBearbeiten }  
        }
    } until ($input -eq '3')
}

Menu