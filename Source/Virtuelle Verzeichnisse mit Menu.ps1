Add-PSSnapin *exchange* -ErrorAction Ignore

function VirtuelleVerzeichnisseAnzeigen() {
    Clear
        
    Write-Host "Anzeigen`n"

    Get-OwaVirtualDirectory | Format-List Name,InternalURL,ExternalURL
    Read-Host "Enter für weiter" | Out-Null

    Get-ECPVirtualDirectory | Format-List Name,InternalURL,ExternalURL
    Read-Host "Enter für weiter" | Out-Null

    Get-WebServicesVirtualDirectory | Format-List Name,InternalURL,ExternalURL
    Read-Host "Enter für weiter" | Out-Null

     Get-ActiveSyncVirtualDirectory | Format-List Name,InternalURL,ExternalURL
    Read-Host "Enter für weiter" | Out-Null

     Get-OabVirtualDirectory | Format-List Name,InternalURL,ExternalURL
    Read-Host "Enter für weiter" | Out-Null


    Get-MapiVirtualDirectory | Format-List Name,InternalURL,ExternalURL
    Read-Host "Enter für weiter" | Out-Null   

   
    Get-PowerShellVirtualDirectory | Format-List Name,InternalURL,ExternalURL
    Read-Host "Enter für weiter" | Out-Null
    

    Get-ClientAccessService | Format-List Name,InternalURL,ExternalURL
    Read-Host "Enter für weiter" | Out-Null 

    Get-OutlookAnywhere | Format-List Name,InternalURL,ExternalURL
    Read-Host "Enter für weiter" | Out-Null
    
}

function VirtuelleVerzeichnisseBearbeiten() {
    Clear
        
    Write-Host "Bearbeiten`n"

    $servername = $env:computername
    $internalhostname = Read-Host -Prompt "Internal-URL"
    $externalhostname = Read-Host -Prompt "External-URL"
    $autodiscoverhostname = Read-Host -Prompt "Autodiscover-URL"

    $owainturl = "https://" + "$internalhostname" + "/owa"
    $owaexturl = "https://" + "$externalhostname" + "/owa"
    $ecpinturl = "https://" + "$internalhostname" + "/ecp"
    $ecpexturl = "https://" + "$externalhostname" + "/ecp"
    $ewsinturl = "https://" + "$internalhostname" + "/EWS/Exchange.asmx"
    $ewsexturl = "https://" + "$externalhostname" + "/EWS/Exchange.asmx"
    $easinturl = "https://" + "$internalhostname" + "/Microsoft-Server-ActiveSync"
    $easexturl = "https://" + "$externalhostname" + "/Microsoft-Server-ActiveSync"
    $oabinturl = "https://" + "$internalhostname" + "/OAB"
    $oabexturl = "https://" + "$externalhostname" + "/OAB"
    $mapiinturl = "https://" + "$internalhostname" + "/mapi"
    $mapiexturl = "https://" + "$externalhostname" + "/mapi"
    $aduri = "https://" + "$autodiscoverhostname" + "/Autodiscover/Autodiscover.xml"
  
    

    Get-OwaVirtualDirectory -Server $servername | Set-OwaVirtualDirectory -internalurl $owainturl -externalurl $owaexturl
    Get-EcpVirtualDirectory -server $servername | Set-EcpVirtualDirectory -internalurl $ecpinturl -externalurl $ecpexturl
    Get-WebServicesVirtualDirectory -server $servername | Set-WebServicesVirtualDirectory -internalurl $ewsinturl -externalurl $ewsexturl
    Get-ActiveSyncVirtualDirectory -Server $servername  | Set-ActiveSyncVirtualDirectory -internalurl $easinturl -externalurl $easexturl
    Get-OabVirtualDirectory -Server $servername | Set-OabVirtualDirectory -internalurl $oabinturl -externalurl $oabexturl
    Get-MapiVirtualDirectory -Server $servername | Set-MapiVirtualDirectory -externalurl $mapiexturl -internalurl $mapiinturl
    Get-OutlookAnywhere -Server $servername | Set-OutlookAnywhere -externalhostname $externalhostname -internalhostname $internalhostname -ExternalClientsRequireSsl:$true -InternalClientsRequireSsl:$true -ExternalClientAuthenticationMethod 'Negotiate'
    Get-ClientAccessService $servername | Set-ClientAccessService -AutoDiscoverServiceInternalUri $aduri
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