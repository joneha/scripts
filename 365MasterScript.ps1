# ///////////////////////////////////////////////////////
# Main PS Script to combine a bunch of different standard tasks in my day-to-day
# -------------------------------------------------------
# Written by @joneha as a combination of scripts from several sources
# most of which I cannot find. If you feel like having credit in the header, please let me know.
# Known source(s):
# Aman Sharma - https://stackoverflow.com/a/36799057
# 
# For easy access to the script within a PowewrShell session, add this to you PowerShell_profile.ps1:
# function EXOScript {
#  . "[path-to-file]\Scripts\MasterScript.ps1"
# }
# 
# Last revised on 26/11/2020.
# //////////////////////////////////////////////////////

#region Function Definitions. These come first before calling them
function MailTraceAcct {
    <#
    # Mer generell uthenting av MailTrace
    #>

    $Kunde = Read-Host -Prompt 'Oppgi tenant-navn for den du vil koble deg til (uten .onmicrosoft)'
    IF (!$Kunde) {
        Write-Host 'Kunde kreves'
        Write-Host 'Scriptet avsluttes'
        Exit;
    }
    <#
    # Dersom du vil unng� � m�tte legge inn epostadresse, s� legger du inn informasjon under.
    # Sparer deg for et trinn i prosessen
    #>
    IF ($env:computername -eq "[your computername here]") {
        $Bruker = "[your email here]"
        Write-Host "Bruker: " $Bruker
    }
    Elseif ($env:computername -eq "[your computername here]") {
        $Bruker = "[your email here]"
        Write-Host "Bruker: " $Bruker
    }
    Else {
        $Bruker = Read-Host -Prompt 'Hva er din innloggingsadresse?'
        IF (!$Bruker) {
                Write-Host 'Brukernavn kreves'
                Write-Host 'Scriptet avsluttes'
                Exit;
        }
    }

    $Avsender = Read-Host -Prompt 'Hvilken avsender vil du sjekke eposter fra?'
    IF (!$Avsender) {
        Write-Host 'Avsender kreves'
        Write-Host 'Scriptet avsluttes'
        Exit;
    }
    [int]$PosDay = Read-Host 'Hvor mange dager bakover �nsker du � sjekke?'
    IF (!$PosDay) {
        $PosDay = 1
    }
    $NegDay = $PosDay * -1

    [int]$PageSize = 5000

    $page = Read-Host -Prompt '�nsket sidetall'
    IF (!$page) {
        $page = 1
    }

    Write-Host ""
    Write-Host ""
    Write-Host ""

    Write-Host "Dine valg er:"
    Write-Host "Din innlogging: " $Bruker
    Write-Host "Kundetenant: " $Kunde".onmicrosoft.com"
    Write-Host "Avsender � sjekke fra: " $Avsender
    Write-Host "Antall dager tilbake i tid: " $PosDay
    Write-Host "Side: " $page

    Write-Host ""
    Write-Host ""

    Write-Host "Velg et tall for � fortsette eller avslutte"
    Write-Host "1. Start script"
    Write-Host "0. Avslutt"

    Write-Host ""
    Write-Host ""

    # Vent p� input f�r tilkobling
    While ($true) {
        $pick = Read-Host "Velg et tall"

        If ($pick -eq "0") {
            "Avslutter"
            Break;
        }

        If (($pick -ne "1") -and ($pick -ne "0")) {
            Write-Host ""
            Write-Host -ForegroundColor Red "Tja. Dette valget er ikke gyldig, pr�v igjen..."
        }
    
        If ($pick -eq "1") {
            # Koble opp til kundes tenant
            Connect-ExchangeOnline -UserPrincipalName $Bruker -DelegatedOrganization $Kunde.onmicrosoft.com

            # Hente ut opptil 5000 elementer fra de siste $PosDay dager - For flere sider (dersom resultat er 5000) legg til -Page # etter -PageSize 5000
            Get-MessageTrace -StartDate (Get-Date).AddDays($NegDay) -EndDate (Get-Date) -SenderAddress $Avsender -PageSize $PageSize | Group-Object -Property SenderAddress | Select-Object Name, Count

            # Vent p� input f�r frakobling
            Read-Host -Prompt "Press Enter for � avslutte"

            # Koble fra for � ikke oppta session
            Disconnect-ExchangeOnline -Confirm:$false
            Break;
        }
    }

}
function FetchAccounts {    
    # Henter dagens dato for � navngi ut-filen: UserLicenseReport 31/08/20.csv for � lett kunne skille mellom filene
    $date = Get-Date -Format "dd/MM/yyyy"

    # Oppretter en PowerShell-session mot Office 365. Du vil f� en dialogboks som ber om innloggingsdetaljer
    # For Front Software er det krav om 2FA, s� forsikre deg om at du har tilgang til 2-faktor koden
    Connect-MsolService

    # Henter ut antall Tenants som er under din administrasjon
    $customers = Get-MsolPartnerContract -All
    Write-Host "Found $($customers.Count) customers for $((Get-MsolCompanyInformation).displayname)." -ForegroundColor DarkGreen

    # Setter ut-filen til � bli lagret i mappen PowerShell har �pen, med navnet i f�rste linje
    $CSVpath = "$(Get-Location)\UserLicenseReport $date.csv"

    # Henter ut alle lisensierte brukere fra 365
    # Printer alle til skjerm og skriver til CSV-fil. 
    foreach ($customer in $customers) {
        Write-Host "Retrieving license info for $($customer.name)" -ForegroundColor Green
        $licensedUsers = Get-MsolUser -TenantId $customer.TenantId -All | Where-Object { $_.islicensed }
  
        foreach ($user in $licensedUsers) {
            Write-Host "$($user.displayname)" -ForegroundColor Yellow  
            $licenses = $user.Licenses
            $licenseArray = $licenses | foreach-Object { $_.AccountSkuId }
            # $licenseString = $licenseArray -join ", " # Tatt bort for � lage ny linje for hver lisens
            foreach ($licenseString in $licenseArray) {
            Write-Host "$($user.displayname) har $licenseString" -ForegroundColor Blue
            $licensedSharedMailboxProperties = [pscustomobject][ordered]@{
                CustomerName      = $customer.Name
                DisplayName       = $user.DisplayName
                License           = $licenseString
                TenantId          = $customer.TenantId
                UserPrincipalName = $user.UserPrincipalName
                AccountCreated    = ($user.WhenCreated).ToShortDateString()
            }
            $licensedSharedMailboxProperties | Export-CSV -Path $CSVpath -Append -NoTypeInformation
        }
        }
    }
    # CleanCsv; # Ikke testet - I teorien rensker den da opp filen selv
}
function CleanCsv {    
    # Denne filen sjekker dagens dato, leter etter en eksportfil fra 365 (opprettet i Fetch-O365Users)
    # Den kj�rer da en enkel find-and-replace for � gj�re sluttresultatet enklere � lese for fakturering

    $date = Get-Date -Format "dd/MM/yyyy"
    $orig_file = "UserLicenseReport $date.csv"
    $dest_file = "Kundebrukere $date.csv"

    if (Test-Path $orig_file) {
        (Get-Content $orig_file) | ForEach-Object {
            $_ -replace 'reseller-account:', '' `
                -replace 'EXCHANGESTANDARD', 'Exchange Online (Plan 1)' `
                -replace 'EXCHANGEDESKLESS', 'Exchange Online Kiosk' `
                -replace 'DESKLESSPACK', 'Office 365 F3' `
                -replace 'TEAMS_EXPLORATORY', 'Microsoft Teams Exploratory' `
                -replace 'TEAMS EXPLORATORY', 'Microsoft Teams Exploratory' `
                -replace 'O365 BUSINESS PREMIUM', 'Microsoft 365 Business Standard' `
                -replace 'O365_BUSINESS_PREMIUM', 'Microsoft 365 Business Standard' `
                -replace 'POWER BI PRO', 'Power BI Pro' `
                -replace 'POWER_BI_PRO', 'Power BI Pro' `
                -replace 'POWER BI STANDARD', 'Power BI (free)' `
                -replace 'POWER_BI_STANDARD', 'Power BI (free)' `
                -replace 'FLOW FREE', '' `
                -replace 'FLOW_FREE', '' `
                -replace 'TEAMS COMMERCIAL TRIAL', '' `
                -replace 'TEAMS_COMMERCIAL_TRIAL', '' `
                -replace 'O365 BUSINESS ESSENTIALS', 'Microsoft 365 Business Basic' `
            -replace 'O365_BUSINESS_ESSENTIALS', 'Microsoft 365 Business Basic'
        } | Set-Content -Path $($dest_file)
        Write-Host "`nFilen $orig_file har blitt redigert. Ny og mer lettleselig versjon er lagret som: $(Get-Location)\$dest_file `n" -ForegroundColor Green
    }
    Else {
        Write-Host "`n$(Get-Location)\$orig_file ble ikke funnet. Kan ikke erstatte tekst. `n" -ForegroundColor Red
    }

}
function ConnectClient {
    
    $Kunde = Read-Host -Prompt 'Oppgi tenant-navn for den du vil koble deg til (uten .onmicrosoft)'
    
    IF (!$Kunde) {
        Write-Host 'Kunde kreves'
        Write-Host 'Scriptet avsluttes'
        Exit;
    }
    <#
    # Dersom du vil unng� � m�tte legge inn epostadresse, s� legger du inn informasjon under.
    # Sparer deg for et trinn i prosessen
    #>
    IF ($env:computername -eq "[your computername here]") {
        $Bruker = "[your email here]"
        Write-Host "Bruker: " $Bruker
    }
    Elseif ($env:computername -eq "[your computername here]") {
        $Bruker = "[your email here]"
        Write-Host "Bruker: " $Bruker
    }
    Else {
        $Bruker = Read-Host -Prompt 'Hva er din innloggingsadresse?'
        IF (!$Bruker) {
                Write-Host 'Brukernavn kreves'
                Write-Host 'Scriptet avsluttes'
                Exit;
        }
    }
    
    IF (!$Bruker) {
        Write-Host 'Brukernavn kreves'
        Write-Host 'Scriptet avsluttes'
        Exit;
    }

    Write-Host ""
    Write-Host "Connect-ExchangeOnline -UserPrincipalName" $Bruker "-DelegatedOrganization" $Kunde".onmicrosoft.com"
    Write-Host ""
    Write-Host "Husk � koble fra f�r du avslutter PowerShell med kommando:"
    Write-Host ""
    Write-Host "Disconnect-ExchangeOnline"
    Write-Host ""
    Exit;
}
function MSOLConnected {
    Get-MsolDomain -ErrorAction SilentlyContinue | out-null
    $result = $?
    return $result
}
#endregion

#region Showing Options
Write-Host -ForegroundColor Yellow "Velkommen til masterscriptet for Office 365"
""
Start-Sleep -Seconds 1
# Write-Host -ForegroundColor Yellow "Velg en oppgave:"

# Write-Host -ForegroundColor Yellow "1. MailTrace mot spesifikk avsender til kundetenant"
# Write-Host -ForegroundColor Yellow "2. Hente ut alle lisensierte kundebrukere" 
# Write-Host -ForegroundColor Yellow "3. Renske CSV med kundebrukere"
# Write-Host -ForegroundColor Yellow "4. Print n�dvendige kommandoer for � koble p� kundetenant"
# Write-Host -ForegroundColor Yellow "0. Exit"

#endregion

#region Getting input
While ($true) {
    Write-Host -ForegroundColor Yellow "Alternativer:"
    ""
    Write-Host -ForegroundColor Yellow "1. MailTrace mot spesifikk avsender til kundetenant"
    Write-Host -ForegroundColor Yellow "2. Hente ut alle lisensierte kundebrukere" 
    Write-Host -ForegroundColor Yellow "3. Renske CSV med kundebrukere"
    Write-Host -ForegroundColor Yellow "4. Print n�dvendige kommandoer for � koble p� kundetenant"
    Write-Host -ForegroundColor Yellow "0. Exit"

    $Valg = Read-Host "Velg en oppgave fra listen over"

    If ($Valg -eq "0") {
        "Takk for meg"; 
        # if (-not (MSOLConnected)) {
            Break;
        # }
        # else {
            # Disconnect-ExchangeOnline -Confirm:$false
            # Break; 
        # }
    }

    If (($Valg -ne "1") -and ($Valg -ne "2") -and ($Valg -ne "3") -and ($Valg -ne "4") -and ($Valg -ne "0")) {
        ""
        Write-Host -ForegroundColor Red "Tja. Dette valget er ikke gyldig, pr�v igjen..."
    }

    #region Main Code

    switch ($Valg) { 
        1 {
            MailTraceAcct;
        } 
        2 {
            FetchAccounts;
        } 
        3 {
            CleanCsv;
        }
        4 {
            ConnectClient;
        }
        default { "Vennligst velg en gyldig oppgave" }
    }
    #endregion
}
#endregion