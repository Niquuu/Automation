$etunimi = Read-Host "Anna käyttäjän etunimi" `n""
$sukunimi = Read-Host "Anna käyttäjän sukunimi" `n""

# Tämä muutetaan asiakkaan mukaan!!!
$vakioAD = "OU=, DC=, DC=" 

$vakio = $etunimi + $sukunimi
# Asettaa käyttäjä nimen oletusarvoksi Etu- ja sukunimi, ellei sitä muuta.
if (!($kayttajaNimi = Read-Host "Paina enter hyväksyäksesi $vakio käyttäjänimeksi tai anna käyttäjänimi" `n"")) {$kayttajaNimi = $vakio} 
$kokoNimiDomain = $etunimi + "." + $sukunimi + "@."

$salasana = Read-Host "Anna salasana"  -AsSecureString  `n""
$puhelinNumero = Read-Host "Anna käyttäjän puhelinnumero" `n""
$kaupunki = "" # Voidaan säätää vakioksi asiakas kohtaisesti
$osoite = "" # Voidaan säätää vakioksi asiakas kohtaisesti
$postinumero = "" # Voidaan säätää vakioksi asiakas kohtaisesti
$yritys = "" # Voidaan säätää vakioksi asiakas kohtaisesti
$yksikko = Read-Host "Anna käyttäjän yksikkö" `n""
$titteli = Read-Host "Anna käyttäjän titteli" `n""


$sahkoposti = $etunimi + "." + $sukunimi + "@.fi"
$smtpAD1 = $etunimi + "." + $sukunimi + "@.fi"
$smtpAD2 = $etunimi + "." + $sukunimi + "@.onmicrosoft.com"

Clear-Host; 

if (Get-ADUser -F { SamAccount -eq $kayttajaNimi }) {

    Write-Host -ForegroundColor Red "Käyttäjä $kayttajaNimi on jo olemassa!"

} else {

    New-ADUser `
    -SamAccountName $kayttajaNimi `
    -UserprincipalName "$kokoNimiDomain" `
    -Name  "$etunimi $sukunimi" `
    -GivenName $etunimi `
    -Surname $sukunimi `
    -Enable $True `
    -ChangePasswordAtLogon $True `
    -DisplayName "$etunimi $sukunimi" `
    -Department $yksikko `
    -EmailAddress $sahkoposti `
    -MobilePhone $puhelinNumero `
    -Company $yritys `
    -City $kaupunki `
    -StreetAddress $osoite `
    -PostalCode $postinumero `
    -Title $titteli `
    -AccountPassword (convertto-securestring $salasana -AsPlainText -Force) `
    -Path $vakioAD

    Set-AdUser $kayttajaNimi -Add @{ProxyAddresses="$smtpAD1"}
    Set-AdUser $kayttajaNimi -Add @{ProxyAddresses="$smtpAD2"}

    # Vaihda asiakas kohtaisesti mikäli tarvetta.
    $HomeFolder = '' + $kayttajanimi
    Set-ADUser -Identity $kayttajaNimi -HomeDirectory $HomeFolder -HomeDrive "";
    
}

$ehto = $true;
    while($ehto) {
        Write-Host "Käyttäjä $kayttajaNimi ei kuulu vielä mihinkään ryhmään."
        (Get-ADUser $kayttajaNimi –Properties MemberOf | Select-Object MemberOf).MemberOf
        $valinta = Read-Host "Haluatko lisätä $kayttajaNimi johonkin ryhmään?" `
        `n"1: Kyllä" `n"2: En" `n""
    switch($valinta) {

        1 {
            $ouUnit = Get-ADGroup -filter * | select-Object Name
            $lista = @{}

            for ($i = 1; $i -le $ouUnit.Count; $i++) {
            write-host "$i. $($ouUnit[$i - 1].Name)"
            $lista.Add($i, ($ouUnit[$i - 1 ].Name))
            } do {
            [int]$ans = Read-Host "Mihin ryhmään haluat lisätä käyttäjän $kayttajaNimi" `n""
            }    
            until ([int]$ans -ige 1 -and [int]$ans -ile $ouUnit.length)
            $selection = $lista.Item($ans) 
            Add-ADGroupMember -Identity $selection -Members $kayttajaNimi          
            }
        2 { $ehto = $false; break;}
    }
    
    }

 Start-AdSyncSyncCycle -PolicyType Initial



