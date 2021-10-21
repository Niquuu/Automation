$ehtoKoko = $true
# Disconnect-ExchangeOnline -Confirm:$false
# Connect-MsolService 
while ($ehtoKoko) {

  # Tyhjentää kaikki aikasemmat ExchangeOnline sessiot. Niin ei aiheuta erroria mikäli avaa liian monta kertaa.
  

  Clear-Host;  
  Write-Host -ForegroundColor Red "
##############################################
#                                            #
#               O365 Pack                    #
#             Work In Progress               #
#                                            #
##############################################
"


  # Jokaisen tulee kerran vähintään asentaa nämä moduulit! Tarvitaanko lisää moduuleita?

  # Install-module MSOnline
  # Import-Module ExchangeOnlineManagement

  # Tarvitaanko muita connect MsolService ja ExchangeOnline

  # Connect-ExchangeOnline -UserPrincipalName $cred -DelegatedOrganization $selection

  # $cred = Get-Credential


  # Kommentoi ulos kun valmis!
  # Connect-MsolService 



  $valintaTenant = Read-Host "Valitse seuraavista vaihtoehdoista" `n"1: Single tenant" `n"2: Multi-tenant" `n"3: Lopetus" `n""
  switch ($valintaTenant) {
    1 {
      Clear-Host;
      write-host -ForegroundColor Red "        
      #########################
      #                       #
      # Single Tenant Scripts #
      #                       #
      #########################
      ";
      
      $array = (Get-MsolPartnerContract | Select-Object defaultdomainname)
      $menu_single = @{}
            
      for ($i = 1; $i -le $array.Count; $i++) {
        Write-host "$i. $($array[$i-1].defaultdomainname)"
        $menu_single.Add($i, ($array[$i - 1].defaultdomainname))
      } do {
        [int]$ans = Read-Host "Valitse listasta asiakas (Numero)" `n""
      } until ([int]$ans -ige 1 -and [int]$ans -ile $array.length)
      $selection = $menu_single.Item($ans)

      # Yhdistää ExchangeOnlineen tenattiin, kun se on valittu. Kommentoi ulos kun valmis!
      # Connect-ExchangeOnline -DelegatedOrganization $selection

      Clear-Host;

      
      Clear-Host;
      $tenant = Get-MsolPartnerContract -DomainName $selection | Select-Object -ExpandProperty tenantid
          
          
      $ehtoLista1 = $true;
      while ($ehtoLista1) {
        $lista = get-msoluser -All -TenantId $tenant | Select-Object UserPrincipalName
        write-host ($lista | Select-Object UserPrincipalName | Out-String)
            
        $userAsked = Read-Host "Anna käyttäjän sähköpostiosoite mihin annetaan lisää oikeuksia (Etunimi.Sukunimi@yritys.com)" `n""
        $userCheck = get-msoluser -all -TenantId $tenant | Where-Object { ($_.UserPrincipalName -eq "$userAsked") } | Select-Object -ExpandProperty UserPrincipalName
            
        if ($userAsked -eq $userCheck) {
          $ehtoLista1 = $false; break;
        }
        else {            
          Write-Host -foregroundcolor Red "Käyttäjää $userAsked ei löytynyt." `n
        }
      }
      Clear-Host;

      $ehtoLista2 = $true;
      while ($ehtoLista2) {
        $lista = get-msoluser -All -TenantId $tenant | Select-Object UserPrincipalName
        write-host ($lista | Select-Object UserPrincipalName | Out-String)

        $Userasked1 = Read-Host "Anna sähköpostiosoite mihin lisäät oikeuksia käyttäjälle $userAsked (Etunimi.sukunimi@yritys.com)" `n""
        $userCheck = Get-MsolUser -All -TenantId $tenant | Where-Object { ($_.userprincipalname -eq $Userasked1) } | Select-Object -ExpandProperty userprincipalname

        if ($userAsked1 -eq $userCheck) {
          Write-Host "Käyttäjä löytyi." $ehtoLista2 = $false; break; 
        }
        else {            
          Write-Host -foregroundcolor Red "Käyttäjää $userAsked1 ei löytynyt." `n
        }
      }
          
      $valintaKayttaja = $true; 
      while ($valintaKayttaja) {
        Clear-Host;
        Write-Host -ForegroundColor Red "
        ##############################################
        #                                            #
        #                                            #
        #             Work In Progress               #
        #                                            #
        ##############################################
        "; 

        $valintaToiminto = Read-Host "Valitse mitä haluat tehdä käyttäjälle: $userAsked" `n"1: Kalenterioikeudet" `n"2: Postilaatikon oikeudet" `n"3: Lähetysoikeudet" `n"4: Poissaoloviesti" `n"5: Jakeluryhmät" `n"6: Tarkistukset" `n"7: Palaa takaisin" `n""
        switch ($valintaToiminto) {
          1 {
            Clear-Host;
            write-Host -ForegroundColor Red "                                        
            #####################
            #                   #
            # Kalenterioikeudet #
            #                   #
            #####################
            ";
            Write-Host -ForegroundColor Yellow "
Mitä oikeudet tekevät:
- [Contributor] - the person can put appointments on your calendar but cannot see details of existing appointments
- [Reviewer] - the person can read everything related to an appointment (except a private one) and see folders, but not subfolders
- [Author] - the person can see appointment details, create appointments, edit appointments they created, and delete appointments they created
- [Editor] - the person can create items, edit all appointments, delete any appointment, and see the full details of all appointments
- [Owner] - the person will have the same permissions to your calendar that you have
- [Oikeuksien poisto] - Jos haluat poistaa kaikki oikeudet henkilöltä kalenteriin"; 
            
            $ehtoKalenteri = $true;
            while ($ehtoKalenteri) {
              $calendar = Get-MailboxFolderStatistics -Identity $userAsked -FolderScope Calendar | Where-Object { $_.FolderType -eq 'Calendar' } | Select-Object Name, FolderId 
              $permCalendar = Get-MailboxFolderPermission -Identity $userAsked":\$($calendar.Name)"

              Write-Host -ForegroundColor Yellow "Olet muuttamassa käyttäjän $userAsked oikeuksia kalenteriin $userAsked1" `n 
              Write-Host -ForegroundColor Yellow "$userasked on seuraaviin kalentereihin oikeudet:" ($permCalendar | Format-table -Force | Out-String) 
              $kalenteri = Read-Host "Mitä oikeuksia haluat muuttaa" `n"1: Contributor" `n"2: Reviewer" `n"3: Author" `n"4: Editor" `
                `n"5: Owner" `n"6: Oikeuksien poisto" `n"7: Palaa takaisin" `n 


              # Hakee oikeudet, mutta ei tulosta ulos niitä oikein!
              Write-Host ($permCalendar | Format-Table | Out-String )

              switch ($kalenteri) {
                1 { Add-MailboxFolderPermission -Identity $userAsked":\Kalenteri" -User $userAsked1 -AccessRights Contributor -SharingPermissionFlags Delegate }
                2 { Add-MailboxFolderPermission -Identity $userAsked":\Kalenteri" -User $userAsked1 -AccessRights Reviewer -SharingPermissionFlags Delegate }
                3 { Add-MailboxFolderPermission -Identity $userAsked":\Kalenteri" -User $userAsked1 -AccessRights Author -SharingPermissionFlags Delegate }
                4 { Add-MailboxFolderPermission -Identity $userAsked":\Kalenteri" -User $userAsked1 -AccessRights Editor -SharingPermissionFlags Delegate }
                5 { Add-MailboxFolderPermission -Identity $userAsked":\Kalenteri" -User $userAsked1 -AccessRights Owner -SharingPermissionFlags Delegate }
                6 { Remove-MailBoxFolderPermission -Identity $userAsked":\Kalenteri" -User $userAsked1 }
                7 { $ehtoKalenteri = $false; break; }               
              }             
            }
          }

          2 {
            Clear-Host;
            Write-host -ForegroundColor Red "         
            ##########################
            #                        #
            # Postilaatikon oikeudet #
            #                        #
            ##########################                
            ";
            $ehtoPosti = $true
            while ($ehtoPosti) {


              $postiOikeudet = Read-Host "Haluatko lisätä käyttäjälle $userAsked FullAccess oikeudet käyttäjän $userAsked1 postilaakkoon" `n"1: Kyllä" `n"2: En" `n""
              Switch ($postiOikeudet) {
                1 { Add-MailboxPermission -Identity $userAsked -User $userAsked1 -AccessRights FullAccess -InheritanceType All }

                # Lisää kaikkiin organisaation postilaakoihin admin oikeudet
                # 2{ Get-Mailbox -ResultSize unlimited -Filter "(RecipientTypeDetails -eq 'UserMailbox') -and (Alias -ne 'Admin')" | Add-MailboxPermission -User $userAsked -AccessRights FullAccess -InheritanceType All }

                2 { $ehtoPosti = $False; break; }
              }
            }
          }

          3 {
            Clear-Host;
            Write-Host -ForegroundColor Red "
            ###################
            #                 #
            # Lähetysoikeudet #
            #                 #
            ###################
            ";
            $ehtoLahetys = $true;
            while ($ehtoLahetys) {

              $laheytusOikeudet = Read-Host "Mitkä oikeudeut annetaan" `n"1: Send On Behalf" `n"2: Send As" `n"3: Palaa takaisin" `n""
              switch ($laheytusOikeudet) {
                1 { Set-Mailbox $userAsked - GrantSendOnBehalfTo @{add = $userAsked1 } }
                2 { Add-ADPermission $userAsked -ExtendedRights Send-As -user $userAsked1 }
                3 { $ehtoLahetys = $false; break; }
              }
            }
          }

          4 {
            Clear-Host;
            write-Host -ForegroundColor Red "
            ####################
            #                  #
            # Poissaoloviestit #
            #                  #
            ####################
            ";
            $ehtoPoissa = $true; 
            while ($ehtoPoissa) {
              Write-Host -ForegroundColor Yellow " Voi käyttää myös automaattisena vastauksena, silloin aseta poissaoloviesti ilman päivämäärää!" `n
              $poissaOlo = Read-Host  "Voit tarkistaa/asettaa henkilön $userAsked poissaoloviestin" `n"1: Tarkista poissaoloviestit " `n"2: Aseta poissaoloviesti " `n"3: Palaa takaisin" `n""
              switch ($poissaOlo) {
                # Tarviiko: Get-Mailbox -ResultSize unlimited | Get-MailboxAutoReplyConfiguration -Identity $userAsked
                1 { Get-MailboxAutoReplyConfiguration -Identity $userAsked }
                2 { 
                  $ehtoViesti = $true;
                  while ($ehtoViesti) {
                    $paivaMaara = Read-Host "Minkä poissaoloviestin haluat lisätä?" `n"1: Päivämäärän & Viestin" `n"2: Pelkästään viestin" <#`n"3: Sisäisen & ulkoisen viestin, sekä päivämäärän"#> `n"3: Palaa takaisin" `n
                    switch ($paivaMaara) {
                      1 {  
                        $aikaAlku = Read-Host "Anna aloitus päivämäärä" `n
                        $aikaAlkuTotal = [datetime]$aikaAlku
                        $aikaAlkuTotal.ToString('dd/MM/yyy')

                        $aikaLoppu = Read-Host "Anna aloitus päivämäärä" `n
                        $aikaLoppuTotal = [datetime]$aikaLoppu
                        $aikaLoppuTotal.ToString('dd/MM/yyy')

                        $poissaOloViesti = Read-Host "Kirjoita poissaoloviesti" `n""
                        Set-MailboxAutoReplyConfiguration -Identity $userAsked -AutoReplyState Scheduled -StartTime $aikaAlkuTotal -EndTime $aikaLoppuTotal -InternalMessage "$poissaOloViesti"  -ExternalMessage "$poissaOloViesti"
                        $ehtoViesti = $false; break;
                      }
                      2 {
                        $poissaOloViesti = Read-Host "Poissaoloviesti tulee käyttöön samantien ja on voimassa toistaiseksi" `n"Kirjoita poissaoloviesti" `n""
                        Set-MailboxAutoReplyConfiguration -Identity $userAsked -AutoReplyState Enabled -InternalMessage "$poissaOloViesti"  -ExternalMessage "$poissaOloViesti" 
                        $ehtoViesti = $false; break;
                      }

                      # Mahdollisuus asettaa erikseen sisäinen ja ulkoinen viesti
                      <# 3 { $aikaAlku = Read-Host "Anna aloitus päivämäärä" `n
                        $aikaAlkuTotal = [datetime]$aikaAlku
                        $aikaAlkuTotal.ToString('dd/MM/yyy')

                        $aikaLoppu = Read-Host "Anna aloitus päivämäärä" `n
                        $aikaLoppuTotal = [datetime]$aikaLoppu
                        $aikaLoppuTotal.ToString('dd/MM/yyy')

                        $poissaOloViesti = Read-Host `n"Sisäinen poissaoloviesti menee vain oman yrityksen sisältä tuleviin sähköposteihin" `n"Kirjoita sisäinen poissaoloviesti" `n""

                        $poissaOloViesti1 = Read-Host `n"Ulkoinen poissaoloviesti menee kaikille säkhöposti osotteille mitkä tulevat yrityksen ulkopuolelta" `n"Kirjoita ulkoinen poissaoloviesti" `n""

                        Set-MailboxAutoReplyConfiguration -Identity $userAsked -AutoReplyState Scheduled -StartTime $aikaAlkuTotal -EndTime $aikaLoppuTotal -InternalMessage "$poissaOloViesti"  -ExternalMessage "$poissaOloViesti1"
                        $ehtoViesti = $false; break; } #>
                      3 { $ehtoViesti = $false; break; }
                    }
                  }
                }
                # Halutessa voidaan koko tenantin poissa olo viestit.
                # 3 { Get-Mailbox -ResultSize unlimited | Get-MailboxAutoReplyConfiguration }
                3 { $ehtoPoissa = $false; break; }
                
              }
            }
          }

          5 {
            Clear-Host;
            Write-Host -ForegroundColor Red "
            ################
            #              #
            # Jakeluryhmät #
            #              #
            ################
            ";
            $ehtoJakelu = $true; 
            while ($ehtoJakelu) {

              $jakelu = Read-Host "Mitkä jakeluryhmät haluat tarkistaa" `n"1: Listaa kaikki jakeluryhmät" `n"2: Lisää $userAsked jakelulistaan" `n"3: Poista $userAsked jakelulistasta" `n"4: Palaa takaisin" `n""

              switch ($jakelu) {
                1 { Get-DistributionGroup | Select-Object Name, DisplayName, PrimarySmtpAddress }
                2 {
                  $lisaaLista = (Get-DistributionGroup | Select-Object Name)
                  $menu_single = @{}
                        
                  for ($i = 1; $i -le $lisaaLista.Count; $i++) {
                    Write-host "$i. $($lisaaLista[$i-1].Name)"
                    $menu_single.Add($i, ($lisaaLista[$i - 1].Name))
                  } do {
                    [int]$ans = Read-Host "Valitse listasta ryhmä mihin $userAsked lisätään (numero)" `n""
                  } until ([int]$ans -ige 1 -and [int]$ans -ile $lisaaLista.length)
                  $selectionLisaa = $menu_single.Item($ans)

                  Add-DistributionGroupMember -Identity $selectionLisaa -Member $userAsked 
                }
                3 {  
                  $poistoLista = (Get-DistributionGroup | Select-Object Name)
                  $menu_single = @{}
                        
                  for ($i = 1; $i -le $poistoLista.Count; $i++) {
                    Write-host "$i. $($poistoLista[$i-1].Name)"
                    $menu_single.Add($i, ($poistoLista[$i - 1].Name))
                  } do {
                    [int]$ans = Read-Host "Valitse listasta ryhmä mistä $userAsked poistetaan (numero)" `n""
                  } until ([int]$ans -ige 1 -and [int]$ans -ile $poistoLista.length)
                  $selectionPoisto = $menu_single.Item($ans)

                  Remove-DistributionGroupMember -Identity $selectionPoisto -Member $userAsked
                }
                4 { $ehtoJakelu = $false; break; }
              }
            }
          }

          6 {
            Clear-Host;
            Write-Host -ForegroundColor Red "                                            
            ################
            #              #
            # Tarkistukset #
            #              #
            ################
            ";
            $ehtoTarkistus = $true; 
            
            # Tarvitseeko ErrorActionPreference? ei näytä erroreita jos ei löydy kyseistä kansiota.
            $ErrorActionPreference = "SilentlyContinue";

            while ($ehtoTarkistus) { 

              $oikeuksienTarkistus = Read-Host "Voit tarkistaa seuraavat oikeudet" `n"1: Kalenterin" `n"2: Postilaatikon" `
                `n"3: Lähetys" `n"4: Poissaoloviestit" `n"5: Jakeluryhmät" `n"6: Palaa takaisin" `n""

              switch ($oikeuksienTarkistus) {
                
                # Hae laatikoiden tiedot alla olevalla koodilla!
                <#
                $obj = Get-MailboxFolderStatistics -Identity $userAsked -FolderScope  | Where-Object { $_.FolderType -eq 'folderName' } | Select-Object Name, FolderId 
                $obj1 = Get-MailboxFolderPermission -Identity $userAsked":\$($folderName.Name)" 
                ($obj1 | Format-table -Force | Out-String) 
                #>


                1 { 
                  $calendar = Get-MailboxFolderStatistics -Identity $userAsked -FolderScope Calendar | Where-Object { $_.FolderType -eq 'Calendar' } | Select-Object Name, FolderId 
                  $permCalendar = Get-MailboxFolderPermission -Identity $userAsked":\$($calendar.Name)" 
                  ($permCalendar | Format-table -Force | Out-String) 
                }
                2 { 
                  $Inbox = Get-MailboxFolderStatistics -Identity $userAsked -FolderScope Inbox | Where-Object { $_.FolderType -eq 'Inbox' } | Select-Object Name, FolderId 
                  $permInbox = Get-MailboxFolderPermission -Identity $userAsked":\$($Inbox.Name)" 
                  ($permInbox | Format-table -Force | Out-String) 
                }
                3 {  
                  $outBox = Get-MailboxFolderStatistics -Identity $userAsked -FolderScope OutBox | Where-Object { $_.FolderType -eq 'OutBox' } | Select-Object Name, FolderId 
                  $permOutBox = Get-MailboxFolderPermission -Identity $userAsked":\$($outBox.Name)" 
                  ($permOutBox | Format-table -Force | Out-String) 

                }
                4 { 
                  Get-MailboxAutoReplyConfiguration -Identity $userAsked
                }
                5 {                       
                  $jakeluLista = (Get-DistributionGroup -ResultSize unlimited | Select-Object name)
                  $menu_single = @{}
      
                  for ($i = 1; $i -le $jakeluLista.Count; $i++) {
                    Write-host "$i. $($jakeluLista[$i-1].name)"
                    $menu_single.Add($i, ($jakeluLista[$i - 1].name))
                  } do {
                    [int]$ans = Read-Host "Valitse minkä ryhmän haluat tarkistaa" `n""
                  } until ([int]$ans -ige 1 -and [int]$ans -ile $jakeluLista.length)
                  $selectionJakelu = $menu_single.Item($ans)

                  $jakeluListat = Get-DistributionGroupMember -Identity $selectionJakelu | Select-Object Name, RecipientType | Format-Table | Out-string
                  Write-Host $jakeluListat
                }
                6 { $ehtoTarkistus = $false; break; }
              }

            }
          }
          7 { $valintaKayttaja = $false; break; } 
        }
      }
    }


    2 { <#Multitenant#> }


    3 { $ehtoKoko = $False; Break; }
    
  }

}
