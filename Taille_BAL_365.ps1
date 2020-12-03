param(
[Parameter(Mandatory = $false)]
 [switch]$MFA,
 [switch]$SharedMBOnly,
 [switch]$UserMBOnly,
 [string]$MBNamesFile,
 [string]$UserName,
 [string]$Password
 )

Function Get_MailboxSize
{
 $Stats=Get-MailboxStatistics -Identity $UPN
 $ItemCount=$Stats.ItemCount
 $TotalItemSize=$Stats.TotalItemSize
 $TotalItemSizeinBytes= $TotalItemSize –replace “(.*\()|,| [a-z]*\)”, “”
 $TotalSize=$stats.TotalItemSize.value -replace "\(.*",""
 $DeletedItemCount=$Stats.DeletedItemCount
 $TotalDeletedItemSize=$Stats.TotalDeletedItemSize
 $TotalDeletedItemSizeinBytes=$TotalDeletedItemSize –replace “(.*\()|,| [a-z]*\)”, “”
 $TotalDeletedSize=$stats.TotalDeletedItemSize.value -replace "\(.*",""  
 
 #Export du résultat en CSV
 $Result=@{'Nom de la BAL'=$DisplayName;'Type de BAL'=$MailboxType;'Nb éléments'=$ItemCount;'Taille Totale'=$TotalSize;'Taille Totale en Bytes'=$TotalItemSizeinBytes;'Nb éléments supprimés'=$DeletedItemCount;'Taille Corbeille'=$TotalDeletedSize;'Taille Corbeille en Bytes'=$TotalDeletedItemSizeinBytes}
 $Results= New-Object PSObject -Property $Result  
 $Results | Select-Object 'Nom de la BAL','Type de BAL','Nb éléments','Taille Totale','Taille Totale en Bytes','Nb éléments supprimés','Taille Corbeille','Taille Corbeille en Bytes' | Export-Csv -Path $ExportCSV -Notype -Append -Encoding UTF8
}

Function main()
{
 #Vérification présence du module EXO v2
 $Module = Get-Module ExchangeOnlineManagement -ListAvailable
 if($Module.count -eq 0) 
 { 
  Write-Host Exchange Online PowerShell V2 non disponible  -ForegroundColor yellow  
  $Confirm= Read-Host Installer le module? [O] Oui [N] Non 
  if($Confirm -match "[Oo]") 
  { 
   Write-host "Installation du module Exchange Online Powershell v2"
   Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force
  } 
  else 
  { 
   Write-Host Le module EXO V2 est nécessaire pour se connecter à Exchange Online.installez ce module en utilisant la commande Install-Module ExchangeOnlineManagement. 
   Exit
  }
 } 

 #Connexion avec MFA
 if($MFA.IsPresent)
 {
  Connect-ExchangeOnline
 }

 #Autehtification sans MFA
 else
 {
  #Stockage du credential
  if(($UserName -ne "") -and ($Password -ne ""))
  {
   $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
   $Credential  = New-Object System.Management.Automation.PSCredential $UserName,$SecuredPassword
  }
  else
  {
   $Credential=Get-Credential -Credential $null
  }
  Connect-ExchangeOnline -Credential $Credential
 }

 #Déclaration fichier de sorti 
 
 $ExportCSV=".\MailboxSizeReport_$((Get-Date -format ddd-dd-MMM-yyyy` hh-mm` tt).ToString()).csv" 

 $Result=""   
 $Results=@()  
 $MBCount=0
 $PrintedMBCount=0
 Write-Host Génération du Rapport de BAL...
 
 
 if([string]$MBNamesFile -ne "") 
 { 
  #Mise en mémoire des données 
  $Mailboxes=@()
  $Mailboxes=Import-Csv -Header "MBIdentity" $MBNamesFile
  foreach($item in $Mailboxes)
  {
   $MBDetails=Get-Mailbox -Identity $item.MBIdentity
   $UPN=$MBDetails.UserPrincipalName  
   $MailboxType=$MBDetails.RecipientTypeDetails
   $DisplayName=$MBDetails.DisplayName
   $MBCount++
   Write-Progress -Activity "`n     Nb de BAL scannées: $MBCount "`n"  Scan en cours de: $DisplayName"
   Get_MailboxSize
   $PrintedMBCount++
  }
 }

 #Récupération des informations des BAL sur 365
 else
 {
  Get-Mailbox -ResultSize Unlimited | foreach {
   $UPN=$_.UserPrincipalName
   $Mailboxtype=$_.RecipientTypeDetails
   $DisplayName=$_.DisplayName
   $MBCount++
   Write-Progress -Activity "`n     Nb de BAL scannées: $MBCount "`n"  Analyse de la BAL: $DisplayName"
   if($SharedMBOnly.IsPresent -and ($Mailboxtype -ne "SharedMailbox"))
   {
    return
   }
   if($UserMBOnly.IsPresent -and ($MailboxType -ne "UserMailbox"))
   {
    return
   }  
   Get_MailboxSize
   $PrintedMBCount++
  }
 }

 #Ouverture du fichier de sortie 
 If($PrintedMBCount -eq 0)
 {
  Write-Host BAL non trouvée
 }
 else
 {
  Write-Host `nLe fichier de sortie contient $PrintedMBCount BAL.
  if((Test-Path -Path $ExportCSV) -eq "True") 
  {
   Write-Host `nFichier de sortie stocké vers $ExportCSV -ForegroundColor Green
  }
 }
 #Deconnexion des session exchange online
 Disconnect-ExchangeOnline -Confirm:$false | Out-Null
}
 . main

