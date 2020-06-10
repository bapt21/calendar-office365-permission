

#On se connecte a la plateforme

$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session 

##################################################################################################################################
#                                                                                                                                #
# On l'exécute pour tous les utilisateurs, avec une double commande au cas ou on aurait des boites mails toujours en anglais.    #
# D'abords on exporte le résultat avant application du script avec une redirection de la sortie.                                 #
# On retire toutes les permissions sur les calendriers sinon ca refusera de les modifiers.                                       #
# On attribue le niveau de permissions qui nous intéresses.                                                                      #
# On exporte le résultat dans un autre fichiers pour comparer les résultats.                                                     #
#                                                                                                                                #
##################################################################################################################################                    
    
    
$users = Get-MailBox | Select -ExpandProperty Alias


#Foreach ($user in $users) {Get-MailboxFolderPermission $user":\Calendrier" >> c:\sources\avantmofification.txt }
#Foreach ($user in $users) {Get-MailboxFolderPermission $user":\Calendar" >> c:\sources\avantmofification.txt }

#Foreach ($user in $users) {Remove-MailboxFolderPermission $user":\Calendrier" -user $user -Confirm:$False}
#Foreach ($user in $users) {Remove-MailboxFolderPermission $user":\Calendar" -user $user -Confirm:$False}

#Foreach ($user in $users) {Add-MailboxFolderPermission $user":\Calendrier" -User $user -accessrights Reviewer}
#Foreach ($user in $users) {Add-MailboxFolderPermission $user":\Calendar" -User $user -accessrights Reviewer}

#Foreach ($user in $users) {Get-MailboxFolderPermission $user":\Calendrier" >> c:\sources\apresmodificaion.txt }
#Foreach ($user in $users) {Get-MailboxFolderPermission $user":\Calendar" >> c:\sources\apresmodification.txt }

#############################################
#                                           #  
#      ZONE DE TEST AVEC UN UTILISATEUR     #
#                                           #
#############################################

Foreach ($user in $users) {Get-MailboxFolderPermission test.test@test.fr:\Calendrier >> c:\sources\avantmofification.txt }
Foreach ($user in $users) {Remove-MailboxFolderPermission test.test@test.fr:\Calendrier -user $user -Confirm:$False}
Foreach ($user in $users) {Add-MailboxFolderPermission test.test@test.fr:\Calendrier -User $user -accessrights Reviewer}
Foreach ($user in $users) {Get-MailboxFolderPermission test.test@test.fr:\Calendrier >> c:\sources\apresmodificaion.txt }


#On ferme la session.

Remove-PSSession $Session

