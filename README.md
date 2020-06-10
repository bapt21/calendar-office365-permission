# calendar-office365-permission
Modify calendar permission

Script en powershell permettant d'extraire les droits sur un calendrier d'un utilisateur et de le partager a toute l'entreprise.


On l'exécute pour tous les utilisateurs, avec une double commande au cas ou on aurait des boites mails toujours en anglais.    
D'abords on exporte le résultat avant application du script avec une redirection de la sortie.                                 
On retire toutes les permissions sur les calendriers sinon ca refusera de les modifiers.                                      
On attribue le niveau de permissions qui nous intéresses.                                                                    
On exporte le résultat dans un autre fichiers pour comparer les résultats.      



Powershell script to extract the rights to a user's calendar and share it with the whole company.
We run it for all users, with a double command in case we have mailboxes still in English or French.    
First we export the result before applying the script with an output redirection.                                 
We remove all permissions on the calendars otherwise it will refuse to modify them.                                      
We assign the level of permissions that we are interested in.                                                                    
We export the result in another file to compare the results.     

