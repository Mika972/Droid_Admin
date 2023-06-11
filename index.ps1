##                                                                       ##
 #====================Programme principal de DroidAdminV3================#
  ##                                                                   ##

#Include les bibliotheques et fonctions du programme
. "C:\Developpement\DroidAdminV3\applications\biblio.ps1"

$formDroid = New-Object System.Windows.Forms.Form
$tabcontrolDroid = New-Object System.Windows.Forms.TabControl
$tabpage1 = New-Object System.Windows.Forms.TabPage
$tabpage2 = New-Object System.Windows.Forms.TabPage
$tabpage3 = New-Object System.Windows.Forms.TabPage
$tabpage4 = New-Object System.Windows.Forms.TabPage
$bouton_Sortie = New-Object System.Windows.Forms.Button
#
# Form du Droid
#
$formDroid.ClientSize = New-Object System.Drawing.Size(860, 720)
$formDroid.Controls.Add($tabcontrolDroid)
$formDroid.Controls.Add($bouton_Sortie)
$formDroid.Name = 'DroidAdminV3'
$formDroid.Text = 'DroidAdminV3'
$formDroid.Opacity = 0.97 #Pour l'opacité comme en CSS
#
# Changement d'icone de l'application avec l'appel de la fonction r2icone
#$icon = [system.drawing.icon]::ExtractAssociatedIcon("C:\Windows\System32\newdev.exe")
#$formDroid.Icon = $icon
r2icone
###
# Appel de la fonction imageDeFondDroid/couleurDeFondDroid qui va charger l'image de fond de l'application
#imageDeFondDroid/couleurDeFondDroid
couleurDeFondDroid
# Appel des fonctions imageFondTabcontrol(n) qui changent l'image de fond des tabpage(n)
imageFondTabpage1
imageFondTabpage2
imageFondTabpage3
imageFondTabpage4
#
# tabcontrol principal du Droid
#
$tabcontrolDroid.Controls.Add($tabpage1)
$tabcontrolDroid.Controls.Add($tabpage2)
$tabcontrolDroid.Controls.Add($tabpage3)
$tabcontrolDroid.Controls.Add($tabpage4)
$tabcontrolDroid.Name = 'tabcontrolDroid'
$tabcontrolDroid.Location = New-Object System.Drawing.Point(24,12)
$tabcontrolDroid.Size = New-Object System.Drawing.Size(810, 630)
$tabcontrolDroid.TabIndex = 0
#
# tabpage1
#
$tabpage1.location = '42, 4'
$tabpage1.Name = 'tabpage1'
$tabpage1.Padding = '3, 3, 3, 3'
$tabpage1.Size = '602, 442'
$tabpage1.TabIndex = 1
$tabpage1.UseVisualStyleBackColor = $True
#
# tabpag 2
#
$tabpage2.Location = '23, 4'
$tabpage2.Name = 'tabpage2'
$tabpage2.Padding = '3, 3, 3, 3'
$tabpage2.Size = '602, 442'
$tabpage2.TabIndex = 2
$tabpage2.Text = 'tabpage2'
$tabpage2.UseVisualStyleBackColor = $True
#
# tabpage3
#
$tabpage3.Location = '23, 4'
$tabpage3.Name = 'tabpage3'
$tabpage3.Padding = '3, 3, 3, 3'
$tabpage3.Size = '602, 442'
$tabpage3.TabIndex = 3
$tabpage3.Text = 'tabpage3'
$tabpage3.UseVisualStyleBackColor = $True
#
# tabpage4
#
$tabpage4.Location = '23, 4'
$tabpage4.Name = 'tabpage3'
$tabpage4.Padding = '3, 3, 3, 3'
$tabpage4.Size = '602, 442'
$tabpage4.TabIndex = 4
$tabpage4.Text = 'tabpage3'
$tabpage4.UseVisualStyleBackColor = $True
#
# bouton quitter
#
$bouton_Sortie.Location = New-Object System.Drawing.Point(708, 665)
$bouton_Sortie.Name = 'button 3'
$bouton_Sortie.Size = New-Object System.Drawing.Size(126, 30)
$bouton_Sortie.TabIndex = 5
$bouton_Sortie.Text = 'Quitter'
$bouton_Sortie.UseVisualStyleBackColor = $true
$bouton_Sortie.Add_Click({$formDroid.Close()})
##=====================================================================================================##
# 							         Include des tabpages(n)											#
##=====================================================================================================##

# Include de InfoUCVEx2, qui donne les propriétés, crée et remplis le fichier excel au choix
. "C:\Developpement\DroidAdminV3\applications\proprietesUC.ps1"

# Include de ListComputerSearchV4, qui récupère les infos de toute une liste de PCs
. "C:\Developpement\DroidAdminV3\applications\ListproprietesUCs.ps1"

# Include de CreateKBLogV1, qui récupère les KB microsoft d'un PC sous forme d'un log fichier txt
. "C:\Developpement\DroidAdminV3\applications\KBLog.ps1"

# Include de forest, qui récupère tous les PC FWPC de l'AD
. "C:\Developpement\DroidAdminV3\applications\forest.ps1"
#=======================================================================================================#
$formDroid.ShowDialog()
