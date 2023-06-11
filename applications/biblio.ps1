##########
#===================================================Bibliotheques et fonctions communes a toutes les applications,
#==================================================================Image du Background de $formDroid
##########
[void][reflection.assembly]::Load('System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
[void][reflection.assembly]::Load('System.Data, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
[void][reflection.assembly]::Load('System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
[void][reflection.assembly]::Load('System.DirectoryServices, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
[void][reflection.assembly]::Load('System.ServiceProcess, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
													#============================================================#
#
# La fonction r2icone permet de changer l'icone de l'application
function r2icone {
	$formDroid.Icon = [system.drawing.icon]::ExtractAssociatedIcon("C:\Windows\System32\newdev.exe")
}
													#============================================================#
##
#	La fonction imageDeFondDroid contient l'image de fond de $formDroid
function imageDeFondDroid	{
	$image = [system.drawing.image]::FromFile("C:\Developpement\DroidAdminV3\images\stars.jpg")
	$formDroid.BackgroundImage = $image
	$formDroid.BackgroundImageLayout = "None" # None, Tile, Center, Stretch, Zoom
	#$formDroid.Width = $image.Width
	#$formDroid.Height = $Image.Height
}
													#============================================================#
###
#	La fonction couleurDeFondDroid configure la couleur de fond de $formDroid
function couleurDeFondDroid {
	$formDroid.BackColor = "black"
}
													#============================================================#
####
# 	La fonction excelSpecial cree une page excel avec ses colonnes. De plus, des variables globales sont crees
# 	afin d'augmenter la portee de certaines variables utiles dans le reste du programme

function excelSpecial {
	$a = New-Object -comobject Excel.Application	#		Declaration d'Excel		#

	$a.Visible = $True

	$b = $a.Workbooks.Add()
	$global:c = $b.Worksheets.Item(1)

	$global:col = 1
	$global:ligne = 1

	$c.Cells.Item(1,1) = "Poste de travail"
	$c.Cells.Item(1,2) = "Utilisateur"
	$c.Cells.Item(1,3) = "Modele"
	$c.Cells.Item(1,4) = "N de serie"
	$c.Cells.Item(1,5) = "Memoire"
	$c.Cells.Item(1,6) = "Mp"
	$c.Cells.Item(1,7) = "Date"
}
													#============================================================#
#####
#	Les fonctions imageFondTabcontrol(n) changent l'image de fond des Tabpage(n)

function imageFondTabpage1 {
	$imagTabpage1 = [system.drawing.image]::FromFile("C:\Developpement\DroidAdminV3\images\R2D2_2.jpg")
	$tabpage1.BackgroundImage = $imagTabpage1
	$tabpage1.BackgroundImageLayout = "Stretch"
}
function imageFondTabpage2 {
	$imagTabpage2 = [system.drawing.image]::FromFile("C:\Developpement\DroidAdminV3\images\droidSpaceShip.jpg")
	$tabpage2.BackgroundImage = $imagTabpage2
	$tabpage2.BackgroundImageLayout = "Stretch"
}
function imageFondTabpage3 {
	$imagTabpage3 = [system.drawing.image]::FromFile("C:\Developpement\DroidAdminV3\images\bb8_4k.jpg")
	$tabpage3.BackgroundImage = $imagTabpage3
	$tabpage3.BackgroundImageLayout = "Stretch"
}
function imageFondTabpage4 {
	$imagTabpage4 = [system.drawing.image]::FromFile("C:\Developpement\DroidAdminV3\images\droidForest.jpg")
	$tabpage4.BackgroundImage = $imagTabpage4
	$tabpage4.BackgroundImageLayout = "Stretch"
}
