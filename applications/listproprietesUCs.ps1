###
# ListComputerSearchV4.ps1			test	01/08/2018
###
$label_postes = New-Object System.Windows.Forms.Label
$label_messages = New-Object System.Windows.Forms.Label
$bouton_demarrer =  New-Object System.Windows.Forms.Button
$bouton_reset	=	New-Object System.Windows.Forms.Button
$progressBarP = New-Object System.Windows.Forms.ProgressBar
#
#	$tabpage2
#
$tabpage2.Controls.Add($label_postes)
$tabpage2.Controls.Add($label_messages)
$tabpage2.Controls.Add($bouton_demarrer)
$tabpage2.Controls.Add($bouton_reset)
$tabpage2.Controls.Add($progressBarP)
$tabpage2.Name = 'ListDesUCs'
$tabpage2.Text = 'Liste proprietes UCs'
#
# label verification de la liste de postes
#
$label_postes.AutoSize = $true
$label_postes.Location = New-Object System.Drawing.Point(60, 31)
$label_postes.Name = 'label_postes'
$label_postes.Size = New-Object System.Drawing.Size(79, 13)
$label_postes.TabIndex = 0
$label_postes.Text = 'Verifier la liste de postes'
$label_postes.ForeColor = "Yellow"
$label_postes.Font = 'Segoe UI, 10.75pt, style=Bold, Italic' 
#
# label 2
#
$label_messages.AutoSize = $true
$label_messages.Location = New-Object System.Drawing.Point(80, 125)
$label_messages.Name = 'label_messages'
$label_messages.Size = New-Object System.Drawing.Size(79, 13)
$label_messages.TabIndex = 2
$label_messages.Text = 'En attente'
$label_messages.ForeColor = "Turquoise"
#$label_messages.Font = 'Segoe UI, 10.75pt, style=Bold, Italic' 
#
# Barre de progression
#
$progressBarP.Location = New-Object System.Drawing.Point(30, 150)
$progressBarP.Name = 'progressBarP'
$progressBarP.Size = New-Object System.Drawing.Size(220, 23)
$progressBarP.Value = 0
$progressBarP.Style = "continuous"
#
# bouton demarrer
#
$bouton_demarrer.Location = New-Object System.Drawing.Point (35, 75)
$bouton_demarrer.Name = 'Bouton demarrer'
$bouton_demarrer.Size = New-Object System.Drawing.Size(90 ,35)
$bouton_demarrer.TabIndex = 1
$bouton_demarrer.Text = 'Demarrer'
$bouton_demarrer.UseVisualStyleBackColor = $true
$bouton_demarrer.Add_Click(
	{
	#Moteur_DEBUT#################################################################
	$path = "C:\TEMP\Outils\Labo\DroidAdminV3\results.xls"
	#Excel
	$objExcel = new-object -comobject excel.application

	if (Test-Path $path) {
			$objWorkbook = $objExcel.WorkBooks.Open($path)
			$objWorksheet = $objWorkbook.Worksheets.Item(1)
			}
	else {
			$objWorkbook = $objExcel.Workbooks.Add()
			$objWorksheet = $objWorkbook.Worksheets.Item(1)
			}

	$objExcel.Visible = $True

	#########Add Header de la page excel#########

	$objWorksheet.Cells.Item(1, 1) = "Machine"
	$objWorksheet.Cells.Item(1, 2) = "Connecte"
	$objWorksheet.Cells.Item(1, 3) = "Modele"
	$objWorksheet.Cells.Item(1, 4) = "N de serie"
	$objWorksheet.Cells.Item(1, 5) = "Memoire"
	$objWorksheet.Cells.Item(1, 6) = "Mp"
	$objWorksheet.Cells.Item(1, 7) = "Date"
	$objWorksheet.Cells.Item(1, 8) = "Ping"

	$machines = gc C:\Developpement\DroidAdminV3\listMachines.txt
	$countMch = $machines.count #N'est pas utilisé dans le reste du code
	$row=2

	$top = Get-Content "C:\Developpement\DroidAdminV3\listMachines.txt"
	$i = 0 #Initialisation du compteur

	$machines | foreach-object{
		$ping = $null
		$loginfo = $null
		$machine = $_
		$ping = Test-Connection $machine -Count 1 -ea silentlycontinue

		if($ping){
			$objWorksheet.Cells.Item($row,1) = $machine
			$objWorksheet.Cells.Item($row,8) = "OUI"

			$loginfo = Get-WmiObject -Computer $machine -Class Win32_ComputerSystem
			$objWorksheet.Cells.Item($row,2) = $loginfo.UserName

			$model = Get-WmiObject -Class Win32_computerSystem -NameSpace "root\CIMV2" -Computer "$machine"
			$objWorksheet.Cells.Item($row,3) = $model.Model

			$snumber = Get-WmiObject -Class Win32_BIOS -NameSpace "root\CIMV2" -Computer "$machine"
			$objWorksheet.Cells.Item($row,4) = $snumber.SerialNumber

			$memory = Get-WmiObject -Class Win32_computerSystem -NameSpace "root\CIMV2" -Computer "$machine"
			$objWorksheet.Cells.Item($row,5) = $memory.TotalPhysicalMemory

			$mp = Get-WmiObject -Class Win32_Processor -Computer "$machine"
			$objWorksheet.Cells.Item($row,6) = $mp.Name

			$date = (get-date).tostring('dd/MM/yyyy')
			$objWorksheet.Cells.Item($row,7) = $date
			$row++}
		else {
			$objWorksheet.Cells.Item($row,1) = $machine
			$objWorksheet.Cells.Item($row,8) = "non"
			$row++}
		#Calcul en pourcentage
		$i++
		[int]$pct = ($i/$machines.count)*100 #On aurait pu utiliser $top à la place de $machines
		#Mise à jour de la barre de progression
		$progressBarP.Value = $pct
		$label_messages.text = "Ordinateur: $machine"
		#Rafraichissement de la windows Fomr5 graphique, à chaque changement de l'état de progressBar1
		#$Form5.Refresh()
		}
	#Moteur_FIN#################################################################
	}
)
#
# Bouton reset
#
$bouton_reset.Location = New-Object System.Drawing.Point (152, 75)
$bouton_reset.Name = 'Bouton demarrer'
$bouton_reset.Size = New-Object System.Drawing.Size(90 ,35)
$bouton_reset.TabIndex = 2
$bouton_reset.Text = 'Reset'
$bouton_reset.UseVisualStyleBackColor = $true
$bouton_reset.Add_Click(
	{
		$progressBarP.Value = 0		
		$label_messages.text = 'En attente'
	}
)
