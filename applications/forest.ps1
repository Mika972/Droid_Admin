##
# forestV1	Test 02/08/2018
##
$textBox_UCf = New-Object System.Windows.Forms.TextBox
$bouton_recherchef = New-Object System.Windows.Forms.Button
#$button_clearf = New-Object System.Windows.Forms.Button
#$button_quitterf = New-Object System.Windows.Forms.Button
$label_posteOU = New-Object System.Windows.Forms.Label
#
# tabpage4
#
#$tabpage4.Controls.Add($textBox_UCf)
$tabpage4.Controls.Add($bouton_recherchef)
$tabpage4.Controls.Add($label_posteOU)
$tabpage4.Name = 'Forest'
$tabpage4.Text = 'Forest'
#
# label_posteOU
#
$label_posteOU.AutoSize = $true
$label_posteOU.Location = New-Object System.Drawing.Point(55, 31)
$label_posteOU.Name = 'label_posteOU'
$label_posteOU.Size = New-Object System.Drawing.Size(79, 13)
$label_posteOU.TabIndex = 0
$label_posteOU.Text = 'Liste ordinateurs'
$label_posteOU.ForeColor = "Yellow"
$label_posteOU.Font = 'Segoe UI, 10.75pt, style=Bold, Italic' 
#
# textBox_UCf
#
#$textBox_UCf.Location = New-Object System.Drawing.Point(211, 31)
#$textBox_UCf.Multiline = $true
#$textBox_UCf.Name = 'textBox_UC'
#$textBox_UCf.Size = New-Object System.Drawing.Size(270, 20)
#$textBox_UCf.TabIndex = 1
#
# bouton valider 
#
$bouton_recherchef.Location = New-Object System.Drawing.Point (25, 90)
$bouton_recherchef.Name = 'Bouton valider'
$bouton_recherchef.Size = New-Object System.Drawing.Size(173 ,35)
$bouton_recherchef.TabIndex = 1
$bouton_recherchef.Text = 'Creation liste UC'
$bouton_recherchef.UseVisualStyleBackColor = $true
$bouton_recherchef.Add_Click(
	{
	$date = (Get-Date).ToString('ddMMyyyy')
	Get-ADComputer -Filter {name -like"FWPC-*"}|out-file C:\TEMP\Outils\Anomalies_kb\"listeComputers"$date.txt		
	}
)

# Vérifie si le fichier.INI contient la ligne en question dans le fichier 
# Get-Content "C:\Program Files (x86)\dossier\fichier.INI" | Where-Object { $_.Contains("la_ligne_en_question") }
