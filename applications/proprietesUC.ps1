##
# InfoUCVEx3		test 01/08/2018
##
#D�claration des variables d'index
$textBox_UCx = New-Object System.Windows.Forms.TextBox
$textBox_marquex = New-Object System.Windows.Forms.TextBox
$textBox_SNumberx = New-Object System.Windows.Forms.TextBox
$textBox_Memoryx = New-Object System.Windows.Forms.TextBox
$textBox_connectex = New-Object System.Windows.Forms.TextBox
$textBox_mpx = New-Object System.Windows.Forms.TextBox
$textBox_Datex = New-Object System.Windows.Forms.TextBox
$bouton_recherchex = New-Object System.Windows.Forms.Button
$bouton_resetx = New-Object System.Windows.Forms.Button
$label_poste = New-Object System.Windows.Forms.Label
$label_cnt = New-Object System.Windows.Forms.Label
$label_mod = New-Object System.Windows.Forms.Label
$label_sn = New-Object System.Windows.Forms.Label
$label_my = New-Object System.Windows.Forms.Label
$label_mp = New-Object System.Windows.Forms.Label
$label_date = New-Object System.Windows.Forms.Label
$bouton_excelx = New-Object System.Windows.Forms.Button
#
# tabpage1
#
$tabpage1.Controls.Add($label_poste)
$tabpage1.Controls.Add($label_cnt)
$tabpage1.Controls.Add($label_mod)
$tabpage1.Controls.Add($label_sn)
$tabpage1.Controls.Add($label_my)
$tabpage1.Controls.Add($label_mp)
$tabpage1.Controls.Add($label_date)
$tabpage1.Controls.Add($bouton_resetx)
$tabpage1.Controls.Add($bouton_recherchex)
$tabpage1.Controls.Add($bouton_excelx)
$tabpage1.Controls.Add($textBox_marquex)
$tabpage1.Controls.Add($textBox_SNumberx)
$tabpage1.Controls.Add($textBox_UCx)
$tabpage1.Controls.Add($textBox_Memoryx)
$tabpage1.Controls.Add($textBox_connectex)
$tabpage1.Controls.Add($textBox_mpx)
$tabpage1.Controls.Add($textBox_Datex)
$tabpage1.Name = 'FormPro'
$tabpage1.Text = 'Proprietes UC'
#
# textBox_UCx
#
$textBox_UCx.Location = New-Object System.Drawing.Point(270, 31)
$textBox_UCx.Multiline = $true
$textBox_UCx.Name = 'textBox_UC'
$textBox_UCx.Size = New-Object System.Drawing.Size(270, 20)
$textBox_UCx.TabIndex = 1
#
# textBox Utilisateur connecte
#
$textBox_connectex.Location = New-Object System.Drawing.Point(270, 81)
$textBox_connectex.Multiline = $true
$textBox_connectex.Name = 'textBox_connectex'
$textBox_connectex.Size = New-Object System.Drawing.Size(270, 20)
$textBox_connectex.TabIndex = 3
#
# textBox mod�le
#
$textBox_marquex.Location = New-Object System.Drawing.Point (270, 131)
$textBox_marquex.Multiline = $true
$textBox_marquex.Name = 'textBox_marque'
$textBox_marquex.Size = New-Object System.Drawing.Point (270, 20)
$textBox_marquex.TabIndex = 5
#
# textBox Num�ro de s�rie
#
$textBox_SNumberx.Location = New-Object System.Drawing.Point (270, 181)
$textBox_SNumberx.Multiline = $true
$textBox_SNumberx.Name = 'textBox_SNumber'
$textBox_SNumberx.Size = New-Object System.Drawing.Size (270, 20)
$textBox_SNumberx.TabIndex = 7
#
# textBox M�moire
#
$textBox_Memoryx.Location = New-Object System.Drawing.Point (270, 231)
$textBox_Memoryx.Multiline = $true
$textBox_Memoryx.Name = 'textBox_Memory'
$textBox_Memoryx.Size = New-Object System.Drawing.Point (270, 20)
$textBox_Memoryx.TabIndex = 9
#
# textBox � processeur
#
$textBox_mpx.Location = New-Object System.Drawing.Point (270, 281)
$textBox_mpx.Multiline = $true
$textBox_mpx.Name = 'textBox_mp'
$textBox_mpx.Size = New-Object System.Drawing.Size (270, 20 )
$textBox_mpx.TabIndex = 11
#
# textBox_Date
#
$textBox_Datex.Location = New-Object System.Drawing.Point (270, 331)
$textBox_Datex.Multiline = $true
$textBox_Datex.Name = 'textBox_Date'
$textBox_Datex.Size = New-Object System.Drawing.Size (270, 20)
$textBox_Datex.TabIndex = 13
#
# label_poste
#
$label_poste.AutoSize = $true
$label_poste.Location = New-Object System.Drawing.Point(80, 31)
$label_poste.Name = 'label_poste'
$label_poste.Size = New-Object System.Drawing.Size(79, 13)
$label_poste.TabIndex = 0
$label_poste.Text = 'Poste de travail'
$label_poste.ForeColor = "Yellow" # Couleur du texte
$label_poste.Font = 'Segoe UI, 10.75pt, style=Bold, Italic' # Pour changer le Font du texte
#
# label utilisateur connect�
#
$label_cnt.AutoSize = $true
$label_cnt.Location = New-Object System.Drawing.Point(80, 81)
$label_cnt.Name = 'label_cnt'
$label_cnt.Size = New-Object System.Drawing.Size(79, 13)
$label_cnt.TabIndex = 2
$label_cnt.Text = 'Utilisateur'
$label_cnt.ForeColor = "Yellow"
$label_cnt.Font = 'Segoe UI, 10.75pt, style=Bold, Italic'
#
# label mod�le
#
$label_mod.AutoSize = $true
$label_mod.Location = New-Object System.Drawing.Point(80, 131)
$label_mod.Name = 'label_mod'
$label_mod.Size = New-Object System.Drawing.Size(79, 13)
$label_mod.TabIndex = 4
$label_mod.Text = 'Model'
$label_mod.ForeColor = "Yellow"
$label_mod.Font = 'Segoe UI, 10.75pt, style=Bold, Italic'
#
# label_sn
#
$label_sn.AutoSize = $true
$label_sn.Location = New-Object System.Drawing.Point(80, 181)
$label_sn.Name = 'label_sn'
$label_sn.Size = New-Object System.Drawing.Size(79, 13)
$label_sn.TabIndex = 6
$label_sn.Text = 'N de serie'
$label_sn.ForeColor = "Yellow"
$label_sn.Font = 'Segoe UI, 10.75pt, style=Bold, Italic'
#
# label_my
#
$label_my.AutoSize = $true
$label_my.Location = New-Object System.Drawing.Point(80, 231)
$label_my.Name = 'label_my'
$label_my.Size = New-Object System.Drawing.Size(79, 13)
$label_my.TabIndex = 8
$label_my.Text = 'Memoire'
$label_my.ForeColor = "Yellow"
$label_my.Font = 'Segoe UI, 10.75pt, style=Bold, Italic'
#
# label micro processeur
#
$label_mp.AutoSize = $true
$label_mp.Location = New-Object System.Drawing.Point(80, 281)
$label_mp.Name = 'labe_mp'
$label_mp.Size = New-Object System.Drawing.Size(79, 13)
$label_mp.TabIndex = 10
$label_mp.Text = 'mp'
$label_mp.ForeColor = "Yellow"
$label_mp.Font = 'Segoe UI, 10.75pt, style=Bold, Italic'
#
# label_date
#
$label_date.AutoSize = $true
$label_date.Location = New-Object System.Drawing.Point(80, 331)
$label_date.Name = 'label_date'
$label_date.Size = New-Object System.Drawing.Size(79, 13)
$label_date.TabIndex = 12
$label_date.Text = 'Date'
$label_date.ForeColor = "Yellow"
$label_date.Font = 'Segoe UI, 10.75pt, style=Bold, Italic'
#
# Bouton Excel pour l'ouverture d'un fixhier Excel à renseigner
#
$bouton_excelx.Location = New-Object System.Drawing.Point(80, 400)
$bouton_excelx.Size = New-Object System.Drawing.size(126, 25)
$bouton_excelx.Name = 'Excel'
$bouton_excelx.TabIndex = 14
$bouton_excelx.Text = 'Fichier Excel'
#Si le bouton excel est cliqué, alors active la fonction excelSpecial.
#Cree, ouvre et remplie une page Excel
$bouton_excelx.Add_Click(
	{
		excelSpecial
	}
)
#
# Bouton_recherchex
#
$bouton_recherchex.Location = New-Object System.Drawing.Point (80, 490)
$bouton_recherchex.Name = 'Boutton recherche'
$bouton_recherchex.Size = New-Object System.Drawing.Size(173 ,35)
$bouton_recherchex.TabIndex = 15
$bouton_recherchex.Text = 'Recherche'
$bouton_recherchex.UseVisualStyleBackColor = $true

$bouton_recherchex.Add_Click(
	{
		[String]$Station = $textBox_UCx.Get_Text()
		$loginfo = Get-WmiObject -Computer $Station -Class Win32_ComputerSystem
		$textBox_connectex.Text += $loginfo.UserName

		$textBox_marquex.Text = (Get-WmiObject -Class Win32_computerSystem -NameSpace "root\CIMV2" -Computer "$Station").Model
		$textBox_SNumberx.Text = (Get-WmiObject -Class Win32_BIOS -NameSpace "root\CIMV2" -Computer "$Station").SerialNumber
		$textBox_Memoryx.Text = (Get-WmiObject -Class Win32_computerSystem -NameSpace "root\CIMV2" -Computer "$Station").TotalPhysicalMemory
		$textBox_mpx.Text = (Get-WmiObject -Class Win32_Processor -Computer "$Station").Name
		$textBox_Datex.Text = (get-date).tostring('dd/MM/yyyy')

		$global:Number += 1; $ligne = 1 + "{0}" -f $global:Number.ToString()	###Le compteur � chaque clique###
		$c.Cells.Item($ligne,$col) = [String]$Station = $textBox_UCx.Text
		$col++
		$c.Cells.Item($ligne,$col) = [String]$User = $textBox_connectex.Text
		$col++
		$c.Cells.Item($ligne,$col) = [String]$Model = $textBox_marquex.Text
		$col++
		$c.Cells.Item($ligne,$col) = [String]$NSerie = $textBox_SNumberx.Text
		$col++
		$c.Cells.Item($ligne,$col) = [String]$Memory = $textBox_Memoryx.Text
		$col++
		$c.Cells.Item($ligne,$col) = [String]$MP = $textBox_mpx.Text
		$col++
		$c.Cells.Item($ligne,$col) = [String]$Date = $textBox_Datex.Text
		$col++
	}

)
#
# Bouton Effacer
#
$bouton_resetx.Location = New-Object System.Drawing.Point(353, 500)
$bouton_resetx.Name = 'Bouton Clear'
$bouton_resetx.Size = New-Object System.Drawing.Size(126,25)
$bouton_resetx.TabIndex = 16
$bouton_resetx.Text = 'Reinitialisation'
$bouton_resetx.UseVisualStyleBackColor = $true
$bouton_resetx.Add_Click(
	{
		$textBox_UCx.Clear()
		$textBox_connectex.Clear()
		$textBox_marquex.Clear()
		$textBox_SNumberx.Clear()
		$textBox_Memoryx.Clear()
		$textBox_mpx.Clear()
		$textBox_Datex.Clear()
	}
)