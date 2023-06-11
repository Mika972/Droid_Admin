####
# CreateKBLog.ps1 18.04.2018              OK            29/06/2018
####
$textBox_UC3 = New-Object System.Windows.Forms.TextBox
$button_valider =  New-Object System.Windows.Forms.Button
$button_clear = New-Object System.Windows.Forms.Button
$label_postesDeTaf = New-Object System.Windows.Forms.Label
#
# $tabpage3
#
$tabpage3.ClientSize = New-Object System.Drawing.Size(530, 145)
$tabpage3.Controls.Add($label_postesDeTaf)
$tabpage3.Controls.Add($textBox_UC3)
$tabpage3.Controls.Add($button_valider)
$tabpage3.Controls.Add($button_clear)
$tabpage3.Name = 'CreateKBLog'
$tabpage3.Text = 'KB Log'
#
# label sur les postes De travail
#
$label_postesDeTaf.AutoSize = $true
$label_postesDeTaf.Location = New-Object System.Drawing.Point(70, 31)
$label_postesDeTaf.Name = 'label1'
$label_postesDeTaf.Size = New-Object System.Drawing.Size(79, 13)
$label_postesDeTaf.TabIndex = 0
$label_postesDeTaf.Text = 'Poste de travail'
$label_postesDeTaf.ForeColor = "Yellow"
$label_postesDeTaf.Font = 'Segoe UI, 10.75pt, style=Bold, Italic' 
#
# textBox_UC3
#
$textBox_UC3.Location = New-Object System.Drawing.Point(211, 31)
$textBox_UC3.Multiline = $true
$textBox_UC3.Name = 'textBox_UC3'
$textBox_UC3.Size = New-Object System.Drawing.Size(270, 20)
$textBox_UC3.TabIndex = 1
#
# bouton valider
#
$button_valider.Location = New-Object System.Drawing.Point (70, 90)
$button_valider.Name = 'Bouton valider'
$button_valider.Size = New-Object System.Drawing.Size(173 ,35)
$button_valider.TabIndex = 2
$button_valider.Text = 'Creation log'
$button_valider.UseVisualStyleBackColor = $true
$button_valider.Add_Click(
	{
	   [String]$Station = $textBox_UC3.Get_Text()
	   function Get-HotFix2{
		     param($computername)
         Get-HotFix @PSBOundParameters |
         Select-Object description,hotfixid,installedby, @{l="InstalledOn";e={[DateTime]::Parse($_.psbase.properties["installedon"].value,$([System.Globalization.CultureInfo]::GetCultureInfo("en-US")))}}
	   }

	   get-Hotfix2 -computername $Station|out-file C:\Developpement\$Station"KBlog".txt
	}
)
#
# bouton effacer
#
$button_clear.Location = New-Object System.Drawing.Point(353, 100)
$button_clear.Name = 'Bouton Clear'
$button_clear.Size = New-Object System.Drawing.Size(126,25)
$button_clear.TabIndex = 3
$button_clear.Text = 'Reinitialisation'
$button_clear.UseVisualStyleBackColor = $true
$button_clear.Add_Click(
		{
		$textBox_UC3.Clear()
		}
)
