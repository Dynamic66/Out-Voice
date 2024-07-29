

function Main {

	Param ([String]$Commandline)
	
	Add-Type -AssemblyName System.speech -IgnoreWarnings
	Add-Type -AssemblyName "System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"

	Show-GUI_psf
	
	
	$script:ExitCode = 0
}

function Invoke-Speaker_ps1
{
	Param
	(
		[parameter(Position = 0)]
		[string]$say,
		[switch]$test,
		[switch]$help,
		[int]$rate,
		[int]$volume,
		$Voice
	)
	
	Add-Type -AssemblyName System.speech -IgnoreWarnings
	
	$script:speaker = New-Object System.Speech.Synthesis.SpeechSynthesizer
	
	$say = $say -replace "’" -replace "'"
	
	If ($rate)
	{
		$speaker.Rate = $rate
	}
	If ($volume)
	{
		$speaker.Volume = $volume
	}
	If ($voice)
	{
		$speaker.SelectVoice($voice)
	}
	
	
	
	If ($test)
	{
		"test" | Write-Output
		$script:speaker.Speak("Test Output")
	}
	ElseIf ($help)
	{
		$speaker | Get-Member
	}
	Else
	{
		$say | Write-Output
		$say | ForEach-Object {
			$script:speaker.Speak("$_")
		}
		
	}
	
}

function Show-GUI_psf
{

	#region Import the Assemblies
	[void][reflection.assembly]::Load('System.Drawing, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
	[void][reflection.assembly]::Load('System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
	[void][reflection.assembly]::Load('System.Design, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
	#endregion Import Assemblies

	#----------------------------------------------
	#region Generated Form Objects
	#----------------------------------------------
	[System.Windows.Forms.Application]::EnableVisualStyles()
	$fOutVoice = New-Object 'System.Windows.Forms.Form'
	$checkboxActive = New-Object 'System.Windows.Forms.CheckBox'
	$cVoice = New-Object 'System.Windows.Forms.ComboBox'
	$tbRate = New-Object 'System.Windows.Forms.TrackBar'
	$tbVolume = New-Object 'System.Windows.Forms.TrackBar'
	$labelRate = New-Object 'System.Windows.Forms.Label'
	$labelVolume = New-Object 'System.Windows.Forms.Label'
	$timer1 = New-Object 'System.Windows.Forms.Timer'
	$trayicon = New-Object 'System.Windows.Forms.NotifyIcon'
	$contextmenustrip1 = New-Object 'System.Windows.Forms.ContextMenuStrip'
	$toolstripmenuitem1 = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState'
	#endregion Generated Form Objects

	#----------------------------------------------
	# User Generated Script
	#----------------------------------------------

	Function Invoke-Speach
	{
		$ErrorActionPreference = 'SilentlyContinue'
		Stop-Process $current  -Verbose -ErrorAction SilentlyContinue
		Write-Verbose -Message ($new) -Verbose -ErrorAction SilentlyContinue
		#$current = Start-Process -FilePath "powershell.exe" -ArgumentList "-Command Invoke-Speaker_ps1 -say 'test' -rate 10 -volume 10" -PassThru
		Invoke-Speaker_ps1 -say $new -rate $($tbRate.Value - 10) -volume $($tbVolume.Value * 10) -voice $($cVoice.Text)
			

		
	}
	
	$fOutVoice_Load = {
		Add-Type -AssemblyName System.speech 
		[string]$script:processed = ((Get-Clipboard -raw) -replace "'" -replace ('/[^a-zA-Z0-9]/g')).trim()
		$script:speaker = New-Object System.Speech.Synthesis.SpeechSynthesizer
		($speaker.GetInstalledVoices().VoiceInfo).name | ForEach-Object {
			$cVoice.Items.Add($_)
		}
	}
	
	$fOutVoice_Shown = {
		$cVoice.Text = $speaker.Voice.name
		$timer1.Enabled = $true
		$labelVolume.Text = "Volume: $($tbVolume.Value * 10)%"
		$labelRate.Text = "Rate: $($tbRate.Value - 10)"
	}
	
	
	$timer1_Tick = {
		[string]$new = ((Get-Clipboard -raw) -replace "'" -replace ('/[^a-zA-Z0-9]/g')).trim()
		If (($new -ne $processed) -and ($null -ne $new))
		{
			$fOutVoice.TopMost = $true
			$fOutVoice.WindowState = 'Normal'
			Invoke-Speach
			$script:processed = $new
			Write-Verbose -Message "Awaiting diferent String..." -Verbose -ErrorAction SilentlyContinue
			$fOutVoice.TopMost = $false
		}
	}
	
	
	$cVoice_SelectedIndexChanged = {
		$cVoice.Text | Out-Host
	}
	
	$tbVolume_Scroll = {
		$labelVolume.Text = "Volume: $($tbVolume.Value * 10)%"
	}
	
	$tbRate_Scroll = {
		$labelRate.Text = "Rate: $($tbRate.Value - 10)"
	}
	
	$trayIcon_MouseDoubleClick = [System.Windows.Forms.MouseEventHandler]{
		Stop-Process $current -Force -Verbose -ErrorAction SilentlyContinue
	}
	
	$fOutVoice_FormClosing = [System.Windows.Forms.FormClosingEventHandler]{
		Stop-Process $current -Force -Verbose -ErrorAction SilentlyContinue
		$trayIcon.Dispose()
	}
	
	$fOutVoice_Activated={
		Stop-Process $current -Force -Verbose -ErrorAction SilentlyContinue
	}
	
	$fOutVoice_Click={
		Stop-Process $current -Force -Verbose -ErrorAction SilentlyContinue
	}
	
	$checkboxActive_CheckedChanged={
		If ($checkboxActive.Checked)
		{
			$timer1.Enabled = $true
		}
		Else
		{
			$timer1.Enabled = $false
		}
	}
	
	
	$labelRate_Click={
		Stop-Process $current -Force -Verbose -ErrorAction SilentlyContinue
	}
	
	$labelVolume_Click={
		Stop-Process $current -Force -Verbose -ErrorAction SilentlyContinue
	}
	
	# --End User Generated Script--
	#----------------------------------------------
	#region Generated Events
	#----------------------------------------------
	
	$Form_StateCorrection_Load=
	{
		#Correct the initial state of the form to prevent the .Net maximized form issue
		$fOutVoice.WindowState = $InitialFormWindowState
	}
	
	$Form_StoreValues_Closing=
	{
		#Store the control values
		$script:GUI_checkboxActive = $checkboxActive.Checked
		$script:GUI_cVoice = $cVoice.Text
		$script:GUI_cVoice_SelectedItem = $cVoice.SelectedItem
		$script:GUI_tbRate = $tbRate.Value
		$script:GUI_tbVolume = $tbVolume.Value
	}

	
	$Form_Cleanup_FormClosed=
	{
		#Remove all event handlers from the controls
		try
		{
			$checkboxActive.remove_CheckedChanged($checkboxActive_CheckedChanged)
			$cVoice.remove_SelectedIndexChanged($cVoice_SelectedIndexChanged)
			$tbRate.remove_Scroll($tbRate_Scroll)
			$tbVolume.remove_Scroll($tbVolume_Scroll)
			$labelRate.remove_Click($labelRate_Click)
			$labelVolume.remove_Click($labelVolume_Click)
			$fOutVoice.remove_Activated($fOutVoice_Activated)
			$fOutVoice.remove_FormClosing($fOutVoice_FormClosing)
			$fOutVoice.remove_Load($fOutVoice_Load)
			$fOutVoice.remove_Shown($fOutVoice_Shown)
			$fOutVoice.remove_Click($fOutVoice_Click)
			$timer1.remove_Tick($timer1_Tick)
			$trayicon.remove_MouseDoubleClick($trayicon_MouseDoubleClick)
			$fOutVoice.remove_Load($Form_StateCorrection_Load)
			$fOutVoice.remove_Closing($Form_StoreValues_Closing)
			$fOutVoice.remove_FormClosed($Form_Cleanup_FormClosed)
		}
		catch { Out-Null <# Prevent PSScriptAnalyzer warning #> }
	}
	#endregion Generated Events

	#----------------------------------------------
	#region Generated Form Code
	#----------------------------------------------
	$fOutVoice.SuspendLayout()
	$tbRate.BeginInit()
	$tbVolume.BeginInit()
	$contextmenustrip1.SuspendLayout()
	#
	# fOutVoice
	#
	$fOutVoice.Controls.Add($checkboxActive)
	$fOutVoice.Controls.Add($cVoice)
	$fOutVoice.Controls.Add($tbRate)
	$fOutVoice.Controls.Add($tbVolume)
	$fOutVoice.Controls.Add($labelRate)
	$fOutVoice.Controls.Add($labelVolume)
	$fOutVoice.AutoScaleDimensions = New-Object System.Drawing.SizeF(6, 13)
	$fOutVoice.AutoScaleMode = 'Font'
	$fOutVoice.AutoSize = $True
	$fOutVoice.AutoSizeMode = 'GrowAndShrink'
	$fOutVoice.BackColor = [System.Drawing.Color]::FromArgb(255, 32, 32, 32)
	$fOutVoice.ClientSize = New-Object System.Drawing.Size(138, 201)
	$fOutVoice.MaximizeBox = $False
	$fOutVoice.MinimizeBox = $False
	$fOutVoice.Name = 'fOutVoice'
	$fOutVoice.ShowIcon = $False
	$fOutVoice.SizeGripStyle = 'Hide'
	$fOutVoice.StartPosition = 'CenterScreen'
	$fOutVoice.TopMost = $True
	$fOutVoice.add_Activated($fOutVoice_Activated)
	$fOutVoice.add_FormClosing($fOutVoice_FormClosing)
	$fOutVoice.add_Load($fOutVoice_Load)
	$fOutVoice.add_Shown($fOutVoice_Shown)
	$fOutVoice.add_Click($fOutVoice_Click)
	#
	# checkboxActive
	#
	$checkboxActive.BackgroundImageLayout = 'Stretch'
	$checkboxActive.Checked = $True
	$checkboxActive.CheckState = 'Checked'
	$checkboxActive.FlatAppearance.BorderColor = [System.Drawing.Color]::Red 
	$checkboxActive.FlatAppearance.CheckedBackColor = [System.Drawing.Color]::Red 
	$checkboxActive.FlatStyle = 'Flat'
	$checkboxActive.ForeColor = [System.Drawing.Color]::White 
	$checkboxActive.Location = New-Object System.Drawing.Point(6, 39)
	$checkboxActive.Name = 'checkboxActive'
	$checkboxActive.Size = New-Object System.Drawing.Size(75, 24)
	$checkboxActive.TabIndex = 16
	$checkboxActive.Text = 'Active'
	$checkboxActive.UseVisualStyleBackColor = $False
	$checkboxActive.add_CheckedChanged($checkboxActive_CheckedChanged)
	#
	# cVoice
	#
	$cVoice.BackColor = [System.Drawing.Color]::FromArgb(255, 32, 32, 32)
	$cVoice.DropDownStyle = 'DropDownList'
	$cVoice.FlatStyle = 'Popup'
	$cVoice.ForeColor = [System.Drawing.Color]::White 
	$cVoice.FormattingEnabled = $True
	$cVoice.Location = New-Object System.Drawing.Point(6, 12)
	$cVoice.Name = 'cVoice'
	$cVoice.Size = New-Object System.Drawing.Size(125, 21)
	$cVoice.TabIndex = 15
	$cVoice.add_SelectedIndexChanged($cVoice_SelectedIndexChanged)
	#
	# tbRate
	#
	$tbRate.BackColor = [System.Drawing.Color]::FromArgb(255, 32, 32, 32)
	$tbRate.LargeChange = 1
	$tbRate.Location = New-Object System.Drawing.Point(86, 82)
	$tbRate.Maximum = 20
	$tbRate.Name = 'tbRate'
	$tbRate.Orientation = 'Vertical'
	$tbRate.Size = New-Object System.Drawing.Size(45, 115)
	$tbRate.TabIndex = 11
	$tbRate.TickStyle = 'Both'
	$tbRate.Value = 11
	$tbRate.add_Scroll($tbRate_Scroll)
	#
	# tbVolume
	#
	$tbVolume.BackColor = [System.Drawing.Color]::FromArgb(255, 32, 32, 32)
	$tbVolume.LargeChange = 1
	$tbVolume.Location = New-Object System.Drawing.Point(6, 82)
	$tbVolume.Name = 'tbVolume'
	$tbVolume.Orientation = 'Vertical'
	$tbVolume.RightToLeft = 'No'
	$tbVolume.Size = New-Object System.Drawing.Size(45, 115)
	$tbVolume.TabIndex = 10
	$tbVolume.TickStyle = 'Both'
	$tbVolume.Value = 8
	$tbVolume.add_Scroll($tbVolume_Scroll)
	#
	# labelRate
	#
	$labelRate.AutoSize = $True
	$labelRate.ForeColor = [System.Drawing.SystemColors]::ButtonFace 
	$labelRate.Location = New-Object System.Drawing.Point(86, 66)
	$labelRate.Name = 'labelRate'
	$labelRate.Size = New-Object System.Drawing.Size(30, 13)
	$labelRate.TabIndex = 13
	$labelRate.Text = 'Rate'
	$labelRate.add_Click($labelRate_Click)
	#
	# labelVolume
	#
	$labelVolume.AutoSize = $True
	$labelVolume.BackColor = [System.Drawing.Color]::FromArgb(255, 32, 32, 32)
	$labelVolume.ForeColor = [System.Drawing.SystemColors]::ButtonFace 
	$labelVolume.Location = New-Object System.Drawing.Point(6, 66)
	$labelVolume.Name = 'labelVolume'
	$labelVolume.Size = New-Object System.Drawing.Size(42, 13)
	$labelVolume.TabIndex = 12
	$labelVolume.Text = 'Volume'
	$labelVolume.add_Click($labelVolume_Click)
	#
	# timer1
	#
	$timer1.add_Tick($timer1_Tick)
	#
	# trayicon
	#
	$trayicon.ContextMenuStrip = $contextmenustrip1

	$trayicon.Icon = [System.Drawing.SystemIcons]::Hand

	$trayicon.Text = 'Out-Voice'
	$trayicon.Visible = $True
	$trayicon.add_MouseDoubleClick($trayicon_MouseDoubleClick)
	#
	# contextmenustrip1
	#
	[void]$contextmenustrip1.Items.Add($toolstripmenuitem1)
	$contextmenustrip1.Name = 'contextmenustrip1'
	$contextmenustrip1.Size = New-Object System.Drawing.Size(135, 26)
	#
	# toolstripmenuitem1
	#
	

	#endregion
	$toolstripmenuitem1.Image = [System.Drawing.SystemIcons]::Hand
	$toolstripmenuitem1.Name = 'toolstripmenuitem1'
	$toolstripmenuitem1.Size = New-Object System.Drawing.Size(134, 22)
	$toolstripmenuitem1.Text = 'Add-Voices'
	$contextmenustrip1.ResumeLayout()
	$tbVolume.EndInit()
	$tbRate.EndInit()
	$fOutVoice.ResumeLayout()
	#endregion Generated Form Code

	#----------------------------------------------

	#Save the initial state of the form
	$InitialFormWindowState = $fOutVoice.WindowState
	#Init the OnLoad event to correct the initial state of the form
	$fOutVoice.add_Load($Form_StateCorrection_Load)
	#Clean up the control events
	$fOutVoice.add_FormClosed($Form_Cleanup_FormClosed)
	#Store the control values when form is closing
	$fOutVoice.add_Closing($Form_StoreValues_Closing)
	#Show the Form
	return $fOutVoice.ShowDialog()

}
#endregion Source: GUI.psf

#Start the application
Main ($CommandLine)
