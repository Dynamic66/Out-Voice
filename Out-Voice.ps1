
try {
    #region load dependencies

    #OPTIONAL
    $ErrorActionPreference = 'Continue' 
    Add-Type -AssemblyName system.Drawing
    Add-Type -AssemblyName System.Design

    function Set-ControlProerty {
        param (
            [parameter(mandatory = $true)]
            [array]$Controls,
            [System.Drawing.Color]$foreColor,
            [System.Drawing.Color]$backColor,
            [System.Windows.Forms.Padding]$padding,
            [System.Windows.Forms.FlatStyle]$flatStyle
        )
        try {
            $ErrorActionPreference = 'Continue'
            foreach ($control in $Controls) {
                if ($foreColor) {
                    $control.ForeColor = $foreColor
                }
                if ($backColor) {
                    $control.BackColor = $backColor
                }
                if ($padding) {
                    $control.Padding = $padding
                }
                if ($null -ne $flatStyle) {
                    # just if($flatStyle) did not work here
                    $control.FlatStyle = $flatStyle
                }
            }
        }
        catch {
            Write-Warning "$($control.GetType()) caused a error:`n$_"
        }
    }

    #DEPENDENT
    $ErrorActionPreference = 'stop'
    
    function Invoke-Speech {
        Param
        (
            [string]$say,
            [switch]$test,
            [switch]$help,
            [switch]$passthrou,
            [int]$rate = 0,
            [int]$volume = 50,
            [string]$Voice
        )

        If ($rate) {
            $sSpeech.Rate = $rate
        }
        If ($volume) {
            $sSpeech.Volume = $volume
        }
        If ($voice) {
            $sSpeech.SelectVoice($voice)
        }
        If ($test) {
            Write-Verbose 'test' -Verbose
            $sSpeech.Speak('Test Output')
        }
        ElseIf ($help) {
            $sSpeech | Get-Member
        }
        Elseif ($say) {
            
            $sSpeech.SpeakAsyncCancelAll()

            if ($passthrou) {
                Write-Output $sSpeech
            }

            $sSpeech.SpeakAsync($say)
        }
    }

    Add-Type -AssemblyName system.windows.forms
    Add-Type -AssemblyName System.speech
    #endregion
}
catch {
    [System.Windows.Forms.MessageBox]::show("$_ Check if the AssemblyName is correct.", 'Error While loading dependencies.', 'ok', 'error')
}

try {
    $ErrorActionPreference = 'stop'
    #region setup
    [System.Windows.Forms.Application]::EnableVisualStyles()

    $form = New-Object System.Windows.Forms.Form
    $bPlayControl = New-Object System.Windows.Forms.Button
    $cVocieSelcetion = New-Object System.Windows.Forms.ComboBox
    $cAutoCopy = New-Object System.Windows.Forms.checkbox
    $rText = New-Object System.Windows.Forms.richTextBox
    $pFill = New-Object System.Windows.forms.panel
    $tVolume = New-Object System.Windows.Forms.TrackBar
    $tUpdate = New-Object System.Windows.Forms.timer
    $pTop = New-Object System.Windows.forms.panel
    $bStop = New-Object System.Windows.Forms.Button
    $cTopmost = New-Object System.Windows.Forms.checkbox
    $pBottom = New-Object System.Windows.forms.panel
    $nRate = New-Object System.Windows.Forms.NumericUpDown
    $cAutoPlay = New-Object System.Windows.Forms.checkbox
    $script:sSpeech = New-Object System.Speech.Synthesis.SpeechSynthesizer
    $script:state = $null

    $theme = @{
        Primary   = [system.drawing.color]::FromArgb(80, 80, 80)
        Secondary = [system.drawing.color]::FromArgb(90, 90, 90)
        ForeColor = [system.drawing.color]::WhiteSmoke
        Padding   = [System.Windows.Forms.Padding]::new(6, 6, 6, 6)
    }
    #endregion setup
}
catch {
    [System.Windows.Forms.MessageBox]::show("$_ Check if dependent Assemblys are loaded.", 'Error while creating controls and variables.', 'ok', 'error')
}

try {
    $ErrorActionPreference = 'Continue'
    #region Controls config
    
    #region Top
    $pTop.Dock = 'top'
    $pTop.size = '100,35'

    $cAutoPlay.dock = 'left'
    $cAutoPlay.Text = 'AutoPlay'
    $cAutoPlay.Appearance = 'button'
    $cAutoPlay.AutoSize = $true
    $cAutoPlay.TextAlign = 'MiddleCenter'
    $cAutoPlay.Checked = $true

    $cTopmost.TextAlign = 'MiddleCenter'
    $cTopmost.Appearance = 'button'
    $cTopmost.text = 'AlwaysOnTop'
    $cTopmost.dock = 'left'
    $cTopmost.Checked = $true
    $cTopmost.Autosize = $true
    $cTopmost.add_click({
            try {
                if ($args[0].checked) {
                    $form.topmost = $true
                }
                else {
                    $form.topmost = $false
                }
            }
            catch {
                [System.Windows.Forms.MessageBox]::show("$_ Check if dependent Assemblys are loaded.", 'Error while clicking the topmost checkbox.', 'ok', 'error')
            }
        })

    $cAutoCopy.text = 'AutoCopy'
    $cAutoCopy.dock = 'left'
    $cAutoCopy.TextAlign = 'MiddleCenter'
    $cAutoCopy.Appearance = 'button'
    $cAutoCopy.AutoSize = $true
    $cAutoCopy.Checked = $true
    $cAutoCopy.add_click({
            try {
                if ($cAutoCopy.Checked) {
                    $tUpdate.Enabled = $true
                    $cAutoPlay.Enabled = $true
                }
                else {
                    $tUpdate.Enabled = $false
                    $cAutoPlay.Enabled = $false
                }
            }
            catch {
                [System.Windows.Forms.MessageBox]::show("$_ Check if dependent Assemblys are loaded.", 'Error while clicking the AutoCopy checkbox.', 'ok', 'error')
            }
        })

    #endregion

    #region bottom
    $pBottom.dock = 'bottom'
    $pBottom.size = '100,35'

    $tVolume.Orientation = 'Horizontal'
    $tVolume.TickStyle = 'TopLeft'
    $tVolume.dock = 'right'
    $tVolume.Maximum = 10
    $tVolume.Value = 8
    $tVolume.AutoSize = $true

    $tVolume.Orientation = 'Horizontal'
    $tVolume.TickStyle = 'TopLeft'
    $tVolume.dock = 'right'
    $tVolume.Maximum = 10
    $tVolume.Value = 8
    $tVolume.AutoSize = $true

    $nRate.Maximum = 10
    $nRate.Minimum = -10
    $nRate.Value = 1
    $nRate.dock = 'right'
    $nRate.autosize = $true

    $bStop.dock = 'left'
    $bStop.AutoSize = $true
    $bStop.AutoSizeMode = 'GrowAndShrink'
    $bStop.Text = '⏹'
    $bStop.add_click({
            try {
                $script:sSpeech.SpeakAsyncCancelAll()
                $script:state = $null
            }
            catch {
                [System.Windows.Forms.MessageBox]::show("$_ Check if dependent Assemblys are loaded.", 'Error while clicking the stop button.', 'ok', 'error')
            }
        })

    $bPlayControl.dock = 'left'
    $bPlayControl.text = '⏯'
    $bPlayControl.Font = 'Segoe UI,12'
    $bPlayControl.add_click({
            try {
                if ($rText.SelectedText) {
                    $script:state = Invoke-Speech -say $rText.SelectedText -volume ($tVolume.Value * 10) -Voice $cVocieSelcetion.text -rate $nRate.Value -passthrou
                }
                else {
                    switch ($script:state.state) {
                        'Ready' {
                            $script:state = Invoke-Speech -say $rText.Text -volume ($tVolume.Value * 10) -Voice $cVocieSelcetion.text -rate $nRate.Value -passthrou
                        }
                        'Speaking' {
                            Write-Verbose 'pause'
                            $sSpeech.Pause()
                        }
                        'Paused' {
                            Write-Verbose 'resume'
                            $sSpeech.Resume()
                        }
                        default {
                            $script:state = Invoke-Speech -say $rText.Text -volume ($tVolume.Value * 10) -Voice $cVocieSelcetion.text -rate $nRate.Value -passthrou
                        }
                    }
                }
            }
            catch {
                [System.Windows.Forms.MessageBox]::show("$_ Check if dependent Assemblys are loaded.",  'Error while clicking the play/pause button.', 'ok', 'error')
            }
        })
    #endregion Controls config


    #region Fill
    $pFill.dock = 'fill'

    $cVocieSelcetion.dock = 'bottom'

    $rText.dock = 'fill'
    $rText.BorderStyle = 'none'
    #endregion

    $tUpdate.Interval = 350
    $tUpdate.add_tick({
            try {
                $ErrorActionPreference = 'Stop'
                #Visual Feedback for playcontrols
                switch ($script:state.state) {
                    'Ready' { $bPlayControl.BackColor = $theme.Secondary }
                    'Speaking' { $bPlayControl.BackColor = [system.drawing.color]::darkgreen }
                    'Paused' { $bPlayControl.BackColor = [system.drawing.color]::Orange }
                    default { $bPlayControl.BackColor = $theme.Secondary }
                }

                #logic for reading the clipboard content
                $currentClipboard = [string](Get-Clipboard)

                if ((-not [string]::IsNullOrWhiteSpace($currentClipboard)) -and ($currentClipboard -ne $rText.text)) {

                    $rText.text = $currentClipboard
                    $rText.SelectionStart = $rText.Text.Length #set the curser at the end of text

                    if ($cAutoPlay.checked) {
                        $script:state = Invoke-Speech -say $rText.Text -volume ($tVolume.Value * 10) -Voice $cVocieSelcetion.text -rate $nRate.Value -passthrou
                    }
                }
                $currentClipboard = $null
            }
            catch { 
                $tUpdate.stop()
                [System.Windows.Forms.MessageBox]::show("$_ Check if something that was Expected was not delivered.", 'Looping error detected. Exiting script.', 'ok', 'error')
                $form.Close()
            }
        })

    $form.text = 'Out-Voice'
    $form.ShowIcon = $false
    $form.ShowInTaskbar = $true
    $form.add_load({
            try {
                Set-ControlProerty -Controls $form -backColor $theme.Primary
                Set-ControlProerty -Controls @($pTop.Controls + $pFill.Controls + $pBottom.Controls + $form.Controls) -foreColor $theme.ForeColor -backColor $theme.Secondary

                Set-ControlProerty -Controls $pTop, $pFill, $pBottom, $form -padding $theme.Padding

                Set-ControlProerty -Controls $cAutoCopy, $cTopmost, $cAutoPlay, $bPlayControl, $bStop, $cVocieSelcetion -flatStyle ([System.Windows.Forms.FlatStyle]::Flat)
                $cVocieSelcetion.Items.AddRange($sSpeech.GetInstalledVoices().voiceinfo.name)
                $cVocieSelcetion.Text = $cVocieSelcetion.Items[0]
                $tUpdate.Start()
                $form.TopMost = $true
            }
            catch {
                [System.Windows.Forms.MessageBox]::show("$_ Check if dependent Assemblys are loaded.", 'Error while loading form', 'ok', 'error')
            }
        })
    $form.add_closing({
            #cleanup
            $ErrorActionPreference = 'SilentlyContinue'
            $tUpdate.stop()
            $tUpdate.Dispose()
            $script:sSpeech.SpeakAsyncCancelAll()
            $script:sSpeech.Dispose()
            $script:sSpeech = $null
            $script:state = $null
        })

    #region add controls
    $pTop.Controls.AddRange(@(
            $cTopmost,
            $cAutoPlay,
            $cAutoCopy
        ))
    $pFill.Controls.AddRange(@(
            $rText,
            $cVocieSelcetion
        ))
    $pBottom.Controls.AddRange(@(
            $bStop,
            $bPlayControl,
            $tVolume,
            $nRate
        ))
    $form.Controls.AddRange(@(
            $pFill,
            $pBottom,
            $pTop
        ))
    #endregion
    
    #endregion
    $form.showdialog()
}
catch {
    [System.Windows.Forms.MessageBox]::show("$_. Check if dependent Assemblys are loaded.", 'Error while setting Corntol propertys.', 'ok', 'error')
}
finally {
    #cleanup in case form building 
    $ErrorActionPreference = 'SilentlyContinue'
    $tUpdate.stop()
    $tUpdate.Dispose()
    $script:sSpeech.SpeakAsyncCancelAll()
    $script:sSpeech.Dispose()
    $script:sSpeech.Dispose() 
    $script:sSpeech = $null
    $script:state = $null
    $form.close()
}
