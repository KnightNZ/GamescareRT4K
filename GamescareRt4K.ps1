 <#
        .SYNOPSIS
        Switch a GamesCare SCART Switch and RetroTink4K scaler simultaneously.

        .DESCRIPTION
        A GamesCare SCART Switch on the local network can be accessed via http://gscartsw.local (by default) - this script will allow you to select the input 
        ports directly instead of using the WebUI, as part of doing so, the relevant "Profile" button will be simulated for an attached RetroTink4K scaler 
        device (https://www.retrotink.com/product-page/retrotink-4k).  Some additional buttons are also added for direct control of a few Rt4K functions.

        .LINK
        Using ideas from Carter300: https://github.com/carter300/RetroTink4K-PCRemote/tree/main
    #>

    function Send-SerialCommand {
        param (
            [string]$Command # Command to send via serial port
         )
    
        # Define the COM port and settings
        $portName               = "COM3"    # Replace with your COM port (e.g., COM1, COM2, etc.)
        $baudRate               = 115200    # The Rt4K natively uses 115200
        $serialPort             = New-Object System.IO.Ports.SerialPort
        $serialPort.PortName    = $portName
        $serialPort.BaudRate    = $baudRate
    
    try {
        $serialPort.Open()
       # Write-Host "Serial port $portName opened successfully." # DEBUG
    
        if ($serialPort.IsOpen) {
                $serialPort.WriteLine($command)
                Write-Host "Command sent: $command"
            } else {
                Write-Host "Error: Serial port is not open."
            }
         Start-Sleep -Milliseconds 50 # Small wait between commands
        } catch {
            Write-Host "Error: $_"
        } finally {
        # Close the serial port
            if ($serialPort.IsOpen) {
                $serialPort.Close()
                #Write-Host "Serial port $portName closed." # DEBUG
            }
        }
    }
    
    
    Function SendRt4K {
        param (
            [Parameter(Mandatory=$true)]
            [string[]]$Commands, # Can chain multiple commands to process in order, primarily to auto-select "OK" on a mode change.
            [Parameter(Mandatory=$false)]
            [string]$Text
         )
    
            If(!([string]::IsNullOrEmpty($Text)))
                {
                    Write-Host $Text -ForegroundColor Green
                }
    
        ForEach ($Command in $Commands) {
    
            # Debug output for the argument
            Write-Host "Passed argument for input: '$Command'" -ForegroundColor Yellow
    
            # Check if the input argument was provided
            if ([string]::IsNullOrEmpty($Command)) {
                Write-Host "Error: No hex string provided. Please run with -Input 'HexString'." -ForegroundColor Red
            #  exit 1
            }
    
            # List of valid commands
            $validCommands = @(
                "pwr", "menu", "up", "down", "left", "right", "ok", "back", "diag", "stat","input", "output", "scaler", "sfx", "adc", "prof", "prof1", "prof2", "prof3",
                "prof4", "prof5", "prof6", "prof7", "prof8", "prof9", "prof10", "prof11", "prof12","gain", "phase", "pause", "safe", "genlock", "buffer", "res4k", "res1080p", "res1440p",
                "res480p", "res1", "res2", "res3", "res4", "aux1", "aux2", "aux3", "aux4", "aux5","aux6", "aux7", "aux8", "pwr on"
            )
    
            # Check if the input is a valid command
            if ($validCommands -contains $Command) {
                # Prepend "remote" to the command
                $Command = "remote $Command`n"
            } else {
                Write-Host "Unknown command: '$Command'. Please enter a valid command." -ForegroundColor Red
                exit 1
            }
    
            # Debug output for the command
            Write-Host "Command to be executed: $Command" -ForegroundColor Cyan
        
            # Call the function to send the serial command
            Send-SerialCommand -Command $Command
        }
    }
    
    Function Switch-Input {
        Param (
            [string]$URL,
            [string]$Command
        )
        Write-Host "URL Query: ""$URL"""
        If ($Command) {SendRt4k -Commands $Command}
        $result     = Invoke-RestMethod -Uri $URL -Method Get
        $Script:portsel    = $result.active
        $label.Text = "Current Port: $portsel"
        Start-Sleep -Milliseconds 50
    }
    
    
    Function Select-Input {
        # Create the form
        $form               = New-Object System.Windows.Forms.Form
        $form.Text          = "GamesCare Switch - RetroTink 4K Profile Selector"
        $form.Size          = New-Object System.Drawing.Size(470, 450)
        $form.StartPosition = "CenterScreen"
    
    
        # Create a label to display the current port
        $label                  = New-Object System.Windows.Forms.Label
        $label.Text             = "Current Port:"
        $label.Location         = New-Object System.Drawing.Point(300, 12)
        $label.Size             = New-Object System.Drawing.Size(100, 20)
        $form.Controls.Add($label)
        $label.BackColor        = [System.Drawing.Color]::FromName("Transparent")
        $label.ForeColor        = [System.Drawing.Color]::Black
    
    
        #GamesCare Button 1
        $GCbutton1              = New-Object System.Windows.Forms.Button
        $GCbutton1.Size         = New-Object System.Drawing.Size(200, 30)
        $GCbutton1.Location     = New-Object System.Drawing.Point(120, 60)
        $GCButton1.Text         = "Port 1"
        $GCButton1.Add_Click({Switch-Input -URL "http://$GamesCareIP/ports?force=1" -Command "prof1"})
        $form.Controls.Add($GCbutton1)
    
    
            #GamesCare Button 2
        $GCbutton2              = New-Object System.Windows.Forms.Button
        $GCbutton2.Size         = New-Object System.Drawing.Size(200, 30)
        $GCbutton2.Location     = New-Object System.Drawing.Point(120, 90)
        $GCButton2.Text         = "Port 2"
        $GCButton2.Add_Click({Switch-Input -URL "http://$GamesCareIP/ports?force=2" -Command "prof2"})
        $form.Controls.Add($GCbutton2)
    
    
            #GamesCare Button 3
        $GCbutton3              = New-Object System.Windows.Forms.Button
        $GCbutton3.Size         = New-Object System.Drawing.Size(200, 30)
        $GCbutton3.Location     = New-Object System.Drawing.Point(120, 120)
        $GCButton3.Text         = "Port 3"
        $GCButton3.Add_Click({Switch-Input -URL "http://$GamesCareIP/ports?force=3" -Command "prof3"})
        $form.Controls.Add($GCbutton3)
    
    
            #GamesCare Button 4
        $GCbutton4              = New-Object System.Windows.Forms.Button
        $GCbutton4.Size         = New-Object System.Drawing.Size(200, 30)
        $GCbutton4.Location     = New-Object System.Drawing.Point(120, 150)
        $GCButton4.Text         = "Port 4"
        $GCButton4.Add_Click({Switch-Input -URL "http://$GamesCareIP/ports?force=4" -Command "prof4"})
        $form.Controls.Add($GCbutton4)
    
    
            #GamesCare Button 5
        $GCbutton5              = New-Object System.Windows.Forms.Button
        $GCbutton5.Size         = New-Object System.Drawing.Size(200, 30)
        $GCbutton5.Location     = New-Object System.Drawing.Point(120, 180)
        $GCButton5.Text         = "Port 5"
        $GCButton5.Add_Click({Switch-Input -URL "http://$GamesCareIP/ports?force=5" -Command "prof5"})
        $form.Controls.Add($GCbutton5)
    
    
            #GamesCare Button 6
        $GCbutton6              = New-Object System.Windows.Forms.Button
        $GCbutton6.Size         = New-Object System.Drawing.Size(200, 30)
        $GCbutton6.Location     = New-Object System.Drawing.Point(120, 210)
        $GCButton6.Text         = "Port 6"
        $GCButton6.Add_Click({Switch-Input -URL "http://$GamesCareIP/ports?force=6" -Command "prof6"})
        $form.Controls.Add($GCbutton6)
    
    
            #GamesCare Button 7
        $GCbutton7              = New-Object System.Windows.Forms.Button
        $GCbutton7.Size         = New-Object System.Drawing.Size(200, 30)
        $GCbutton7.Location     = New-Object System.Drawing.Point(120, 240)
        $GCButton7.Text         = "Port 7"
        $GCButton7.Add_Click({Switch-Input -URL "http://$GamesCareIP/ports?force=7" -Command "prof7"})
        $form.Controls.Add($GCbutton7)
    
    
            #GamesCare Button 8
        $GCbutton8              = New-Object System.Windows.Forms.Button
        $GCbutton8.Size         = New-Object System.Drawing.Size(200, 30)
        $GCbutton8.Location     = New-Object System.Drawing.Point(120, 270)
        $GCButton8.Text = "Port 8"
        $GCButton8.Add_Click({Switch-Input -URL "http://$GamesCareIP/ports?force=8" -Command "prof8"})
        $form.Controls.Add($GCbutton8)
    
    
            #GamesCare Button Auto
        $GCbutton0              = New-Object System.Windows.Forms.Button
        $GCbutton0.Size         = New-Object System.Drawing.Size(200, 30)
        $GCbutton0.Location     = New-Object System.Drawing.Point(120, 300)
        $GCButton0.Text         = "Auto"
        $GCButton0.Add_Click({Switch-Input -URL "http://$GamesCareIP/ports?force=0" -Command "ok"})
        $form.Controls.Add($GCbutton0)
    
    
              #GamesCare Button Reset
    #   $GCbuttonR              = New-Object System.Windows.Forms.Button
    #   $GCbuttonR.Size         = New-Object System.Drawing.Size(120, 30)
    #    $GCbuttonR.Location     = New-Object System.Drawing.Point(320, 150)
    #    $GCButtonR.Text         = "Reboot"
    #    $GCButtonR.Add_Click({
    #    Write-Host "Rebooting Switch"
    #            $result         = Invoke-RestMethod -Uri "http://$GamesCareIP/settings?reboot=1" -Method Get -ErrorAction Ignore # Will timeout, to be expected.
    #            $portsel        = $result.active
    #            $label.Text     = "Current Port: $portsel"
    #            })
      #  $form.Controls.Add($GCbuttonR) - Don't add, behaves strangely.
    
    
        # Create a label for Rt4K settings
        $label2                 = New-Object System.Windows.Forms.Label
        $label2.Text            = "RetroTink 4K:"
        $label2.Location        = New-Object System.Drawing.Point(20, 350)
        $label2.Size            = New-Object System.Drawing.Size(100, 20)
        $form.Controls.Add($label2)
        $label2.BackColor       = [System.Drawing.Color]::FromName("Transparent")
        $label2.ForeColor       = [System.Drawing.Color]::Black
    
    
        # 4:3 AR Button
        $RtButton1              = New-Object System.Windows.Forms.Button
        $RtButton1.Text         = "4:3"
        $RtButton1.Size         = New-Object System.Drawing.Size(60,30)
        $RtButton1.Location     = New-Object System.Drawing.Point(20,370)
        $RtButton1.Add_Click({SendRt4k -Commands "aux2" -Text "Aspect Ratio: 4:3"})
        $form.Controls.Add($RtButton1)
        
    
        # 16:9 AR Button
        $RtButton2          = New-Object System.Windows.Forms.Button
        $RtButton2.Text     = "16:9"
        $RtButton2.Size     = New-Object System.Drawing.Size(60,30)
        $RtButton2.Location = New-Object System.Drawing.Point(80,370)
        $RtButton2.Add_Click({SendRt4k -Commands "aux3" -Text "Aspect Ratio: 16:9"})
        $form.Controls.Add($RtButton2)
    
    
        # Phase Detect Button
        $RtButton3          = New-Object System.Windows.Forms.Button
        $RtButton3.Text     = "Phase"
        $RtButton3.Size     = New-Object System.Drawing.Size(60,30)
        $RtButton3.Location = New-Object System.Drawing.Point(140,370)
        $RtButton3.Add_Click({SendRt4k -Commands "phase" -text "Phase Adjust"})
        $form.Controls.Add($RtButton3)
    
    
        # Gain Detect Button
        $RtButton4          = New-Object System.Windows.Forms.Button
        $RtButton4.Text     = "Gain"
        $RtButton4.Size     = New-Object System.Drawing.Size(60,30)
        $RtButton4.Location = New-Object System.Drawing.Point(200,370)
        $RtButton4.Add_Click({SendRt4k -Commands "gain" -Text "Adjust Gain"})
        $form.Controls.Add($RtButton4)
    
    
        # 1080p Button
        $RtButton5          = New-Object System.Windows.Forms.Button
        $RtButton5.Text     = "1080P"
        $RtButton5.Size     = New-Object System.Drawing.Size(60,30)
        $RtButton5.Location = New-Object System.Drawing.Point(260,370)
        $RtButton5.Add_Click({SendRt4k -Commands "res1080p", "right", "ok" -Text "1080P"})
        $form.Controls.Add($RtButton5)
    
    
        # 4K Button
        $RtButton6          = New-Object System.Windows.Forms.Button
        $RtButton6.Text     = "4K"
        $RtButton6.Size     = New-Object System.Drawing.Size(60,30)
        $RtButton6.Location = New-Object System.Drawing.Point(320,370)
        $RtButton6.Add_Click({SendRt4k -Commands "res4k", "right", "ok" -Text "4k"})
        $form.Controls.Add($RtButton6)
    
    
        # 480P Button
        $RtButton7          = New-Object System.Windows.Forms.Button
        $RtButton7.Text     = "480P"
        $RtButton7.Size     = New-Object System.Drawing.Size(60,30)
        $RtButton7.Location = New-Object System.Drawing.Point(380,370)
        $RtButton7.Add_Click({SendRt4k -Commands "res480p", "right", "ok" -Text "480p"})
        $form.Controls.Add($RtButton7)
    
    
        # Background image
        $Image                      = [system.drawing.image]::FromFile("$PSScriptRoot\GamesCareBG.png") 
        $Form.BackgroundImage       = $Image
        $Form.BackgroundImageLayout = "Center"
    
        # Add timer to auto-refresh the form every couple of seconds to detect automatic switch input changes
    
        # Check if $Timer exists, stop it, and dispose of it
        Cleanup-Timer
    
        $Global:Timer           = New-Object System.Windows.Forms.Timer
        $Global:Timer.Interval  = 2000 # 2000 = 2 seconds between queries
        $Global:Timer.Add_Tick({
            $CurrPort           = (Invoke-RestMethod -Uri "http://$GamesCareIP/ports" -Method Get).Active
            If ($portsel -ne $CurrPort) {
                Write-Host "Port change to port $CurrPort Detected"
                Switch-Input -URL "http://$GamesCareIP/ports?force=$CurrPort" -Command "prof$($CurrPort)"
            }
        })
    
        $global:Timer.Start()

        # Cleanup timer or it will may remain active depending on launch method.
        $Form.Add_FormClosed({
            $global:Timer.Stop()
            $global:Timer.Dispose()
            $global:Timer = $null
        })
        
        # Show the form
        Switch-Input -URL "http://$GamesCareIP/ports" # Perform initial port query
        $form.Add_Shown
        [void]$form.ShowDialog()
    }
    
    
# Main Loop

Clear-Host

$SwitchName = "gcswitch.local" # gcswitch is the default, change if necessary.
#$GamesCareIP = "10.0.1.125" # Uncomment this and set IP appropriately if switch is not detected via DNS.

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Write-Host "Searching for GamesCare switch on local network...`n"

#If (Test-Connection $GamesCareIP -Count 1 -ErrorAction SilentlyContinue) {$SwitchName = $GamesCareIP}

$Script:GamesCareIP = (Test-Connection $SwitchName -Count 1 -ErrorAction SilentlyContinue).IPV4Address.IPAddressToString
If ($GamesCareIP) 
    {
        Write-Host "Games Care Switch found @ $GamesCareIP"
    }
    ELSE
    {
        Write-Host "Games Care Switch not found - You may need to specify your IP explicitly."
        Exit 1
    }


Select-Input

