function Send-SerialCommand {
    param (
     [string]$Command # Command to send via serial port
     )

    # Define the COM port and settings
    $portName = "COM3"    # Replace with your COM port (e.g., COM1, COM2, etc.)
    $baudRate = 115200      # The Rt4K natively uses 115200
    $serialPort = New-Object System.IO.Ports.SerialPort
    $serialPort.PortName = $portName
    $serialPort.BaudRate = $baudRate

try {
    # Open the serial port
    $serialPort.Open()
    Write-Host "Serial port $portName opened successfully."

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
            Write-Host "Serial port $portName closed."
        }
    }
}


Function SendRt4K {
    param (
     [string]$Command # Command to send via serial port
     )

    # Code block based on original by Carter300: https://github.com/carter300/RetroTink4K-PCRemote/tree/main

    # Debug output for the argument
    Write-Host "Passed argument for input: '$Command'" -ForegroundColor Yellow

    # Check if the input argument was provided
    if ([string]::IsNullOrEmpty($Command)) {
        Write-Host "Error: No hex string provided. Please run with -Input 'HexString'." -ForegroundColor Red
      #  exit 1
    }

    # List of valid commands
    $validCommands = @(
        "pwr", "menu", "up", "down", "left", "right", "ok", "back", "diag", "stat",
        "input", "output", "scaler", "sfx", "adc", "prof", "prof1", "prof2", "prof3",
        "prof4", "prof5", "prof6", "prof7", "prof8", "prof9", "prof10", "prof11", "prof12",
        "gain", "phase", "pause", "safe", "genlock", "buffer", "res4k", "res1080p", "res1440p",
        "res480p", "res1", "res2", "res3", "res4", "aux1", "aux2", "aux3", "aux4", "aux5",
        "aux6", "aux7", "aux8", "pwr on"
    )

    # Check if the input is a valid command
    if ($validCommands -contains $Command) {
        # Prepend "remote" to the command
        $command = "remote $Command`n"
    } else {
        Write-Host "Unknown command: '$Command'. Please enter a valid command." -ForegroundColor Red
        exit 1
    }

    # Debug output for the command
    Write-Host "Command to be executed: $command" -ForegroundColor Cyan
  
    # Call the function to send the serial command
    Send-SerialCommand -Command $Command

    # List of commands: 
    # https://consolemods.org/wiki/AV:RetroTINK-4K#USB_Serial_Configuration
    # Execute with: .\RT4KRemoteSerial.ps1 -Input "down"
}


Function Select-Input { # Massively inefficient form by Steve Wells for use with the GamesCare 8 port SCART switch
    # Create the form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "GamesCare Switch - RetroTink 4K Profile Selector"
    $form.Size = New-Object System.Drawing.Size(470, 450)
    $form.StartPosition = "CenterScreen"

    # Create a label to display the current port
    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Current Port:"
    $label.Location = New-Object System.Drawing.Point(300, 12)
    $label.Size = New-Object System.Drawing.Size(100, 20)
    $form.Controls.Add($label)
    $label.BackColor 		 = [System.Drawing.Color]::FromName("Transparent")
    $label.ForeColor 		 = [System.Drawing.Color]::Black

    #GamesCare Button 1
    $GCbutton1 = New-Object System.Windows.Forms.Button
    $GCbutton1.Size = New-Object System.Drawing.Size(200, 30)
    $GCbutton1.Location = New-Object System.Drawing.Point(120, 60)
    $GCButton1.Text = "Port 1"
    $GCButton1.Add_Click({
    Write-Host "Changing Input to: "http://$GamesCareIP/ports?force=1""
            SendRt4K -Command "prof1"
            $result = Invoke-RestMethod -Uri "http://$GamesCareIP/ports?force=1" -Method Get
            $portsel = $result.active
            $label.Text = "Current Port: $portsel"
            })
    $form.Controls.Add($GCbutton1)

        #GamesCare Button 2
    $GCbutton2 = New-Object System.Windows.Forms.Button
    $GCbutton2.Size = New-Object System.Drawing.Size(200, 30)
    $GCbutton2.Location = New-Object System.Drawing.Point(120, 90)
    $GCButton2.Text = "Port 2"
    $GCButton2.Add_Click({
    Write-Host "Changing Input to: "http://$GamesCareIP/ports?force=2""
    SendRt4K -Command "prof2"
            $result = Invoke-RestMethod -Uri "http://$GamesCareIP/ports?force=2" -Method Get
            $portsel = $result.active
            $label.Text = "Current Port: $portsel"
            })
    $form.Controls.Add($GCbutton2)


        #GamesCare Button 3
    $GCbutton3 = New-Object System.Windows.Forms.Button
    $GCbutton3.Size = New-Object System.Drawing.Size(200, 30)
    $GCbutton3.Location = New-Object System.Drawing.Point(120, 120)
    $GCButton3.Text = "Port 3"
    $GCButton3.Add_Click({
    Write-Host "Changing Input to: "http://$GamesCareIP/ports?force=3""
    SendRt4K -Command "prof3"
            $result = Invoke-RestMethod -Uri "http://$GamesCareIP/ports?force=3" -Method Get
            $portsel = $result.active
            $label.Text = "Current Port: $portsel"
            })
    $form.Controls.Add($GCbutton3)


        #GamesCare Button 4
    $GCbutton4 = New-Object System.Windows.Forms.Button
    $GCbutton4.Size = New-Object System.Drawing.Size(200, 30)
    $GCbutton4.Location = New-Object System.Drawing.Point(120, 150)
    $GCButton4.Text = "Port 4"
    $GCButton4.Add_Click({
    Write-Host "Changing Input to: "http://$GamesCareIP/ports?force=4""
    SendRt4K -Command "prof4"
            $result = Invoke-RestMethod -Uri "http://$GamesCareIP/ports?force=4" -Method Get
            $portsel = $result.active
            $label.Text = "Current Port: $portsel"
            })
    $form.Controls.Add($GCbutton4)


        #GamesCare Button 5
    $GCbutton5 = New-Object System.Windows.Forms.Button
    $GCbutton5.Size = New-Object System.Drawing.Size(200, 30)
    $GCbutton5.Location = New-Object System.Drawing.Point(120, 180)
    $GCButton5.Text = "Port 5"
    $GCButton5.Add_Click({
    Write-Host "Changing Input to: "http://$GamesCareIP/ports?force=5""
    SendRt4K -Command "prof5"
            $result = Invoke-RestMethod -Uri "http://$GamesCareIP/ports?force=5" -Method Get
            $portsel = $result.active
            $label.Text = "Current Port: $portsel"
            })
    $form.Controls.Add($GCbutton5)


        #GamesCare Button 6
    $GCbutton6 = New-Object System.Windows.Forms.Button
    $GCbutton6.Size = New-Object System.Drawing.Size(200, 30)
    $GCbutton6.Location = New-Object System.Drawing.Point(120, 210)
    $GCButton6.Text = "Port 6"
    $GCButton6.Add_Click({
    Write-Host "Changing Input to: "http://$GamesCareIP/ports?force=6""
    SendRt4K -Command "prof6"
            $result = Invoke-RestMethod -Uri "http://$GamesCareIP/ports?force=6" -Method Get
            $portsel = $result.active
            $label.Text = "Current Port: $portsel"
            })
    $form.Controls.Add($GCbutton6)


        #GamesCare Button 7
    $GCbutton7 = New-Object System.Windows.Forms.Button
    $GCbutton7.Size = New-Object System.Drawing.Size(200, 30)
    $GCbutton7.Location = New-Object System.Drawing.Point(120, 240)
    $GCButton7.Text = "Port 7"
    $GCButton7.Add_Click({
    Write-Host "Changing Input to: "http://$GamesCareIP/ports?force=7""
    SendRt4K -Command "prof7"
            $result = Invoke-RestMethod -Uri "http://$GamesCareIP/ports?force=7" -Method Get
            $portsel = $result.active
            $label.Text = "Current Port: $portsel"
            })
    $form.Controls.Add($GCbutton7)


        #GamesCare Button 8
    $GCbutton8 = New-Object System.Windows.Forms.Button
    $GCbutton8.Size = New-Object System.Drawing.Size(200, 30)
    $GCbutton8.Location = New-Object System.Drawing.Point(120, 270)
    $GCButton8.Text = "Port 8"
    $GCButton8.Add_Click({
    Write-Host "Changing Input to: "http://$GamesCareIP/ports?force=8""
    SendRt4K -Command "prof8"
            $result = Invoke-RestMethod -Uri "http://$GamesCareIP/ports?force=8" -Method Get
            $portsel = $result.active
            $label.Text = "Current Port: $portsel"
            })
    $form.Controls.Add($GCbutton8)


        #GamesCare Button Auto
    $GCbutton0 = New-Object System.Windows.Forms.Button
    $GCbutton0.Size = New-Object System.Drawing.Size(200, 30)
    $GCbutton0.Location = New-Object System.Drawing.Point(120, 300)
    $GCButton0.Text = "Auto"
    $GCButton0.Add_Click({
    Write-Host "Changing Input to: "http://$GamesCareIP/ports?force=0""
    #SendRt4K -Command "prof2"
            $result = Invoke-RestMethod -Uri "http://$GamesCareIP/ports?force=0" -Method Get
            $portsel = $result.active
            $label.Text = "Current Port: $portsel"
            })
    $form.Controls.Add($GCbutton0)


          #GamesCare Button Reset
    $GCbuttonR = New-Object System.Windows.Forms.Button
    $GCbuttonR.Size = New-Object System.Drawing.Size(120, 30)
    $GCbuttonR.Location = New-Object System.Drawing.Point(320, 150)
    $GCButtonR.Text = "Reboot"
    $GCButtonR.Add_Click({
    Write-Host "Rebooting Switch"
            $result = Invoke-RestMethod -Uri "http://$GamesCareIP/settings?reboot=1" -Method Get -ErrorAction Ignore # Will timeout, to be expected.
            $portsel = $result.active
            $label.Text = "Current Port: $portsel"
            })
  #  $form.Controls.Add($GCbuttonR) - Don't add, behaves strangely.


    # Create a label for Rt4K settings
    $label2 = New-Object System.Windows.Forms.Label
    $label2.Text = "RetroTink 4K:"
    $label2.Location = New-Object System.Drawing.Point(20, 350)
    $label2.Size = New-Object System.Drawing.Size(100, 20)
    $form.Controls.Add($label2)
    $label2.BackColor 		 = [System.Drawing.Color]::FromName("Transparent")
    $label2.ForeColor 		 = [System.Drawing.Color]::Black


    # 4:3 AR Button
    $RtButton1 = New-Object System.Windows.Forms.Button
    $RtButton1.Text = "4:3"
    $RtButton1.Size = New-Object System.Drawing.Size(60,30)
    $RtButton1.Location = New-Object System.Drawing.Point(20,370)
    $RtButton1.Add_Click({
        Write-Host "Aspect Ratio: 4:3" -ForegroundColor Green
                SendRt4K -Command "aux2"
                SendRt4K -Command "right"
                SendRt4K -Command "ok"
            })
             $form.Controls.Add($RtButton1)
    
    # 16:9 AR Button
    $RtButton2 = New-Object System.Windows.Forms.Button
    $RtButton2.Text = "16:9"
    $RtButton2.Size = New-Object System.Drawing.Size(60,30)
    $RtButton2.Location = New-Object System.Drawing.Point(80,370)
    $RtButton2.Add_Click({
        Write-Host "Aspect Ratio: 16:9" -ForegroundColor Green
                SendRt4K -Command "aux3"
                SendRt4K -Command "right"
                SendRt4K -Command "ok"
            })
             $form.Controls.Add($RtButton2)

    # Phase Detect Button
    $RtButton3 = New-Object System.Windows.Forms.Button
    $RtButton3.Text = "Phase"
    $RtButton3.Size = New-Object System.Drawing.Size(60,30)
    $RtButton3.Location = New-Object System.Drawing.Point(140,370)
    $RtButton3.Add_Click({
        Write-Host "Phase Adjust" -ForegroundColor Green
                SendRt4K -Command "phase"
            })
             $form.Controls.Add($RtButton3)

    # Gain Detect Button
    $RtButton4 = New-Object System.Windows.Forms.Button
    $RtButton4.Text = "Gain"
    $RtButton4.Size = New-Object System.Drawing.Size(60,30)
    $RtButton4.Location = New-Object System.Drawing.Point(200,370)
    $RtButton4.Add_Click({
        Write-Host "Gain Adjust" -ForegroundColor Green
                SendRt4K -Command "gain"
            })
             $form.Controls.Add($RtButton4)


    # 1080p Button
    $RtButton5 = New-Object System.Windows.Forms.Button
    $RtButton5.Text = "1080P"
    $RtButton5.Size = New-Object System.Drawing.Size(60,30)
    $RtButton5.Location = New-Object System.Drawing.Point(260,370)
    $RtButton5.Add_Click({
        Write-Host "1080p" -ForegroundColor Green
                SendRt4K -Command "res1080p"
                SendRt4K -Command "right"
                SendRt4K -Command "ok"
            })
             $form.Controls.Add($RtButton5)

    # 4K Button
    $RtButton6 = New-Object System.Windows.Forms.Button
    $RtButton6.Text = "4K"
    $RtButton6.Size = New-Object System.Drawing.Size(60,30)
    $RtButton6.Location = New-Object System.Drawing.Point(320,370)
    $RtButton6.Add_Click({
        Write-Host "4K" -ForegroundColor Green
                SendRt4K -Command "res4k"
                SendRt4K -Command "right"
                SendRt4K -Command "ok"
            })
             $form.Controls.Add($RtButton6)

    # Safe Button
    $RtButton7 = New-Object System.Windows.Forms.Button
    $RtButton7.Text = "480P"
    $RtButton7.Size = New-Object System.Drawing.Size(60,30)
    $RtButton7.Location = New-Object System.Drawing.Point(380,370)
    $RtButton7.Add_Click({
        Write-Host "480P" -ForegroundColor Green
                SendRt4K -Command "res480p"
                SendRt4K -Command "right"
                SendRt4K -Command "ok"
            })
             $form.Controls.Add($RtButton7)

             
             $Image = [system.drawing.image]::FromFile("$PSScriptRoot\GamesCareBG.png") 
              $Form.BackgroundImage = $Image
 $Form.BackgroundImageLayout = "Center"

    # Show the form
 $result = Invoke-RestMethod -Uri "http://$GamesCareIP/ports?force=1" -Method Get
            $portsel = $result.active
            $label.Text = "Current Port: $portsel"
    $form.Add_Shown
    [void]$form.ShowDialog()
}

# Main Loop
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Write-Host "Searching for GamesCare switch on local network...`n"
$Script:GamesCareIP = (Test-Connection gcswitch.local -Count 1).IPV4Address.IPAddressToString
If ($GamesCareIP) 
    {
        Write-Host "Games Care Switch found @ $GamesCareIP"
    }
    ELSE
    {
        Write-Host "Games Care Switch not found - You may need to specify your IP explicitly."
        # GamesCareIP = "10.0.1.100" # Set IP here if necessary.
     #   Exit 1
    }

Select-Input
