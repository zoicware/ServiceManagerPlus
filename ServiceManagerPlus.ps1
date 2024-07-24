if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]'Administrator')) {
    Start-Process PowerShell.exe -ArgumentList ("-NoProfile -ExecutionPolicy Bypass -File `"{0}`"" -f $PSCommandPath) -Verb RunAs
    Exit	
}

function Get-ServicePlus {
    #Get all Services
    $services = Get-Service -Name * -ErrorAction SilentlyContinue
    $svcWMI = Get-WmiObject -Class Win32_Service
    $servicesPlus = @()

    foreach ($service in $services) {
        #Get the Binary Path
        $query = sc.exe qc $service.ServiceName | Select-String 'BINARY_PATH_NAME'
        $binPath = if ($query) { ($query -split ':', 2)[1].Trim() } else { '' }
        #Get the Service Description
        $svcDesc = $svcWMI | Where-Object { $_.Name -eq $service.ServiceName }
        #Custom Object to fill Data Grid Table
        $servicePlus = [PSCustomObject]@{
            DisplayName           = $service.DisplayName
            Name                  = $service.ServiceName
            Status                = $service.Status
            StartType             = $service.StartType
            'DependentService(s)' = ($service.ServicesDependedOn | ForEach-Object { $_.DisplayName }) -join ', '
            BinaryPath            = $binPath
            Description           = $svcDesc.Description
        }

        $servicesPlus += $servicePlus
    }

    return $servicesPlus
}

#returns string array
function Get-SelectedService {
    $selectedRows = $dataGridView.SelectedRows
    $serviceNames = @()
    foreach ($row in $selectedRows) {
        $cells = $row.Cells
        $serviceNames += $cells[1].Value
    }
    return $serviceNames
}

function Run-Trusted([String]$command) {

    Stop-Service -Name TrustedInstaller -Force -ErrorAction SilentlyContinue
    #get bin path to revert later
    $query = sc.exe qc TrustedInstaller | Select-String 'BINARY_PATH_NAME'
    #limit split to only 2 parts
    $binPath = $query -split ':', 2
    #convert command to base64 to avoid errors with spaces
    $bytes = [System.Text.Encoding]::Unicode.GetBytes($command)
    $base64Command = [Convert]::ToBase64String($bytes)
    #change bin to command
    sc.exe config TrustedInstaller binPath= "cmd.exe /c powershell.exe -encodedcommand $base64Command" | Out-Null
    #run the command
    sc.exe start TrustedInstaller | Out-Null
    #set bin back to default
    sc.exe config TrustedInstaller binPath= "$($binPath[1].Trim())" | Out-Null
    Stop-Service -Name TrustedInstaller -Force -ErrorAction SilentlyContinue

}

Write-Host 'Getting Services Please Wait...'


Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Service Manager Plus'
$form.BackColor = [System.Drawing.Color]::FromArgb(43, 43, 42)
$form.WindowState = 'Maximized'

$openReg = New-Object System.Windows.Forms.Button
$openReg.Text = 'Open in Registry'
$openReg.Size = New-Object System.Drawing.Size(150, 28)
try {
    $image = [System.Drawing.Image]::FromFile('Assets\Registry.png')
    #width,height
    $resizedImage = New-Object System.Drawing.Bitmap $image, 25, 19

    # Set the button image
    $openReg.Image = $resizedImage
    $openReg.ImageAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $openReg.TextImageRelation = [System.Windows.Forms.TextImageRelation]::ImageBeforeText
}
catch {
    Write-Host 'Missing Asset (Registry Icon)' -ForegroundColor Red
}
$openReg.ForeColor = 'White'
$openReg.Location = New-Object System.Drawing.Point(280, 2)
#add mouse over effect without using flat style
$openReg.Add_MouseEnter({
        $openReg.BackColor = [System.Drawing.Color]::FromArgb(64, 64, 64)
    })

$openReg.Add_MouseLeave({
        $openReg.BackColor = [System.Drawing.Color]::FromArgb(43, 43, 42)
    })
$openReg.Add_Click({
        $selected = Get-SelectedService
        if (!($selected)) {
            Write-Host 'No Service Selected...'
        }
        else {
            #close regedit if its open
            Stop-Process -Name regedit -Force -ErrorAction SilentlyContinue
            #open registry to first row selected
            #if array only has 1 item just use $selected
            if ($selected.Count -eq 1) {
                $Path = "HKLM\SYSTEM\ControlSet001\Services\$($selected)"
            }
            else {
                $Path = "HKLM\SYSTEM\ControlSet001\Services\$($selected[0])"
            }
            Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Applets\Regedit' -Name Lastkey -Value $Path -Type String -Force
            Start-Process 'regedit.exe'
        }
    })
$form.Controls.Add($openReg)

# Create the TextBox for search
$searchBox = New-Object System.Windows.Forms.TextBox
$searchBox.Location = New-Object System.Drawing.Point(5, 5)
$searchBox.Size = New-Object System.Drawing.Size(200, 20)
$form.Controls.Add($searchBox)

# Add a TextChanged event to the TextBox
$searchBox.add_TextChanged({
        $dataTable.DefaultView.RowFilter = "DisplayName LIKE '%" + $searchBox.Text + "%' OR Name LIKE '%" + $searchBox.Text + "%'"
    })

# Create the PictureBox
$pictureBox = New-Object System.Windows.Forms.PictureBox
$pictureBox.Location = New-Object System.Drawing.Point(215, 5) 
$pictureBox.Size = New-Object System.Drawing.Size(30, 20) 
$pictureBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
$imagePath = 'Assets\Search.png'
try {
    $image = [System.Drawing.Image]::FromFile($imagePath)
}
catch {
    Write-Host 'Missing Asset (Search Icon)' -ForegroundColor Red
}
$pictureBox.Image = $image
$form.Controls.Add($pictureBox)


# Create the DataGridView
$Global:dataGridView = New-Object System.Windows.Forms.DataGridView
$dataGridView.Location = New-Object System.Drawing.Point(-40, 30)
$dataGridView.ReadOnly = $true
$dataGridView.BackgroundColor = [System.Drawing.Color]::FromArgb(43, 43, 42)
$dataGridView.ForeColor = 'White'
$dataGridView.GridColor = [System.Drawing.Color]::FromArgb(64, 64, 64)
$dataGridView.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(43, 43, 42)
$dataGridView.DefaultCellStyle.ForeColor = 'White'
$dataGridView.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(176, 176, 176) 
$dataGridView.DefaultCellStyle.SelectionForeColor = 'Black'
$dataGridView.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(64, 64, 64)
$dataGridView.ColumnHeadersDefaultCellStyle.ForeColor = 'White'
$dataGridView.RowHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(64, 64, 64)
$dataGridView.RowHeadersDefaultCellStyle.ForeColor = 'White'
$dataGridView.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(64, 64, 64)
$dataGridView.AlternatingRowsDefaultCellStyle.ForeColor = 'White'
$dataGridView.ScrollBars = [System.Windows.Forms.ScrollBars]::Both
$dataGridView.AllowUserToResizeColumns = $true
$dataGridView.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect

$form.add_Resize({
        # Adjust the DataGridView's size when the form is resized
        $dataGridView.Size = New-Object System.Drawing.Size(($form.Width + 25), ($form.Height - 70))
    })

$form.Controls.Add($dataGridView)

# Retrieve the service data
$servicesPlus = Get-ServicePlus

# Convert the data to a DataTable
$dataTable = New-Object System.Data.DataTable
$columnNames = $servicesPlus[0].PSObject.Properties.Name 
foreach ($name in $columnNames) {
    $dataTable.Columns.Add($name) | Out-Null
}
foreach ($service in $servicesPlus) {
    $row = $dataTable.NewRow()
    $rows = $service.PSObject.Properties
    foreach ($r in $rows) {
        $row[$r.Name] = $service.$($r.Name) 
    }
    $dataTable.Rows.Add($row) | Out-Null
}

# Bind the DataTable to the DataGridView
$dataGridView.DataSource = $dataTable

# Create the label
$label = New-Object System.Windows.Forms.Label
$label.Text = 'Set Service:'
$label.ForeColor = 'White'
$label.Location = New-Object System.Drawing.Point(440, 5)
$label.AutoSize = $true
$label.Font = New-Object System.Drawing.Font('Segoe UI', 10)
$form.Controls.Add($label)

# Create the buttons
$manualButton = New-Object System.Windows.Forms.Button
$manualButton.Text = 'Manual'
$manualButton.Size = New-Object System.Drawing.Size(90, 30)
$manualButton.Location = New-Object System.Drawing.Point(520, 2)
$manualButton.ForeColor = 'White'
$manualButton.Add_MouseEnter({
        $manualButton.BackColor = [System.Drawing.Color]::FromArgb(64, 64, 64)
    })

$manualButton.Add_MouseLeave({
        $manualButton.BackColor = [System.Drawing.Color]::FromArgb(43, 43, 42)
    })
$form.Controls.Add($manualButton)

$automaticButton = New-Object System.Windows.Forms.Button
$automaticButton.Text = 'Automatic'
$automaticButton.Size = New-Object System.Drawing.Size(90, 30)
$automaticButton.Location = New-Object System.Drawing.Point(610, 2)
$automaticButton.ForeColor = 'White'
$automaticButton.Add_MouseEnter({
        $automaticButton.BackColor = [System.Drawing.Color]::FromArgb(64, 64, 64)
    })

$automaticButton.Add_MouseLeave({
        $automaticButton.BackColor = [System.Drawing.Color]::FromArgb(43, 43, 42)
    })
$form.Controls.Add($automaticButton)

$disabledButton = New-Object System.Windows.Forms.Button
$disabledButton.Text = 'Disabled'
$disabledButton.Size = New-Object System.Drawing.Size(90, 30)
$disabledButton.Location = New-Object System.Drawing.Point(700, 2)
$disabledButton.ForeColor = 'White'
$disabledButton.Add_MouseEnter({
        $disabledButton.BackColor = [System.Drawing.Color]::FromArgb(64, 64, 64)
    })

$disabledButton.Add_MouseLeave({
        $disabledButton.BackColor = [System.Drawing.Color]::FromArgb(43, 43, 42)
    })
$form.Controls.Add($disabledButton)

$manualButton.Add_Click({
        $selectedServices = Get-SelectedService
        if (!($selectedServices)) {
            Write-Host 'No Service Selected...'
        }
        else {
            if ($selectedServices.Count -gt 1) {
                foreach ($service in $selectedServices) {
                    $command += "Set-Service -Name $service -StartupType Manual; "
                }
                Run-Trusted -command $command
            }
            else {
                $command = "Set-Service -Name $selectedServices -StartupType Manual"
                Run-Trusted -command $command
            }
        }
    })

$automaticButton.Add_Click({
        $selectedServices = Get-SelectedService
        if (!($selectedServices)) {
            Write-Host 'No Service Selected...'
        }
        else {
            if ($selectedServices.Count -gt 1) {
                foreach ($service in $selectedServices) {
                    $command += "Set-Service -Name $service -StartupType Automatic; "
                }
                Run-Trusted -command $command
            }
            else {
                $command = "Set-Service -Name $selectedServices -StartupType Automatic"
                Run-Trusted -command $command
            }
        }
    })

$disabledButton.Add_Click({
        $selectedServices = Get-SelectedService
        if (!($selectedServices)) {
            Write-Host 'No Service Selected...'
        }
        else {
            if ($selectedServices.Count -gt 1) {
                foreach ($service in $selectedServices) {
                    $command += "Set-Service -Name $service -StartupType Disabled; "
                }
                Run-Trusted -command $command
            }
            else {
                $command = "Set-Service -Name $selectedServices -StartupType Disabled"
                Run-Trusted -command $command
            }
        }
    })

#set autosize mode for each column
$dataGridView.Columns['DisplayName'].AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::None
$dataGridView.Columns['Name'].AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::None
$dataGridView.Columns['Status'].AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::None
$dataGridView.Columns['StartType'].AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::None
$dataGridView.Columns['DependentService(s)'].AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::None
$dataGridView.Columns['BinaryPath'].AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::None
$dataGridView.Columns['Description'].AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells

#set custom width
$dataGridView.Columns['DisplayName'].Width = 230
$dataGridView.Columns['Name'].Width = 170
$dataGridView.Columns['Status'].Width = 80
$dataGridView.Columns['StartType'].Width = 80
$dataGridView.Columns['DependentService(s)'].Width = 200
$dataGridView.Columns['BinaryPath'].Width = 200

#remove first row selection
$form.Add_Shown({
        $dataGridView.ClearSelection()
    })

# Show the form
[void]$form.ShowDialog()