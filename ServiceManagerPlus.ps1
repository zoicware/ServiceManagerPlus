if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]'Administrator')) {
    Start-Process PowerShell.exe -ArgumentList ("-NoProfile -ExecutionPolicy Bypass -File `"{0}`"" -f $PSCommandPath) -Verb RunAs
    Exit	
}

#set mode service by default
$Global:Mode = 'service'

function Get-DriverPlus {
    $drivers = Get-CimInstance -ClassName Win32_SystemDriver
    $driversList = [System.Collections.ArrayList]::new()
    $i = 0
    $totalDrivers = $drivers.Count
    foreach ($driver in $drivers) {
        $i++
        $progressbar1.Value = $(($i / $totalDrivers) * 100) 
        $driverPlus = [PSCustomObject]@{
            DisplayName = $driver.DisplayName
            Name        = $driver.Name
            Status      = $driver.State
            StartType   = $driver.StartMode
            BinaryPath  = $driver.PathName
        }
        $driversList.Add($driverPlus) | Out-Null
    }
    return $driversList
}


function Get-ServicePlus {
    #Get all Services
    $services = Get-Service -Name * -ErrorAction SilentlyContinue
    $totalServices = $services.Count
    $svcWMI = Get-CimInstance -Class Win32_Service | Group-Object -Property Name -AsHashTable -AsString
    #use .net array list since .Add method is much faster than +=
    $servicesPlus = [System.Collections.ArrayList]::new()
    $i = 0
    foreach ($service in $services) {
        $i++
        #Write-Progress -Activity 'Getting Services' -Status "$([math]::Round(($i / $totalServices) * 100)) % Complete" -PercentComplete $(($i / $totalServices) * 100) 
        $progressbar1.Value = $(($i / $totalServices) * 100) 
        #Get the Binary Path
        $binPath = $svcWMI[$service.ServiceName].PathName
        #Get the Service Description
        $svcDesc = $svcWMI[$service.ServiceName]
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

        $null = $servicesPlus.Add($servicePlus)
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
    $service = Get-WmiObject -Class Win32_Service -Filter "Name='TrustedInstaller'"
    $DefaultBinPath = $service.PathName
    #convert command to base64 to avoid errors with spaces
    $bytes = [System.Text.Encoding]::Unicode.GetBytes($command)
    $base64Command = [Convert]::ToBase64String($bytes)
    #change bin to command
    sc.exe config TrustedInstaller binPath= "cmd.exe /c powershell.exe -encodedcommand $base64Command" | Out-Null
    #run the command
    sc.exe start TrustedInstaller | Out-Null
    #set bin back to default
    sc.exe config TrustedInstaller binpath= "`"$DefaultBinPath`"" | Out-Null
    Stop-Service -Name TrustedInstaller -Force -ErrorAction SilentlyContinue

}

function Stop-SelectedService {
    param(
        [switch]$service,
        [switch]$driver
    )
    if ($service) {
        $selectedServices = Get-SelectedService
        if (!($selectedServices)) {
            Write-Host 'No Service Selected...'
        }
        else {
            if ($selectedServices.Count -gt 1) {
                foreach ($serviceName in $selectedServices) {
                    $command += "Stop-Service -Name $serviceName -Force; "
                }
                Run-Trusted -command $command
            }
            else {
                $command = "Stop-Service -Name $selectedServices -Force"
                Run-Trusted -command $command
            }
        }
    }
    else {
        #stop driver
        $selectedDrivers = Get-SelectedService
        if (!$selectedDrivers) {
            Write-Host 'No Driver Selected...'
        }
        else {
            if ($selectedDrivers -gt 1) {
                foreach ($driverName in $selectedDrivers) {
                    $command += "Sc.exe stop $driverName; "
                }
                Run-Trusted -command $command
            }
            else {
                $command = "Sc.exe stop $selectedDrivers"
                Run-Trusted -command $command
            }
        }
    }
    
}

function Start-SelectedService {
    param (
        [switch]$service,
        [switch]$driver
    )
    if ($service) {
        $selectedServices = Get-SelectedService
        if (!($selectedServices)) {
            Write-Host 'No Service Selected...'
        }
        else {
            if ($selectedServices.Count -gt 1) {
                foreach ($serviceName in $selectedServices) {
                    $command += "Start-Service -Name $serviceName; "
                }
                Run-Trusted -command $command
            }
            else {
                $command = "Start-Service -Name $selectedServices"
                Run-Trusted -command $command
            }
        }
    }
    else {
        #start driver
        $selectedDrivers = Get-SelectedService
        if (!$selectedDrivers) {
            Write-Host 'No Driver Selected...'
        }
        else {
            if ($selectedDrivers -gt 1) {
                foreach ($driverName in $selectedDrivers) {
                    $command += "Sc.exe start $driverName; "
                }
                Run-Trusted -command $command
            }
            else {
                $command = "Sc.exe start $selectedDrivers"
                Run-Trusted -command $command
            }
        }
    }
    
}

function Set-Manual {
    param (
        [switch]$driver,
        [switch]$service
    )
    if ($service) {
        $selectedServices = Get-SelectedService
        if (!($selectedServices)) {
            Write-Host 'No Service Selected...'
        }
        else {
            if ($selectedServices.Count -gt 1) {
                foreach ($serviceName in $selectedServices) {
                    $command += "Set-Service -Name $serviceName -StartupType Manual; "
                }
                Run-Trusted -command $command
            }
            else {
                $command = "Set-Service -Name $selectedServices -StartupType Manual"
                Run-Trusted -command $command
            }
        }
    }
    else {
        #set driver
        $selectedDrivers = Get-SelectedService
        if (!$selectedDrivers) {
            Write-Host 'No Driver Selected...'
        }
        else {
            if ($selectedDrivers.Count -gt 1) {
                foreach ($driverName in $selectedDrivers) {
                    $command += "Sc.exe config $driverName start= demand; "
                }
                Run-Trusted -command $command
            }
            else {
                $command = "Sc.exe config $selectedDrivers start= demand"
                Run-Trusted -command $command
            }
        }
        
    }
    
}

function Set-Automatic {
    param(
        [switch]$driver,
        [switch]$service
    )
    if ($service) {
        $selectedServices = Get-SelectedService
        if (!($selectedServices)) {
            Write-Host 'No Service Selected...'
        }
        else {
            if ($selectedServices.Count -gt 1) {
                foreach ($serviceName in $selectedServices) {
                    $command += "Set-Service -Name $serviceName -StartupType Automatic; "
                }
                Run-Trusted -command $command
            }
            else {
                $command = "Set-Service -Name $selectedServices -StartupType Automatic"
                Run-Trusted -command $command
            }
        }
    }
    else {
        #set driver
        $selectedDrivers = Get-SelectedService
        if (!$selectedDrivers) {
            Write-Host 'No Driver Selected...'
        }
        else {
            if ($selectedDrivers.Count -gt 1) {
                foreach ($driverName in $selectedDrivers) {
                    $command += "Sc.exe config $driverName start= auto; "
                }
                Run-Trusted -command $command
            }
            else {
                $command = "Sc.exe config $selectedDrivers start= auto"
                Run-Trusted -command $command
            }
        }
        
    }
    
}

function Set-Disabled {
    param(
        [switch]$driver,
        [switch]$service
    )
    if ($service) {
        $selectedServices = Get-SelectedService
        if (!($selectedServices)) {
            Write-Host 'No Service Selected...'
        }
        else {
            if ($selectedServices.Count -gt 1) {
                foreach ($serviceName in $selectedServices) {
                    $command += "Set-Service -Name $serviceName -StartupType Disabled; "
                }
                Run-Trusted -command $command
            }
            else {
                $command = "Set-Service -Name $selectedServices -StartupType Disabled"
                Run-Trusted -command $command
            }
        }
    }
    else {
        #set driver
        $selectedDrivers = Get-SelectedService
        if (!$selectedDrivers) {
            Write-Host 'No Driver Selected...'
        }
        else {
            if ($selectedDrivers.Count -gt 1) {
                foreach ($driverName in $selectedDrivers) {
                    $command += "Sc.exe config $driverName start= disabled; "
                }
                Run-Trusted -command $command
            }
            else {
                $command = "Sc.exe config $selectedDrivers start= disabled"
                Run-Trusted -command $command
            }
        }
    }
    
}


function Set-System {
    #set driver
    $selectedDrivers = Get-SelectedService
    if (!$selectedDrivers) {
        Write-Host 'No Driver Selected...'
    }
    else {
        if ($selectedDrivers.Count -gt 1) {
            foreach ($driverName in $selectedDrivers) {
                $command += "Sc.exe config $driverName start= system; "
            }
            Run-Trusted -command $command
        }
        else {
            $command = "Sc.exe config $selectedDrivers start= system"
            Run-Trusted -command $command
        }
    }
}

function Set-Boot {
    #set driver
    $selectedDrivers = Get-SelectedService
    if (!$selectedDrivers) {
        Write-Host 'No Driver Selected...'
    }
    else {
        if ($selectedDrivers.Count -gt 1) {
            foreach ($driverName in $selectedDrivers) {
                $command += "Sc.exe config $driverName start= boot; "
            }
            Run-Trusted -command $command
        }
        else {
            $command = "Sc.exe config $selectedDrivers start= boot"
            Run-Trusted -command $command
        }
    }
}


Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Service Manager Plus'
$form.BackColor = [System.Drawing.Color]::FromArgb(43, 43, 42)
$form.WindowState = 'Maximized'
try {
    $form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon('Assets\WindowsLogo.ico')
}
catch {
    Write-Host 'Missing Asset (Title Icon)' -ForegroundColor Red
}


$Global:progressBar1 = New-Object System.Windows.Forms.ProgressBar
$progressBar1.Location = New-Object System.Drawing.Point(500, 400)
$progressBar1.Size = New-Object System.Drawing.Size(300, 25)
$progressBar1.Style = 'Marquee'
$progressBar1.Visible = $false
$form.Controls.Add($progressBar1)

$labelLoading = New-Object System.Windows.Forms.Label
$labelLoading.Text = 'Loading'
$labelLoading.ForeColor = 'White'
$labelLoading.Location = New-Object System.Drawing.Point(350, 400)
$labelLoading.AutoSize = $true
$labelLoading.Font = New-Object System.Drawing.Font('Segoe UI', 13)
$labelLoading.Visible = $false
$form.Controls.Add($labelLoading)

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
$searchBox.Location = New-Object System.Drawing.Point(60, 5)
$searchBox.Size = New-Object System.Drawing.Size(150, 20)
$form.Controls.Add($searchBox)

# Add a TextChanged event to the TextBox
$searchBox.add_TextChanged({
        $dataTable.DefaultView.RowFilter = "DisplayName LIKE '%" + $searchBox.Text + "%' OR Name LIKE '%" + $searchBox.Text + "%'"
    })

# Create the PictureBox
$pictureBox = New-Object System.Windows.Forms.PictureBox
$pictureBox.Location = New-Object System.Drawing.Point(220, 5) 
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

#add right click context menu
$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
$lookUpService = New-Object System.Windows.Forms.ToolStripMenuItem
$lookUpService.Text = 'Look Up Service'
try {
    $lookUpService.Image = [System.Drawing.Image]::FromFile('Assets\LookUp.png')
}
catch {
    Write-Host 'Missing Asset (Lookup Icon)' -ForegroundColor Red
}
$stop = New-Object System.Windows.Forms.ToolStripMenuItem
$stop.Text = 'Stop'
try {
    $stop.Image = [System.Drawing.Image]::FromFile('Assets\Stop.png')
}
catch {
    Write-Host 'Missing Asset (Stop Icon)' -ForegroundColor Red
}
$start = New-Object System.Windows.Forms.ToolStripMenuItem
$start.Text = 'Start'
try {
    $start.Image = [System.Drawing.Image]::FromFile('Assets\Start.png')
}
catch {
    Write-Host 'Missing Asset (Start Icon)' -ForegroundColor Red
}
$disable = New-Object System.Windows.Forms.ToolStripMenuItem
$disable.Text = 'Disable'
try {
    $disable.Image = [System.Drawing.Image]::FromFile('Assets\Disable.png')
}
catch {
    Write-Host 'Missing Asset (Disable Icon)' -ForegroundColor Red
}
$manual = New-Object System.Windows.Forms.ToolStripMenuItem
$manual.Text = 'Manual'
try {
    $manual.Image = [System.Drawing.Image]::FromFile('Assets\Manual.png')
}
catch {
    Write-Host 'Missing Asset (Manual Icon)' -ForegroundColor Red
}
$auto = New-Object System.Windows.Forms.ToolStripMenuItem
$auto.Text = 'Automatic'
try {
    $auto.Image = [System.Drawing.Image]::FromFile('Assets\Auto.png')
}
catch {
    Write-Host 'Missing Asset (Automatic Icon)' -ForegroundColor Red
}
$Global:system = New-Object System.Windows.Forms.ToolStripMenuItem
$system.Text = 'System'
$system.Visible = $false
try {
    $system.Image = [System.Drawing.Image]::FromFile('Assets\System.png')
}
catch {
    Write-Host 'Missing Asset (System Icon)' -ForegroundColor Red
}
$Global:boot = New-Object System.Windows.Forms.ToolStripMenuItem
$boot.Text = 'Boot'
$boot.Visible = $false
try {
    $boot.Image = [System.Drawing.Image]::FromFile('Assets\Boot.png')
}
catch {
    Write-Host 'Missing Asset (Boot Icon)' -ForegroundColor Red
}
$contextMenu.Items.Add($lookUpService) | Out-Null
$contextMenu.Items.Add($stop) | Out-Null
$contextMenu.Items.Add($start) | Out-Null
$contextMenu.Items.Add($disable) | Out-Null
$contextMenu.Items.Add($manual) | Out-Null
$contextMenu.Items.Add($auto) | Out-Null
$contextMenu.Items.Add($boot) | Out-Null
$contextMenu.Items.Add($system) | Out-Null
$dataGridView.ContextMenuStrip = $contextMenu

# Handle the MouseDown event to show the context menu only when a row is selected
$dataGridView.add_MouseDown({
        param($sender, $e)
        if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Right) {
            $hitTestInfo = $dataGridView.HitTest($e.X, $e.Y)
            if ($hitTestInfo.RowIndex -ge 0) {
                #$dataGridView.ClearSelection()
                $dataGridView.Rows[$hitTestInfo.RowIndex].Selected = $true
                $contextMenu.Show($dataGridView, $e.Location)
            }
        }
    })

$lookUpService.Add_Click({
        $selectedServices = Get-SelectedService
        if ($selectedServices.Count -gt 1) {
            foreach ($service in $selectedServices) {
                Start-Process "https://www.google.com/search?q=$service"
            }
        }
        else {
            Start-Process "https://www.google.com/search?q=$selectedServices"
        }
    })

$stop.Add_Click({
        Invoke-Expression -Command "Stop-SelectedService -$Mode"
    })

$start.Add_Click({
        Invoke-Expression -Command "Start-SelectedService -$Mode"
    })

$manual.Add_Click({
        Invoke-Expression -Command "Set-Manual -$Mode"
    })

$auto.Add_Click({
        Invoke-Expression -Command "Set-Automatic -$Mode"
    })

$disable.Add_Click({
        Invoke-Expression -Command "Set-Disabled -$Mode"
    })

$system.Add_Click({
        Set-System
    })

$boot.Add_Click({
        Set-Boot
    })
function addDriversToGrid {
    # Retrieve the driver data
    $driversPlus = Get-DriverPlus

    # Convert the data to a DataTable
    $Global:dataTable = New-Object System.Data.DataTable
    $columnNames = $driversPlus[0].PSObject.Properties.Name 
    foreach ($name in $columnNames) {
        $dataTable.Columns.Add($name) | Out-Null
    }
    foreach ($driver in $driversPlus) {
        $row = $dataTable.NewRow()
        $rows = $driver.PSObject.Properties
        foreach ($r in $rows) {
            $row[$r.Name] = $driver.$($r.Name) 
        }
        $dataTable.Rows.Add($row) | Out-Null
    }

    # Bind the DataTable to the DataGridView
    $dataGridView.DataSource = $dataTable

    #set autosize mode for each column
    $dataGridView.Columns['DisplayName'].AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells
    $dataGridView.Columns['Name'].AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells
    $dataGridView.Columns['Status'].AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells
    $dataGridView.Columns['StartType'].AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells
    $dataGridView.Columns['BinaryPath'].AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells

}


function addServicesToGrid {
    # Retrieve the service data
    $servicesPlus = Get-ServicePlus

    # Convert the data to a DataTable
    $Global:dataTable = New-Object System.Data.DataTable
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

}
addServicesToGrid

# Create the label
$label = New-Object System.Windows.Forms.Label
$label.Text = 'Set Service:'
$label.ForeColor = 'White'
$label.Location = New-Object System.Drawing.Point(470, 5)
$label.AutoSize = $true
$label.Font = New-Object System.Drawing.Font('Segoe UI', 10)
$form.Controls.Add($label)

$refreshGrid = New-Object System.Windows.Forms.Button
$refreshGrid.Text = 'View Services'
$refreshGrid.Size = New-Object System.Drawing.Size(90, 30)
$refreshGrid.Location = New-Object System.Drawing.Point(50, 2)
$refreshGrid.ForeColor = 'White'
$refreshGrid.Add_Click({
        #$dataGridView.Rows.Clear()
        $dataGridView.Columns.Clear()
        $dataGridView.Refresh()
        $boot.Visable = $false
        $system.Visable = $false
        $progressBar1.Visible = $true
        $labelLoading.Visible = $true
        addServicesToGrid
        $labelLoading.Visible = $false
        $progressBar1.Visible = $false
        $dataGridView.Refresh()
        $dataGridView.ClearSelection()
        
    })
$refreshGrid.Add_MouseEnter({
        $refreshGrid.BackColor = [System.Drawing.Color]::FromArgb(64, 64, 64)
    })

$refreshGrid.Add_MouseLeave({
        $refreshGrid.BackColor = [System.Drawing.Color]::FromArgb(43, 43, 42)
    })
$refreshGrid.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($refreshGrid)

$viewDrivers = New-Object System.Windows.Forms.Button
$viewDrivers.Text = 'View Drivers'
$viewDrivers.Size = New-Object System.Drawing.Size(90, 30)
$viewDrivers.Location = New-Object System.Drawing.Point(150, 2)
$viewDrivers.ForeColor = 'White'
$viewDrivers.Add_Click({
        $Global:Mode = 'driver'
        $system.Visible = $true
        $boot.Visable = $true
        #$dataGridView.Rows.Clear()
        $dataGridView.Columns.Clear()
        $dataGridView.Refresh()
        $progressBar1.Visible = $true
        $labelLoading.Visible = $true
        addDriversToGrid
        $labelLoading.Visible = $false
        $progressBar1.Visible = $false
        $dataGridView.Refresh()
        $dataGridView.ClearSelection()
        
    })
$viewDrivers.Add_MouseEnter({
        $viewDrivers.BackColor = [System.Drawing.Color]::FromArgb(64, 64, 64)
    })

$viewDrivers.Add_MouseLeave({
        $viewDrivers.BackColor = [System.Drawing.Color]::FromArgb(43, 43, 42)
    })
$viewDrivers.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($viewDrivers)

$exportMenu = New-Object System.Windows.Forms.Button
$exportMenu.Size = New-Object System.Drawing.Size(40, 30)
$exportMenu.Location = New-Object System.Drawing.Point(3, 2)
$exportMenu.ForeColor = 'White'
$exportMenu.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$exportMenu.FlatAppearance.BorderSize = 0
#$exportMenu.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(62, 62, 64)
#$exportMenu.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(27, 27, 28)
try {
    $image = [System.Drawing.Image]::FromFile('Assets\Menu.png')

    $resizedImage = New-Object System.Drawing.Bitmap $image, 25, 19

    # Set the button image
    $exportMenu.Image = $resizedImage
    $exportMenu.ImageAlign = [System.Drawing.ContentAlignment]::MiddleCenter
}
catch {
    Write-Host 'Missing Asset (Menu Icon)' -ForegroundColor Red
}

$exportMenu.Add_Click({
        $exportContextMenu.Show($exportMenu, 0, $exportMenu.Height)
    })
$form.Controls.Add($exportMenu)

$exportContextMenu = New-Object System.Windows.Forms.ContextMenuStrip
$form.ContextMenuStrip = $exportContextMenu
# Add menu items to the context menu strip
$exportServices = New-Object System.Windows.Forms.ToolStripMenuItem
$exportServices.Text = 'Export Services'
$exportContextMenu.Items.Add($exportServices) | Out-Null
$exportServices.Add_Click({
        $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderBrowser.Description = 'Select a Destination'
        $folderBrowser.RootFolder = [System.Environment+SpecialFolder]::Desktop
        $folderBrowser.ShowNewFolderButton = $true

        # Show the dialog and get the selected folder
        $result = $folderBrowser.ShowDialog()

        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            $selectedFolder = $folderBrowser.SelectedPath
        }
        #build reg file
        $key = 'HKLM\SYSTEM\ControlSet001\Services'
        $header = 'Windows Registry Editor Version 5.00'
        New-Item -Path $selectedFolder -Name 'ServicesBackup.reg' -Force | Out-Null
        Add-Content -Path "$selectedFolder\ServicesBackup.reg" -Value $header -Force
        $regContent = ''
        foreach ($row in $dataGridView.Rows) {
            $name = $row.Cells[1].Value
            if ($null -ne $name) {
                $startValue = Get-ItemPropertyValue -Path "registry::$key\$name" -Name 'Start'
                $regContent += "[$key\$name] `n" + "`"Start`"=dword:0000000$($startValue)`n `n"
            }  
        }
        Add-Content -Path "$selectedFolder\ServicesBackup.reg" -Value $regContent -Force
    })

$exportDrivers = New-Object System.Windows.Forms.ToolStripMenuItem
$exportDrivers.Text = 'Export Drivers'
$exportContextMenu.Items.Add($exportDrivers) | Out-Null
$exportDrivers.Add_Click({
        $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderBrowser.Description = 'Select a Destination'
        $folderBrowser.RootFolder = [System.Environment+SpecialFolder]::Desktop
        $folderBrowser.ShowNewFolderButton = $true

        # Show the dialog and get the selected folder
        $result = $folderBrowser.ShowDialog()

        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            $selectedFolder = $folderBrowser.SelectedPath
        }
        #build reg file
        $key = 'HKLM\SYSTEM\ControlSet001\Services'
        $header = 'Windows Registry Editor Version 5.00'
        New-Item -Path $selectedFolder -Name 'DriversBackup.reg' -Force | Out-Null
        Add-Content -Path "$selectedFolder\DriversBackup.reg" -Value $header -Force
        $regContent = ''
        foreach ($row in $dataGridView.Rows) {
            $name = $row.Cells[1].Value
            if ($null -ne $name) {
                $startValue = Get-ItemPropertyValue -Path "registry::$key\$name" -Name 'Start'
                $regContent += "[$key\$name] `n" + "`"Start`"=dword:0000000$($startValue)`n `n"
            }  
        }
        Add-Content -Path "$selectedFolder\DriversBackup.reg" -Value $regContent -Force
    })

$stopService = New-Object System.Windows.Forms.Button
$stopService.Text = 'Stop Service'
$stopService.Size = New-Object System.Drawing.Size(90, 30)
$stopService.Location = New-Object System.Drawing.Point(900, 2)
$stopService.ForeColor = 'White'
$stopService.Add_Click({
        Stop-SelectedService
    })
$stopService.Add_MouseEnter({
        $stopService.BackColor = [System.Drawing.Color]::FromArgb(64, 64, 64)
    })

$stopService.Add_MouseLeave({
        $stopService.BackColor = [System.Drawing.Color]::FromArgb(43, 43, 42)
    })
$form.Controls.Add($stopService)

$startService = New-Object System.Windows.Forms.Button
$startService.Text = 'Start Service'
$startService.Size = New-Object System.Drawing.Size(90, 30)
$startService.Location = New-Object System.Drawing.Point(990, 2)
$startService.ForeColor = 'White'
$startService.Add_Click({
        Start-SelectedService
    })
$startService.Add_MouseEnter({
        $startService.BackColor = [System.Drawing.Color]::FromArgb(64, 64, 64)
    })

$startService.Add_MouseLeave({
        $startService.BackColor = [System.Drawing.Color]::FromArgb(43, 43, 42)
    })
$form.Controls.Add($startService)

# Create the buttons
$manualButton = New-Object System.Windows.Forms.Button
$manualButton.Text = 'Manual'
$manualButton.Size = New-Object System.Drawing.Size(90, 30)
$manualButton.Location = New-Object System.Drawing.Point(550, 2)
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
$automaticButton.Location = New-Object System.Drawing.Point(650, 2)
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
$disabledButton.Location = New-Object System.Drawing.Point(750, 2)
$disabledButton.ForeColor = 'White'
$disabledButton.Add_MouseEnter({
        $disabledButton.BackColor = [System.Drawing.Color]::FromArgb(64, 64, 64)
    })

$disabledButton.Add_MouseLeave({
        $disabledButton.BackColor = [System.Drawing.Color]::FromArgb(43, 43, 42)
    })
$form.Controls.Add($disabledButton)

$manualButton.Add_Click({
        Invoke-Expression -Command "Set-Manual -$Mode"
    })

$automaticButton.Add_Click({
        Invoke-Expression -Command "Set-Automatic -$Mode"
    })

$disabledButton.Add_Click({
        Invoke-Expression -Command "Set-Disabled -$Mode"
    })


#remove first row selection
$form.Add_Shown({
        $dataGridView.ClearSelection()
    })

# Show the form
[void]$form.ShowDialog()
