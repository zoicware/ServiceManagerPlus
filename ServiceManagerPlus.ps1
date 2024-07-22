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



Write-Host 'Getting Services Please Wait...'


Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Service Manager Plus'
$form.BackColor = [System.Drawing.Color]::FromArgb(43, 43, 42)
$form.WindowState = 'Maximized'


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
$dataGridView = New-Object System.Windows.Forms.DataGridView
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