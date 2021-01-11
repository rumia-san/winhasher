<#
 # Written by Rumia<rumia-san@outlook.com>
 # Features:
 # 1. Simple File Hasher with GUI
 # 2. Written in Powershell without external modules
 # 3. You could simply drag and drop files to the window!
 #>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$ALGORITHMS = @("MD5", "SHA1", "SHA256")

#Create File ListView
$fileListView = New-Object System.Windows.Forms.ListView

#Set Sytle of ListView
$fileListView.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$fileListView.Dock = [System.Windows.Forms.DockStyle]::Fill
$fileListView.View = [System.Windows.Forms.View]::Details
$fileListView.FullRowSelect = $true
$fileListView.GridLines = $true
$fileListView.MultiSelect = $true
$fileListView.AutoResizeColumns([System.Windows.Forms.ColumnHeaderAutoResizeStyle]::HeaderSize);

#Add column header
$fileListView.Columns.Add("File", -2, [System.Windows.Forms.HorizontalAlignment]::Left)
$ALGORITHMS | ForEach-Object {
    $fileListView.Columns.Add($_, -2, [System.Windows.Forms.HorizontalAlignment]::Left)
}
function hashFile ($filePath) {
    return $ALGORITHMS | ForEach-Object {
        Get-FileHash $filePath -Algorithm $_ | Select-Object -ExpandProperty Hash
    }
}
function addFilesToListView ($filePaths) {
    foreach ($filename in $filePaths)
    {
        if (-Not [System.IO.File]::Exists($filename)) {
            continue
        }
        $hashResults = hashFile($filename)
        $listRow = New-Object -TypeName System.Windows.Forms.ListViewItem -ArgumentList $filename
        $listRow.SubItems.AddRange($hashResults)
        $fileListView.Items.Add($listRow)
    }
    $fileListView.AutoResizeColumns([System.Windows.Forms.ColumnHeaderAutoResizeStyle]::ColumnContent);
}

#Add drag-n-grop event handler
$dragOverHandler = [System.Windows.Forms.DragEventHandler]{
    if ($_.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) # $_ = [System.Windows.Forms.DragEventArgs]
    {
        $_.Effect = 'Copy'
    }
    else
    {
        $_.Effect = 'None'
    }
}
$dragDropHandler = [System.Windows.Forms.DragEventHandler]{
    addFilesToListView($_.Data.GetData([Windows.Forms.DataFormats]::FileDrop))
}
$fileListView.AllowDrop = $true
$fileListView.Add_DragOver($dragOverHandler)
$fileListView.Add_DragDrop($dragDropHandler)

#Add context menu to the file list view
$contextMenuStrip = New-Object System.Windows.Forms.ContextMenuStrip
$clearAllMenuItem = New-Object -TypeName System.Windows.Forms.ToolStripMenuItem `
    -ArgumentList @('Clear All', $null, {$fileListView.Items.Clear()})
$contextMenuStrip.Items.Add($clearAllMenuItem)
$fileListView.ContextMenuStrip = $contextMenuStrip

#Create Add File Button
$addFileButton = New-Object System.Windows.Forms.Button
$addFileButton.Text = "Add File"
$addFileButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$addFileButton.Dock = [System.Windows.Forms.DockStyle]::Top

$addFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$addFileDialog.Multiselect = $true
$addFileDialog.RestoreDirectory = $true
$addFileButton.Add_Click({    
    if ($addFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        addFilesToListView($addFileDialog.FileNames)
    }
})


#Create Main Window
$mainForm = New-Object system.Windows.Forms.Form
$mainForm.ClientSize = '500,230'
$mainForm.text = "File Hasher"
$mainForm.StartPosition = 'CenterScreen'
$mainForm.MinimumSize = '500,230'
$mainForm.Add_KeyDown({
    if($_.KeyCode -eq "Escape") {
        $mainForm.Close()
    }
})
$mainForm.KeyPreview = $true
$mainForm.Controls.Add($fileListView)
$mainForm.Controls.Add($addFileButton)

#Hide the console window
Add-Type -name user32 -member '[DllImport("user32.dll")] public static extern bool ShowWindow(int handle, int state);' -namespace Win32
[Win32.user32]::ShowWindow(([System.Diagnostics.Process]::GetCurrentProcess() | Get-Process).MainWindowHandle, 0)

#If command line arguments is not empty, add to list view
if ($args) {
    addFilesToListView($args)
}

#Show Main window
$mainForm.ShowDialog()
