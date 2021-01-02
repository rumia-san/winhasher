<#
 # Written by Rumia<rumia-san@outlook.com>
 # Features:
 # 1. Simple File Hasher with GUI
 # 2. Written in Powershell without external modules
 # 3. You could simple drag and drop files to the window!
 #>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# TODO: Add File Button
# $addFileButton = New-Object System.Windows.Forms.Button
# $addFileButton.Text = "Add File"
# $addFileButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top `
# -bOR [System.Windows.Forms.AnchorStyles]::Left
# $addFileButton.Dock = [System.Windows.Forms.DockStyle]::Top

$fileListView = New-Object System.Windows.Forms.ListView

#Set Sytle of ListView
$fileListView.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$fileListView.Dock = [System.Windows.Forms.DockStyle]::Fill
$fileListView.View = [System.Windows.Forms.View]::Details
$fileListView.GridLines = $true
$fileListView.AutoResizeColumns([System.Windows.Forms.ColumnHeaderAutoResizeStyle]::HeaderSize);

#Add column header
$algorithms = @("MD5", "SHA1", "SHA256")
$fileListView.Columns.Add("File", -2, [System.Windows.Forms.HorizontalAlignment]::Left)
$algorithms | ForEach-Object {
    $fileListView.Columns.Add($_, -2, [System.Windows.Forms.HorizontalAlignment]::Left)
}
function hashFile ($filePath) {
    return $algorithms | ForEach-Object {
        Get-FileHash $filePath -Algorithm $_ | Select-Object -ExpandProperty Hash
    }
}

#Add drag-n-grop event handler
$listBox_DragOver = [System.Windows.Forms.DragEventHandler]{
    if ($_.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) # $_ = [System.Windows.Forms.DragEventArgs]
    {
        $_.Effect = 'Copy'
    }
    else
    {
        $_.Effect = 'None'
    }
}
$listBox_DragDrop = [System.Windows.Forms.DragEventHandler]{
    foreach ($filename in $_.Data.GetData([Windows.Forms.DataFormats]::FileDrop)) # $_ = [System.Windows.Forms.DragEventArgs]
    {
        $hashResults = hashFile($filename)
        $listRow = New-Object -TypeName System.Windows.Forms.ListViewItem -ArgumentList $filename
        $listRow.SubItems.AddRange($hashResults)
        $fileListView.Items.Add($listRow)
    }
    $fileListView.AutoResizeColumns([System.Windows.Forms.ColumnHeaderAutoResizeStyle]::ColumnContent);
}
$fileListView.AllowDrop = $true
$fileListView.Add_DragOver($listBox_DragOver)
$fileListView.Add_DragDrop($listBox_DragDrop)

#Create Main Window
$mainForm = New-Object system.Windows.Forms.Form
$mainForm.ClientSize = '500,230'
$mainForm.text = "File Hasher"
$mainForm.StartPosition = 'CenterScreen'
$mainForm.MinimumSize = '500,230'
$mainForm.Controls.Add($fileListView)
$mainForm.Controls.Add($addFileButton)
$mainForm.ShowDialog()
