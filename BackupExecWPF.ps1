# todo! 
# add a "suggest tapes to export" button. (req. see what makes a tape "exportable", i.e. full, not appendable, overwritable til whatever, no active)
#   means we need to make sure it isnt active which we don't have a way of doing yet but Im sure BEMCLI will find a way.
# "specify printer, if even by name, have default in place "
# specify credentials to connect to BE,
# specify BE server
# popup pane with arrow
# print or get data button
# wildcards? no we've gone too far.
# wpf inotifyrpoperyt thing : https://smsagent.blog/2017/02/03/powershell-deepdive-wpf-data-binding-and-inotifypropertychanged/

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Drawing
Add-PSSnapin VeeamPSSnapin
# this snapin comes from installing the veeam console.


function Invoke-ImagePrint {
    param([string]$imageName = $(throw "Enter image name to print"),
       [string]$printer = "",
       [bool]$fitImageToPaper = $true)
   
    trap { break; }
   
    # check out Lee Holmes' blog(http://www.leeholmes.com/blog/HowDoIEasilyLoadAssembliesWhenLoadWithPartialNameHasBeenDeprecated.aspx)
    # on how to avoid using deprecated "LoadWithPartialName" function
    # To load assembly containing System.Drawing.Printing.PrintDocument
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
   
    # Bitmap image to use to print image
    $bitmap = $null
   
    $doc = new-object System.Drawing.Printing.PrintDocument
    # if printer name not given, use default printer
    $doc.defaultpagesettings.margins.left=15
    $doc.defaultpagesettings.margins.right=15
    $doc.defaultpagesettings.margins.top=0
    $doc.defaultpagesettings.margins.bottom=0
    if ($printer -ne "") {
     $doc.PrinterSettings.PrinterName = $printer
    }
    
    $doc.DocumentName = [System.IO.Path]::GetFileName($imageName)
   
    $doc.add_BeginPrint({
     Write-Debug "==================== $($doc.DocumentName) ===================="
    })
    
    # clean up after printing...
    $doc.add_EndPrint({
     if ($null -ne $bitmap ) {
      $bitmap.Dispose()
      $bitmap = $null
     }
     Write-Debug "xxxxxxxxxxxxxxxxxxxx $($doc.DocumentName) xxxxxxxxxxxxxxxxxxxx"
    })
    
    # Adjust image size to fit into paper and print image
    $doc.add_PrintPage({
     Write-Debug "Printing $imageName..."
    
     #$g = $_.Graphics
     $pageBounds = $_.MarginBounds
     $img = new-object Drawing.Bitmap($imageName)
     
     $adjustedImageSize = $img.Size
     $ratio = [double] 1;
     
     # Adjust image size to fit on the paper
     if ($fitImageToPaper) {
      $fitWidth = [bool] ($img.Size.Width > $img.Size.Height)
      if (($img.Size.Width -le $_.MarginBounds.Width) -and
       ($img.Size.Height -le $_.MarginBounds.Height)) {
       $adjustedImageSize = new-object System.Drawing.SizeF($img.Size.Width, $img.Size.Height)
      } else {
       if ($fitWidth) {
        $ratio = [double] ($_.MarginBounds.Width / $img.Size.Width);
       } else {
        $ratio = [double] ($_.MarginBounds.Height / $img.Size.Height)
       }
       
       $adjustedImageSize = new-object System.Drawing.SizeF($_.MarginBounds.Width, [float]($img.Size.Height * $ratio))

      }
     }
   
     # calculate destination and source sizes
     $recDest = new-object Drawing.RectangleF($pageBounds.Location, $adjustedImageSize)
     $recSrc = new-object Drawing.RectangleF(0, 0, $img.Width, $img.Height)
     
     # Print to the paper
     $_.Graphics.DrawImage($img, $recDest, $recSrc, [Drawing.GraphicsUnit]"Pixel")
     
     $_.HasMorePages = $false; # nothing else to print
    })
    
    $doc.Print()
    $doc.Dispose()
   }

   function printlabel {
    param( $HeaderText="OSCO", [system.datetime]$dateinfo, $weekno, $tapeno,$printer="EPSON TM-T88VI Receipt" )

    $line1=$headertext
    $dateline=[system.string] ($dateinfo | Get-Date -format d)
    $Weekline="Week # - $($weekno)"
    $tapeline="Tape #$($tapeno)"

    $randname= -join ((65..90) + (97..122) | Get-Random -Count 10 | foreach-object {[char]$_})
    $filename = "$env:temp\$($randname).png" 
    
    $bmp = new-object System.Drawing.Bitmap 400,250 
    $font = new-object System.Drawing.Font Consolas,60 
    $fontsmaller = new-object System.Drawing.Font Consolas,30
    
    $brushBg = [System.Drawing.Brushes]::White 
    $brushFg = [System.Drawing.Brushes]::Black 
    
    $graphics = [System.Drawing.Graphics]::FromImage($bmp) 
    $graphics.FillRectangle($brushBg,0,0,$bmp.Width,$bmp.Height) 
    $graphics.DrawString($line1,$font,$brushFg,10,5) 
    $graphics.DrawString($weekline,$fontsmaller,$brushFg,130,100) 
    $graphics.DrawString($dateline,$fontsmaller,$brushFg,130,140) 
    $graphics.DrawString($tapeline,$fontsmaller,$brushFg,130,180) 
    
    $graphics.Dispose() 
    $bmp.rotateflip("Rotate90FlipNone")

    $bmp.Save($filename) 
    Invoke-ImagePrint -printer $printer -imageName $filename -fitImageToPaper $true
    $bmp.dispose()
    # remove-item $filename -force
    # this (remove-item) doesn't seem to work, so we will just spam TEMP with images and assume it will do cleanup.
}
function get-weeknumber {
    param ([system.datetime]$date)
    $tempdate=new-object System.DateTime -ArgumentList $date.year,$date.Month, 1,12,0,0
    $tmpweekday = [int]($tempdate.dayofweek)
    if($tmpweekday -ge 1) {
        $offset = $tmpweekday -2
     }elseif ($tmpweekday -eq 0){
        $offset = 5
     }
     $weeknumber = [system.math]::floor(($date.day + $offset)/7)+1
     return $weeknumber
}

function Get-VeeamTapeInfo{
    param([System.Collections.ArrayList]$tapestring=@(), $Server, [Pscredential]$Credential)
#use tape user in keepass
if($Credential -eq $null)
{
    Connect-VBRServer -server $Server
# if we don't specify a credential we try to connect as the user.
}else{
# otherwise connect with pscredential object specified.
    Connect-VBRServer -server $Server -credential $credential
}
    [System.Collections.ArrayList]$selected =@()
    [System.Collections.ArrayList]$tapelist = $tapestring
    $alltapes = get-vbrtapemedium
    $tapelist | foreach-object{ $cur =$_; $alltapes | foreach-object{if($_.name -like "*$cur*"){$selected+=$_ | select-object Name, @{N='AllocatedDate'; E={$_.LastWriteTime}}, MediaSet,@{N='Type';E={"Veeam"}};}}}#, ExpirationDate; }}}

#Properties to care about:
#   Name -> barcode basically
#   Barcode -> also barcode
#   ExpirationData -> determine if Weekly, Monthly, Year-end
#   LastWriteTime -> when tape last written ~ Allocation date
#   MediaSet -> basically the easiest way to determine Month-end or nah.
    write-host $selected
Disconnect-VBRServer
    return $selected
}

# END OF FUNCTION ->
function get-tapeinfo{
    param( [System.Collections.ArrayList]$tapestring, $server, [Pscredential]$Credential)
    $tapeinfo = @()
    
    [System.Collections.ArrayList]$tapelist=$tapestring
    if ($credential){
        $tapeinfo = invoke-command -credential $credential -argumentlist (,$tapelist) -computername $server -scriptblock {param ($tapelist); import-module bemcli;$selected = @();  $media=get-bemedia; $tapelist | foreach-object{ $cur =$_; $media | foreach-object{if($_.name -like "*$cur*"){$selected+=$_ | select-object Name, AllocatedDate, MediaSet,@{N='Type';E={"BackupExec"}}; }}}; return $selected}
 
    }else{
        $tapeinfo = invoke-command -argumentlist (,$tapelist) -computername $server -scriptblock {param ($tapelist); import-module bemcli;$selected = @();  $media=get-bemedia; $tapelist | foreach-object{ $cur =$_; $media | foreach-object{if($_.name -like "*$cur*"){$selected+=$_ | select-object Name, AllocatedDate, MediaSet,@{N='Type';E={"BackupExec"}}; }}}; return $selected}
    }
    #write-host $tapelist
               #$tapeinfo = invoke-command -argumentlist $curtape -computername "ocibackup" -scriptblock {param ($curtape); import-module bemcli; get-bemedia | where-object { $_.name -like "*$curtape*" } | select Name, AllocatedDate, MediaSet}
        
        #get-bemedia | where-object { $_.name -like "*$curtape*" } | select Name, AllocatedDate, MediaSet
    #}
    #write-host $tapeinfo[1]
    return $tapeinfo
}
function save-settingsfile{
param([PSCustomObject]$SettingsObject)
#todo! Add some settings validation to make sure we're not just chunking any old pscustomobject to disk.
    $exist=Test-Path "$($env:appdata)\TapeExplorerCL\settings.json"
    if(! (test-path "$($env:appdata)\TapeExplorerCL")) {
        New-Item "$($env:appdata)\TapeExplorerCL" -itemtype directory
    }
    $settingsobject | convertto-json | out-file "$($env:appdata)\TapeExplorerCL\settings.json"
}
function Load-settingsfile{
    $exist=Test-Path "$($env:appdata)\TapeExplorerCL\settings.json"
    if ($exist){
        $settingsobject = get-content "$($env:appdata)\TapeExplorerCL\settings.json" | convertfrom-json
    }else {
        if (! (test-path "$($env:appdata)\TapeExplorerCL")) {
            New-Item "$($env:appdata)\TapeExplorerCL" -itemtype directory
        }
        $settingsobject=[PSCustomObject]@{
            UseVeeam         = $true
            VeeamServer      = "VeeamMgmt.oci.ad"
            UseAltVeeamCreds = $false
            VeeamCredential  = "~\Administrator"

            UseBackupExec    = $true
            BackupExecServer = "ocibackup.oci.ad"
            UseAltBackupExecCreds = $false
            BackupExecCredential = "oci\sm_backup"

            PrinterName = "EPSON TM-T88VI Receipt"
         }
         $settingsobject | convertto-json | out-file "$($env:appdata)\TapeExplorerCL\settings.json"
    }
    return $settingsobject

}

$settingsobject =Load-settingsfile
[xml]$settingsxaml= @"
<Window x:Name="SettingsWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        Title="Settings" Height="320" Width="500" ResizeMode="NoResize" Opacity="1.00" AllowsTransparency="False">
    <Grid Width="500" Height="300">
        <CheckBox x:Name="UseAltVeeamCreds" Content="Alternate Veeam Credentials" HorizontalAlignment="Left" Margin="19,59,0,0" VerticalAlignment="Top" />
        <TextBox x:Name="VeeamUserName" HorizontalAlignment="Left" Height="23" Margin="96,73,0,0" TextWrapping="Wrap" Text="~\Administrator" VerticalAlignment="Top" Width="120"/>
        <CheckBox x:Name="EnableVeeam" Content="Use Veeam Server" HorizontalAlignment="Left" Margin="19,16,0,0" VerticalAlignment="Top" IsChecked="True"/>
        <TextBox x:Name="VeeamServer" HorizontalAlignment="Left" Height="23" Margin="96,32,0,0" TextWrapping="Wrap" Text="VeeamMgmt" VerticalAlignment="Top" Width="120"/>
        <CheckBox x:Name="UseAltBackupExecCreds" Content="Alternate BackupExec Credentials" HorizontalAlignment="Left" Margin="243,59,0,0" VerticalAlignment="Top"/>
        <TextBox x:Name="BackupExecUser" HorizontalAlignment="Left" Height="23" Margin="320,73,0,0" TextWrapping="Wrap" Text="oci\sm_backup" VerticalAlignment="Top" Width="120"/>
        <CheckBox x:Name="EnableBackupExec" Content="Use BackupExec Server" HorizontalAlignment="Left" Margin="243,16,0,0" VerticalAlignment="Top" IsChecked="True"/>
        <TextBox x:Name="BackupExecServer" HorizontalAlignment="Left" Height="23" Margin="320,32,0,0" TextWrapping="Wrap" Text="ocibackup.oci.ad" VerticalAlignment="Top" Width="120"/>
        <PasswordBox x:Name="VeeamPassword" HorizontalAlignment="Left" Margin="96,107,0,0" VerticalAlignment="Top" Width="120"/>
        <PasswordBox x:Name="BackupExecPassword" HorizontalAlignment="Left" Margin="320,105,0,0" VerticalAlignment="Top" Width="120"/>
        <Label x:Name="VeeamUserNameLabel" Content="Username&#xD;&#xA;" HorizontalAlignment="Left" Margin="19,72,0,0" VerticalAlignment="Top" Height="24"/>
        <Label x:Name="VeeamPasswordLabel" Content="Password&#xA;" HorizontalAlignment="Left" Margin="19,102,0,0" VerticalAlignment="Top" Height="26"/>
        <Label x:Name="BackupExecPasswordLabel" Content="Password&#xA;" HorizontalAlignment="Left" Margin="243,102,0,0" VerticalAlignment="Top" Height="26"/>
        <Label x:Name="BackupExecUsernameLabel" Content="Username&#xA;" HorizontalAlignment="Left" Margin="243,72,0,0" VerticalAlignment="Top" Height="26"/>
        <Label x:Name="BackupExecServerNameLabel" Content="Server&#xA;" HorizontalAlignment="Left" Margin="243,29,0,0" VerticalAlignment="Top" Height="26"/>
        <Label x:Name="VeeamServerNameLabel" Content="Server&#xA;" HorizontalAlignment="Left" Margin="19,33,0,0" VerticalAlignment="Top" Height="26"/>
        <Label x:Name="PrinterSelectLabel" Content="Printer&#xD;&#xA;" HorizontalAlignment="Left" Margin="19,176,0,0" VerticalAlignment="Top" Height="27"/>
        <TextBox x:Name="PrinterName" HorizontalAlignment="Left" Height="23" Margin="96,180,0,0" TextWrapping="Wrap" Text="EPSON TM-T88VI Receipt" VerticalAlignment="Top" Width="190"/>
        <Button x:Name="PrinterSelect" Content="Browse" HorizontalAlignment="Left" Margin="291,180,0,0" VerticalAlignment="Top" Width="94"/>
        <Button x:Name="OKSettings" Content="Ok" HorizontalAlignment="Left" Margin="96,251,0,0" VerticalAlignment="Top" Width="75"/>
        <Button x:Name="CancelSettings" Content="Cancel" HorizontalAlignment="Left" Margin="291,251,0,0" VerticalAlignment="Top" Width="75"/>

    </Grid>
</Window>
"@
[xml]$xaml = @"
<Window 
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        Title="TapeExplorer" Height="525" Width="460" ResizeMode="NoResize">
        <Grid x:Name="Grid" Height="500" VerticalAlignment="Top" Margin="0,0,0,-1">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto" MinHeight="340"/>
            <RowDefinition Height="Auto" MinHeight="210"/>
        </Grid.RowDefinitions>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="185"/>
            <ColumnDefinition Width="270" x:Name = "InfoColumn" >
                <ColumnDefinition.Style>
                    <Style TargetType="ColumnDefinition">
                        <Setter Property="Width" Value="*" />
                        <Style.Triggers>
                            <DataTrigger Binding="{Binding IsColumnVisible}" Value="False">
                                <Setter Property="Width" Value="0" />
                            </DataTrigger>
                        </Style.Triggers>
                    </Style>
                </ColumnDefinition.Style>
            </ColumnDefinition>
            <ColumnDefinition Width="40"/>
        </Grid.ColumnDefinitions>
        <TextBox x:Name="PathTextBox" HorizontalAlignment="Left" Height="305" Margin="7,27,0,0" VerticalAlignment="Top" Width="144" Cursor="Arrow" FontSize="24" Text="1234" ScrollViewer.HorizontalScrollBarVisibility="Disabled" ScrollViewer.CanContentScroll="True" VerticalScrollBarVisibility="Auto" AcceptsReturn="True"/>
        <Button x:Name = "ClearButton" 
                Content="Clear Info"
                Grid.Row="1" Margin="10,92,102,94" 
                />
        <Button x:Name = "PrintButton"
                Content="Print"
                Margin="10,18,13,162"
                Grid.Row="1"
        />
        <Button x:Name="InfoButton" 
            Content="Get Info" 
            HorizontalAlignment="Left" 
            Margin="10,55,0,0" 
            VerticalAlignment="Top" 
            Width="162" Height="30" 
            Grid.Row="1"/>
        <Button x:Name="SettingsButton" 
                Content="Settings" 
                HorizontalAlignment="Left" 
                Margin="91,92,0,0" 
                Grid.Row="1" 
                VerticalAlignment="Top" 
                Width="81" Height="24" 
                />
        <Button x:Name="DebugButton" 
                Content="DebugConsole" 
                HorizontalAlignment="Left" 
                Margin="10,120,0,0" 
                Grid.Row="1" 
                VerticalAlignment="Top" 
                Height="20" Width="162"/>
        <Label x:Name="TextBoxLabel" 
               Content="Enter Tape Numbers:" 
               HorizontalAlignment="Left" 
               Margin="10,1,0,0" 
               VerticalAlignment="Top" Width="141" Height="26"/>
        <Button x:Name="ExpandButton" Content="&gt;" 
                HorizontalAlignment="Left" 
                Margin="156,132,0,0" 
                VerticalAlignment="Top" 
                Width="25" Height="60" 
                FontSize="24"/>
        <TextBox x:Name="DebugText"  Background="Transparent" BorderThickness="0" Grid.Column="1" HorizontalAlignment="Left" Margin="10,32,0,0" Grid.Row="1" TextWrapping="Wrap" VerticalAlignment="Top" Height="108" Width="249" ScrollViewer.CanContentScroll="True" ScrollViewer.HorizontalScrollBarVisibility="Visible" ScrollViewer.VerticalScrollBarVisibility="Visible" IsEnabled="False" Visibility="Hidden"></TextBox>
        <Label x:Name="DebugLabel" Content="Debug Console" Grid.Column="1" HorizontalAlignment="Left" Margin="0,6,0,0" VerticalAlignment="Top" Grid.Row="1" IsEnabled="False" Height="26" Width="91" />
        <TextBox x:Name="TapeInfoBox" Grid.Column="1" Background="Transparent" BorderThickness="0" HorizontalAlignment="Left" Margin="10,27,0,0" TextWrapping="Wrap" Text="No Info Loaded Yet." VerticalAlignment="Top" Height="305" Width="241" IsReadOnly="True" IsUndoEnabled="False" />
        <Label Content="Tape Information:" Grid.Column="1" HorizontalAlignment="Left" Margin="40,-2,0,0" VerticalAlignment="Top" FontSize="14" FontWeight="Bold" Height="29" Width="128" />

    </Grid>
</Window>
"@
$settingsreader = (New-Object System.Xml.XmlNodeReader $settingsxaml)

$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$global:creds= @()
$window = [Windows.Markup.XamlReader]::Load($reader)
$settingsbutton = $window.FindName("SettingsButton")
$validateButton = $window.FindName("PrintButton")
$infoButton = $window.findname("InfoButton")
$DebugButton = $window.findname("DebugButton")
$pathTextBox = $window.FindName("PathTextBox")
$Expandbutton = $window.FindName("ExpandButton")
$infopane = $window.FindName("InfoColumn")
$tapeinfobox = $window.findname("TapeInfoBox")
$debugtextbox =$window.findname("DebugText")
$expandbutton.add_click({
    if (-not($infopane.width -eq [system.windows.gridlength]0)) {
        $infopane.width = 0
        $window.width=460-250

    }else {
        $window.width=460
        $infopane.width = 270
    }
})

$ValidateButton.Add_Click({
    If(-not ($pathTextBox.Text -eq "")){
        #write-host "$($pathTextBox.text)`n Raw text`n`n"
        $tapeinfobox.text =""
        $parsedtext = new-object System.collections.generic.list[System.string]
        (($pathtextbox.text).split()).replace('\s*','')| where-object {($_.length -ge 4)}|foreach-object{ $parsedtext.add($_)}
        $debugtextbox.text = $parsedtext + "`n`n" + $debugtextbox.text
        #write-host "$($parsedtext[1] -match '\s')"
        #$parsedtext | foreach-object {write-host "X: $_`n"}
        $tapeinfo =@()
        $settings = Load-settingsfile
        if($settings.UseBackupExec){
            write-host $global:creds.type
            if (($global:creds | where-object {$_.type -eq "BackupExec"})){
                $tapeinfo += Get-TapeInfo -tapestring $parsedtext -credential ($global:creds | where-object {$_.type -eq "BackupExec"}).Credential -server $settings.BackupExecServer
            }else{
                $tapeinfo += Get-TapeInfo -tapestring $parsedtext -server $settings.BackupExecServer
            }
        }
        if($settings.UseVeeam){
            if (($global:creds | where-object {$_.type -eq "Veeam"})){
                $tapeinfo += Get-VeeamTapeInfo -tapestring $parsedtext -credential ($global:creds | where-object {$_.type -eq "Veeam"}).Credential -Server $settings.veeamserver
                write-host "Using Veeam, Credential"
            }else{
                $tapeinfo += Get-VeeamTapeInfo -tapestring $parsedtext -Server $settings.veeamserver
                write-host "Using Veeam, no credential"
            }
        }
        write-host $tapeinfo
        #$tapeinfo =  Get-TapeInfo -tapestring $parsedtext
        # dont parse it!
        #debug line
        #write-host $tapeinfo.size
        
        $tapeinfo | foreach-object { if(($_.MediaSet -like "Keep Data for 4 Weeks*") -or ($_.MediaSet -like "*Weekly media set*") ){$weekno = get-weeknumber -date $_.allocateddate;} else{ $weekno ="ME"};#printlabel -dateinfo $_.allocateddate -weekno $weekno -tapeno $_.number }
        $debugtextbox.text = "Location: $($_.Type)`nDate: $($_.allocateddate)`nWeek #: $($weekno)`nNumber: $($_.Name)`n$($_.Mediaset)`n`n"+$debugtextbox.text;
        $tapeinfobox.text += "Location: $($_.Type)`nDate: $($_.allocateddate)`nWeek #: $($weekno)`nNumber: $($_.Name)`n$($_.Mediaset)`n`n";
        #printlabel -printer "EPSON TM-T88VI Receipt" -header "OSCO-$(($_.Type)[0])" -dateinfo $_.allocateddate -weekno $weekno -tapeno ($_.name).substring(2);
    }
        $tapeinfo | foreach-object { if(($_.MediaSet -like "Keep Data for 4 Weeks*") -or ($_.MediaSet -like "*Weekly media set*") ){$weekno = get-weeknumber -date $_.allocateddate;} else{ $weekno ="ME"}; printlabel -printer "EPSON TM-T88VI Receipt" -header "OSCO-$(($_.Type)[0])" -dateinfo $_.allocateddate -weekno $weekno -tapeno ($_.name).substring(2) }

    
    }
})
$InfoButton.Add_Click({
    If(-not ($pathTextBox.Text -eq "")){
        $tapeinfobox.text =""
        $parsedtext = new-object System.collections.generic.list[System.string]
        (($pathtextbox.text).split()).replace('\s*','')| where-object {($_.length -ge 4)}|foreach-object{ $parsedtext.add($_)}
        $debugtextbox.text = $parsedtext + "`n`n" + $debugtextbox.text
        #write-host "$($parsedtext[1] -match '\s')"
        #$parsedtext | foreach-object {write-host "X: $_`n"}
        $tapeinfo =@()
        $settings = Load-settingsfile
        if($settings.UseBackupExec){
            write-debug $global:creds.type
            if (($global:creds | where-object {$_.type -eq "BackupExec"})){
                $tapeinfo += Get-TapeInfo -tapestring $parsedtext -credential ($global:creds | where-object {$_.type -eq "BackupExec"}).Credential -server $settings.BackupExecServer
            }else{
                $tapeinfo += Get-TapeInfo -tapestring $parsedtext -server $settings.BackupExecServer
            }
        }
        if($settings.UseVeeam){
            if (($global:creds | where-object {$_.type -eq "Veeam"})){
                $tapeinfo += Get-VeeamTapeInfo -tapestring $parsedtext -credential ($global:creds | where-object {$_.type -eq "Veeam"}).Credential -Server $settings.veeamserver
                write-debug "Using Veeam, Credential"
            }else{
                $tapeinfo += Get-VeeamTapeInfo -tapestring $parsedtext -Server $settings.veeamserver
                write-debug "Using Veeam, no credential"
            }
        }
        
        $tapeinfo | foreach-object { if(($_.MediaSet -like "Keep Data for 4 Weeks*") -or ($_.MediaSet -like "*Weekly media set*") ){$weekno = get-weeknumber -date $_.allocateddate;} else{ $weekno ="ME"};#printlabel -dateinfo $_.allocateddate -weekno $weekno -tapeno $_.number }
        $debugtextbox.text = "Location: $($_.Type)`nDate: $($_.allocateddate)`nWeek #: $($weekno)`nNumber: $($_.Name)`n$($_.Mediaset)`n`n"+$debugtextbox.text;
        $tapeinfobox.text += "Location: $($_.Type)`nDate: $($_.allocateddate)`nWeek #: $($weekno)`nNumber: $($_.Name)`n$($_.Mediaset)`n`n";
    
    }
}
})

$Debugbutton.add_Click({
    if($debugtextbox.IsEnabled -eq "True" ){
        $debugtextbox.IsEnabled = "False"
        $debugtextbox.Visibility = "Hidden"

    }else{
        $debugtextbox.IsEnabled = "True"
        $debugTextbox.visibility = "Visible"
    }
})
$SettingsButton.add_Click({
    $settings = Load-settingsfile

    $settingsreader = (New-Object System.Xml.XmlNodeReader $settingsxaml)

    $settingswindow = [Windows.Markup.XamlReader]::Load($settingsreader)

    #username textbox
    $veeamuser = $settingswindow.FindName("VeeamUserName")
    #servername textbox
    $veeamserver = $settingswindow.FindName("VeeamServer")
    #boolean flag
    $UseAltVeeamCreds = $settingswindow.FindName("UseAltVeeamCreds")
    #boolean flag for enabling veeam
    $enableveeam = $settingswindow.FindName("EnableVeeam")
    #password box need to be handled differently.
    $veeampwd = $settingswindow.FindName("VeeamPassword")

    $UseAltBackupExecCreds = $settingswindow.FindName("UseAltBackupExecCreds")
    $backupexecserver = $settingswindow.findname("BackupExecServer")
    $BackupExecUser = $settingswindow.FindName("BackupExecUser")
    $enablebackupexec = $settingswindow.FindName("EnableBackupExec")
    $backupexecpwd = $settingswindow.FindName("BackupExecPassword")
    
    $Printerselect = $settingswindow.FindName("PrinterSelect")
    $PrinterField = $settingswindow.FindName("PrinterName")
    

    $Okbutton = $settingswindow.FindName("OKSettings")
    $CancelButton = $settingswindow.FindName("CancelSettings")

    $enableveeam.ischecked = $settings.UseVeeam 
    $enablebackupexec.ischecked = $settings.UseBackupExec
    $UseAltBackupExecCreds.ischecked = $settings.UseAltBackupExecCreds
    $UseAltVeeamCreds.IsChecked = $settings.usealtveeamcreds

    $printerfield.text = $settings.printername
    $veeamuser.text = $settings.VeeamCredential 
    $backupexecuser.text = $settings.backupexeccredential
    $veeamserver.text = $settings.VeeamServer
    $backupexecserver.text = $settings.backupexecserver
    
    $okbutton.add_click({
        $settings.UseVeeam = $enableveeam.ischecked
        $settings.UseBackupExec = $enablebackupexec.ischecked
        $settings.UseAltBackupExecCreds = $UseAltBackupExecCreds.ischecked
        $settings.usealtveeamcreds = $UseAltVeeamCreds.IsChecked

        $settings.printername = $printerfield.text
        $settings.veeamcredential = $veeamuser.text
        $settings.backupexeccredential = $backupexecuser.text
        $settings.veeamserver = $veeamserver.text
        $settings.backupexecserver = $backupexecserver.text

        $credentiallist=@()
        save-settingsfile -SettingsObject $settings

        if ( $settings.useveeam -and $settings.usealtveeamcreds){
            $veeamobj=[PSCustomObject]@{
                Type         = "Veeam"
                Credential   = New-Object System.Management.Automation.PSCredential ($settings.veeamcredential, $veeampwd.securepassword) 
             }
        }
        if($settings.usebackupexec -and $settings.UseAltBackupExecCreds){
            $backupexecobj=[PSCustomObject]@{
                Type         = "BackupExec"
                Credential   = New-Object System.Management.Automation.PSCredential ($settings.backupexeccredential, $backupexecpwd.securepassword) 
             }
        }
        if($veeamobj){
            $credentiallist+= $veeamobj
        }
        if($backupexecobj){
            $credentiallist+=$backupexecobj
        }
        $settingswindow.close()
        $global:creds= $credentiallist
    })
    $CancelButton.add_click({
        $settingswindow.close()
    })
    $printerselect.add_click({
        
    })

    $settingswindow.ShowDialog()

})

$DataContext = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
$infopanestatus = [int32]1
$datacontext.add($infopanestatus)
$expandbutton.datacontext =$datacontext
$Binding = New-Object System.Windows.Data.Binding
$Binding.Path = "[0]"
$Binding.Mode = [System.Windows.Data.BindingMode]::OneWay
#[void][System.Windows.Data.BindingOperations]::SetBinding($expandbutton,[System.Windows.Controls.TextBox]::v, $Binding)

$window.ShowDialog()
