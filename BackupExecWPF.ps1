Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Drawing


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
     Write-Host "==================== $($doc.DocumentName) ===================="
    })
    
    # clean up after printing...
    $doc.add_EndPrint({
     if ($null -ne $bitmap ) {
      $bitmap.Dispose()
      $bitmap = $null
     }
     Write-Host "xxxxxxxxxxxxxxxxxxxx $($doc.DocumentName) xxxxxxxxxxxxxxxxxxxx"
    })
    
    # Adjust image size to fit into paper and print image
    $doc.add_PrintPage({
     Write-Host "Printing $imageName..."
    
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
    remove-item $filename -force
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
function get-tapeinfo{
    param( [System.Collections.ArrayList]$tapestring)
    $tapeinfo = @()
    
    [System.Collections.ArrayList]$tapelist=$tapestring
    #write-host $tapelist
        $tapeinfo = invoke-command -argumentlist (,$tapelist) -computername "ocibackup" -scriptblock {param ($tapelist); import-module bemcli;$selected = @();  $media=get-bemedia; $tapelist | foreach-object{ $cur =$_; $media | foreach-object{if($_.name -like "*$cur*"){$selected+=$_ | select-object Name, AllocatedDate, MediaSet; }}}; return $selected}
        #$tapeinfo = invoke-command -argumentlist $curtape -computername "ocibackup" -scriptblock {param ($curtape); import-module bemcli; get-bemedia | where-object { $_.name -like "*$curtape*" } | select Name, AllocatedDate, MediaSet}
        
        #get-bemedia | where-object { $_.name -like "*$curtape*" } | select Name, AllocatedDate, MediaSet
    #}
    #write-host $tapeinfo[1]
    return $tapeinfo
}

[xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    x:Name="Window" Height="580" Width="500" ResizeMode="NoResize">
    <Grid x:Name="Grid">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <Viewbox MaxWidth="400" MaxHeight="500" Name="VBText" StretchDirection="Both" Stretch="Fill">
        <TextBox x:Name = "PathTextBox" AcceptsReturn="True" TextWrapping="Wrap"
            Grid.Column="0"
            Grid.Row="0"
        />
        </Viewbox>
        <Button x:Name = "ValidateButton"
            Content="Validate"
            Grid.Column="1"
            Grid.Row="0"
        />
        <Button x:Name = "RemoveButton"
            Content="Remove"
            Grid.Column="0"
            Grid.Row="1"
        />

    </Grid>
</Window>
"@
$reader = (New-Object System.Xml.XmlNodeReader $xaml)

$window = [Windows.Markup.XamlReader]::Load($reader)
$validateButton = $window.FindName("ValidateButton")
$pathTextBox = $window.FindName("PathTextBox")
$ValidateButton.Add_Click({
    If(-not ($pathTextBox.Text -eq "")){
        #write-host "$($pathTextBox.text)`n Raw text`n`n"

        $parsedtext=(($pathtextbox.text).split()).replace('\s*','')| where-object {($_.length -ge 4)}
        #write-host "$($parsedtext[1] -match '\s')"
        #$parsedtext | foreach-object {write-host "X: $_`n"}
        $tapeinfo =  Get-TapeInfo -tapestring $parsedtext
        # dont parse it!
        #debug line
        #write-host $tapeinfo.size
        
        $tapeinfo | foreach-object { if($_.MediaSet -like "Keep Data for 4 Weeks*" ){$weekno = get-weeknumber -date $_.allocateddate;} else{ $weekno ="ME"}; write-host "Date: $($_.allocateddate)`nWeek #: $($weekno)`nNumber: $($_.Name)`n$($_.Mediaset)"}#printlabel -dateinfo $_.allocateddate -weekno $weekno -tapeno $_.number }
        $tapeinfo | foreach-object { if($_.MediaSet -like "Keep Data for 4 Weeks*" ){$weekno = get-weeknumber -date $_.allocateddate;} else{ $weekno ="ME"}; printlabel -printer "EPSON TM-T88VI Receipt" -dateinfo $_.allocateddate -weekno $weekno -tapeno ($_.name).substring(2) }

    
    }
})
$window.ShowDialog()
