Add-PSSnapin VeeamPSSnapin
# this snapin comes from installing the veeam console.

function Get-VeeamTapeInfo{
    param([System.Collections.ArrayList]$tapelist=@(), $Server, [Pscredential]$Credential)
#use tape user in keepass
if($Credential -eq $null)
{
# if we don't specify a credential we try to connect as the user.
}else{
# otherwise connect with pscredential object specified.
    Connect-VBRServer -server $Server -credential $credential
}
    [System.Collections.ArrayList]$selected =@()

    $alltapes = get-vbrtapemedium
    $tapelist | foreach-object{ $cur =$_; $alltapes | foreach-object{if($_.name -like "*$cur*"){$selected+=$_ | select-object Name, @{N='AllocatedDate'; E={$_.LastWriteTime}}, MediaSet, ExpirationDate; }}}

#Properties to care about:
#   Name -> barcode basically
#   Barcode -> also barcode
#   ExpirationData -> determine if Weekly, Monthly, Year-end
#   LastWriteTime -> when tape last written ~ Allocation date
#   MediaSet -> basically the easiest way to determine Month-end or nah.

Disconnect-VBRServer
    return $selected
}

# END OF FUNCTION ->
# Anything below here is for testing :)

$veeamcreds=get-credential
$veeamservername="localhost"
[System.Collections.ArrayList]$testtapes = @("1767","1758","1765","1771","1774","1773","1772","1740","1744","1757","1756","1764","1769","1759")
Get-VeeamTapeInfo -server $veeamservername -tapelist $testtapes -Credential $veeamcreds
