cd "P:\ANG_System_Files"
#cd "D:\MyFirstProject\src\AngSystemFiles"

function Load-Dll
{
    param(
        [string]$assembly
    )
    Write-Host "Loading $assembly"

    $driver = $assembly
    $fileStream = ([System.IO.FileInfo] (Get-Item $driver)).OpenRead();
    $assemblyBytes = new-object byte[] $fileStream.Length
    $fileStream.Read($assemblyBytes, 0, $fileStream.Length) | Out-Null;
    $fileStream.Close();
    $assemblyLoaded = [System.Reflection.Assembly]::Load($assemblyBytes);
}

function DuplicateChecker
{
    $outputFilePT = "Label_PT_DuplicateChecker.txt"

    $pos = foreach ($row in $pt.Rows)
    {
        $row.Cells[0].Value
    }

    $pos = $pos | where {$_ -ne $null -and ([string]$_).Split("-", [StringSplitOptions]::RemoveEmptyEntries).Count -gt 1} | sort
    $uniquePos = $pos | select –unique

    Compare -ReferenceObject $uniquePos -DifferenceObject $pos | Out-File $outputFilePT

    If ((Get-Content $outputFilePT) -ne $Null) 
    {
        Start $outputFilePT
        Sleep -Seconds 2
        Stop-Process -Id $pid
    }
}

function Get-ComparisonObjects
{
    param([Smartsheet.Api.Models.Sheet]$sheet)

    Write-Host "Getting Sheet $($sheet.Name) Comparison Objects"

    $data = $sheet.Rows | foreach {
        $checkVal = $false
        
        [pscustomobject]@{
            Attachments = $_.Attachments;
            RowId = $_.Id;
            RowNumber = $_.RowNumber;
            PoCol = $_.Cells[0].ColumnId;
            Po = $_.Cells[0].Value;
            JobsCol = $_.Cells[1].ColumnId;
            Jobs = $_.Cells[1].Value;
            DescCol = $_.Cells[2].ColumnId;
            Desc = $_.Cells[2].Value;
            AssignCol = $_.Cells[5].ColumnId;
            Assign = $_.Cells[5].Value;
            ReceivedCol = $_.Cells[9].ColumnId;
            Received = $_.Cells[9].Value;
            PrintCol = $_.Cells[17].ColumnId;
            Print = $_.Cells[17].Value;
            SKUCol = $_.Cells[18].ColumnId;
            SKU = $_.Cells[18].Value;

        }                                                  
    } | where {![string]::IsNullOrWhiteSpace($_.Po)} 

    Write-Host "$($data.Count) Returned"      
    return $data                                           
}   

Load-Dll ".\smartsheet-csharp-sdk.dll"                     
Load-Dll ".\RestSharp.dll"
Load-Dll ".\Newtonsoft.Json.dll"
Load-Dll ".\NLog.dll"
Load-Dll ".\QRCoder.dll"
Load-Dll ".\UnityEngine.dll"

$token = ""
$smartsheet = [Smartsheet.Api.SmartSheetBuilder]::new()
$builder = $smartsheet.SetAccessToken($token)
$client = $builder.Build()
$includes =  @([Smartsheet.Api.Models.SheetLevelInclusion]::ATTACHMENTS)
$includes = [System.Collections.Generic.List[Smartsheet.Api.Models.SheetLevelInclusion]]$includes
$ptId = "5779331080316804"
$pt  = $client.SheetResources.GetSheet($ptId, $includes, $null, $null, $null, $null, $null, $null);
$poLabelCol = $pt.Columns | where {$_.Title -eq ("PO/WO #")}
$jobsLabelCol = $pt.Columns | where {$_.Title -eq ("Job")}
$descLabelCol = $pt.Columns | where {$_.Title -eq ("Description")}
$assignLabelCol = $pt.Columns | where {$_.Title -eq ("Assigned To")}
$printLabelCol = $pt.Columns | where {$_.Title -eq ("Print Label")}
$SkuNumCol     = $pt.Columns | where {$_.Title -eq ("SKU")}

DuplicateChecker

$ptCOs  = Get-ComparisonObjects $pt

foreach ($ptCO in $ptCOs)
{
    if(![string]::IsNullOrWhiteSpace($ptCO.Print))
    {
        if ($($ptCO.Print) -gt 1)
        {
            $count = $($ptCO.Print) + 1
            
            for ($i = 1; $i -le ($ptCO.Print); $i++)
            {
                $newCount = $count -1

                $subPOvalue = "$($ptCO.Po) Piece #$newCount"
                
                $poCell = [Smartsheet.Api.Models.Cell]::new()
                $poCell.ColumnId     = $poLabelCol.Id
                $poCell.Value        = $subPOvalue
                
                $jobsCell = [Smartsheet.Api.Models.Cell]::new()
                $jobsCell.ColumnId   = $jobsLabelCol.Id
                $JobsCell.Value      =  $($ptCO.Jobs)
                
                $descCell = [Smartsheet.Api.Models.Cell]::new()
                $descCell.ColumnId   = $descLabelCol.Id
                $descCell.Value      =  $($ptCO.Desc)
                
                $AssignCell = [Smartsheet.Api.Models.Cell]::new()
                $AssignCell.ColumnId = $assignLabelCol.Id
                $AssignCell.Value    =  $($ptCO.Assign)
                
                $row = [Smartsheet.Api.Models.Row]::new()
                $row.ToBottom = $true
                $row.parentId = $($ptCO.RowId)
                $row.Cells = [Smartsheet.Api.Models.Cell[]]@($poCell,$jobsCell,$descCell,$AssignCell) 
                
                $newRow = $client.SheetResources.RowResources.AddRows($ptId, [Smartsheet.Api.Models.Row[]]@($row))

                $barcodePath = "P:\ANG_System_Files\commonFormsUsedInScripts\Todds Labels\$subPOvalue.png"

                $qrinput = "$subPOvalue`n$($ptCO.Desc)`n$($ptCO.Assign)`nAll New Glass"
                $ecclevel = [QRCoder.QRCodeGenerator+ECCLevel]::Q
                $qrgenerator = [QRCoder.QRCodeGenerator]::new()
                $qrcodedata = $qrgenerator.CreateQrCode($qrinput,$ecclevel) 
                $qrcode = [QRCoder.QRCode]::new($qrcodedata)
                $bitmap = $qrcode.GetGraphic(40)
                $bitmap.Save($barcodePath) 

                $xl = New-Object -ComObject Excel.Application -Property @{
                 Visible = $false
                 DisplayAlerts = $false
                }
                
                $wb = $xl.WorkBooks.Add()
                
                $sh = $wb.Sheets.Item('Sheet1')

                # Excel Constants
                # MsoTriState
                Set-Variable msoFalse 0 -Option Constant -ErrorAction SilentlyContinue
                Set-Variable msoTrue 1 -Option Constant -ErrorAction SilentlyContinue
                
                # own Constants
                # cell width and height in points
                Set-Variable cellWidth 48 -Option Constant -ErrorAction SilentlyContinue
                Set-Variable cellHeight 15 -Option Constant -ErrorAction SilentlyContinue
                
                # arguments to insert the image through the Shapes.AddPicture Method
                
                $imgPath = $barcodePath
                $LinkToFile = $msoFalse
                $SaveWithDocument = $msoTrue
                $Left = $cellWidth * 0
                $Top = $cellHeight * 6
                $Width = $cellWidth * 2
                $Height = $cellHeight * 7
                
                # add image to the Sheet
                $img = $sh.Shapes.AddPicture($imgPath, $LinkToFile, $SaveWithDocument,
                $Left, $Top, $Width, $Height)
                $sh.Range("A1:B5").Font.Size = 18
                $sh.Range("A5").ColumnWidth = 19
                $sh.Cells.Item(2, 1)  = if ($subPOvalue -ne $null){$subPOvalue} else {[string]::Empty}
                $sh.Cells.Item(3, 1)  = if ($ptCO.Jobs -ne $null){$ptCO.Jobs} else {[string]::Empty}
                $sh.Cells.Item(4, 1)  = if ($ptCO.Assign -ne $null){$ptCO.Assign} else {[string]::Empty}
                $sh.Cells.Item(5, 1)  = if ($($ptCO.Received) -ne $null){$($ptCO.Received)} else {[string]::Empty}
                
                (New-Object -ComObject WScript.Network).SetDefaultPrinter(‘Zebra ZP 500 (ZPL)’)
                
                $sh.PrintOut(1,1,1)
                
                $wb.Close($false)
                $xl.Quit()

                (New-Object -ComObject WScript.Network).SetDefaultPrinter(‘HP LaserJet Pro M402-M403 PCL 6’)

                $SkuCell = [Smartsheet.Api.Models.Cell]::new()
                $SkuCell.ColumnId = $SkuNumCol.Id
                $SkuCell.Value    = $qrinput

                $row = [Smartsheet.Api.Models.Row]::new()
                $row.Id = $newRow.Id
                $row.Cells = [Smartsheet.Api.Models.Cell[]]@($SkuCell)

                $updateRow = $client.SheetResources.RowResources.UpdateRows($ptId, [Smartsheet.Api.Models.Row[]]@($row))

                $count = $newCount
            }
        }

        if ($($ptCO.Print) -eq 1)
        {
            $barcodePath = "P:\ANG_System_Files\commonFormsUsedInScripts\Todds Labels\$($ptCO.Po).png"

            $qrinput = "$($ptCO.Po)`n$($ptCO.Desc)`n$($ptCO.Assign)`nAll New Glass"
            $ecclevel = [QRCoder.QRCodeGenerator+ECCLevel]::Q
            $qrgenerator = [QRCoder.QRCodeGenerator]::new()
            $qrcodedata = $qrgenerator.CreateQrCode($qrinput,$ecclevel) 
            $qrcode = [QRCoder.QRCode]::new($qrcodedata)
            $bitmap = $qrcode.GetGraphic(40)
            $bitmap.Save($barcodePath) 

            $xl = New-Object -ComObject Excel.Application -Property @{
             Visible = $false
             DisplayAlerts = $false
            }
            
            $wb = $xl.WorkBooks.Add()
            
            $sh = $wb.Sheets.Item('Sheet1')

            # Excel Constants
            # MsoTriState
            Set-Variable msoFalse 0 -Option Constant -ErrorAction SilentlyContinue
            Set-Variable msoTrue 1 -Option Constant -ErrorAction SilentlyContinue
            
            # own Constants
            # cell width and height in points
            Set-Variable cellWidth 48 -Option Constant -ErrorAction SilentlyContinue
            Set-Variable cellHeight 15 -Option Constant -ErrorAction SilentlyContinue
            
            # arguments to insert the image through the Shapes.AddPicture Method
            
            $imgPath = $barcodePath
            $LinkToFile = $msoFalse
            $SaveWithDocument = $msoTrue
            $Left = $cellWidth * 0
            $Top = $cellHeight * 6
            $Width = $cellWidth * 2
            $Height = $cellHeight * 7
            
            # add image to the Sheet
            $img = $sh.Shapes.AddPicture($imgPath, $LinkToFile, $SaveWithDocument,
            $Left, $Top, $Width, $Height)
            $sh.Range("A1:B5").Font.Size = 18
            $sh.Range("A5").ColumnWidth = 19
            $sh.Cells.Item(2, 1)  = if ($($ptCO.Po) -ne $null){$($ptCO.Po)} else {[string]::Empty}
            $sh.Cells.Item(3, 1)  = if ($ptCO.Jobs -ne $null){$ptCO.Jobs} else {[string]::Empty}
            $sh.Cells.Item(4, 1)  = if ($ptCO.Assign -ne $null){$ptCO.Assign} else {[string]::Empty}
            $sh.Cells.Item(5, 1)  = if ($($ptCO.Received) -ne $null){$($ptCO.Received)} else {[string]::Empty}
            
            (New-Object -ComObject WScript.Network).SetDefaultPrinter(‘Zebra ZP 500 (ZPL)’)
            
            $sh.PrintOut(1,1,1)
            
            $wb.Close($false)
            $xl.Quit()

            (New-Object -ComObject WScript.Network).SetDefaultPrinter(‘HP LaserJet Pro M402-M403 PCL 6’)

            $SkuCell = [Smartsheet.Api.Models.Cell]::new()
            $SkuCell.ColumnId = $SkuNumCol.Id
            $SkuCell.Value    = $qrinput

            $row = [Smartsheet.Api.Models.Row]::new()
            $row.Id = $ptCO.RowId
            $row.Cells = [Smartsheet.Api.Models.Cell[]]@($SkuCell)

            $updateRow = $client.SheetResources.RowResources.UpdateRows($ptId, [Smartsheet.Api.Models.Row[]]@($row))
        }

    $PrintCell = [Smartsheet.Api.Models.Cell]::new()
    $PrintCell.ColumnId = $printLabelCol.Id
    $PrintCell.Value    =  ""

    $row = [Smartsheet.Api.Models.Row]::new()
    $row.Id = $ptCO.RowId
    $row.Cells = [Smartsheet.Api.Models.Cell[]]@($PrintCell)

    $updateRow = $client.SheetResources.RowResources.UpdateRows($ptId, [Smartsheet.Api.Models.Row[]]@($row))

    }
}

# (Copy 1)
